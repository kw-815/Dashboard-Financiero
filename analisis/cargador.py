#!/usr/bin/env python3
"""
analisis/cargador.py
Carga los datos financieros (SharePoint o local) reutilizando las funciones
ya validadas de update_dashboard.py, y los expone como pandas.DataFrame para
análisis ad hoc o reportes. No modifica update_dashboard.py ni escribe nada
en index.html / git.
"""

import sys
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).parent.parent


def _importar_update_dashboard():
    """update_dashboard.py ejecuta parser.parse_args() a nivel de módulo,
    así que hay que neutralizar sys.argv mientras se importa para no chocar
    con los argumentos de quien lo esté importando."""
    sys.path.insert(0, str(BASE_DIR))
    argv_original = sys.argv
    sys.argv = ['update_dashboard.py']
    try:
        import update_dashboard as ud
    finally:
        sys.argv = argv_original
    return ud


def _completar_xls_faltantes(ud, contenidos):
    """cargar_local() en update_dashboard.py solo tiene fallback .xls<->.xlsx
    para BANCOS y FACTURAS. RESULTADOS y FACTURACION no lo tienen, así que si
    contabilidad dejó el .xls viejo en datos/, cargar_local() lo salta en
    silencio. abrir_workbook() SÍ sabe convertir .xls->.xlsx en memoria, solo
    falta que el .xls llegue al dict `contenidos`. Se completa acá, sin tocar
    update_dashboard.py."""
    for nombre_xlsx in ('RESULTADOS.xlsx', 'FACTURACION.xlsx'):
        if nombre_xlsx in contenidos:
            continue
        nombre_xls = nombre_xlsx[:-1]  # .xlsx -> .xls
        path_xls = ud.DATOS_DIR / nombre_xls
        if path_xls.exists():
            contenidos[nombre_xls] = path_xls.read_bytes()
            print(f"    ✓ {nombre_xls} (local, fallback agregado por analisis/cargador.py)")
    return contenidos


def _cargar_pyg_historico(ud):
    """Busca años anteriores en subcarpetas datos/<AAAA>/ (ej. datos/2025/
    RESULTADOS 2025.xls) y los procesa con procesar_pyg(), igual que el año
    en curso. Genérico: sirve para cualquier año que se agregue así, no solo
    2025. abrir_workbook() ya sabe convertir .xls->.xlsx en memoria, así que
    basta con ponerlo en el dict bajo la clave 'RESULTADOS.xls'."""
    pyg_historico = {}
    if not ud.DATOS_DIR.exists():
        return pyg_historico
    for carpeta in sorted(ud.DATOS_DIR.iterdir()):
        if not carpeta.is_dir() or not carpeta.name.isdigit():
            continue
        candidatos = list(carpeta.glob('RESULTADOS*.xls*'))
        if not candidatos:
            continue
        archivo = candidatos[0]
        print(f"  Leyendo histórico {archivo.relative_to(ud.BASE_DIR)} ...")
        clave = 'RESULTADOS.xlsx' if archivo.suffix == '.xlsx' else 'RESULTADOS.xls'
        contenidos_hist = {clave: archivo.read_bytes()}
        resultado = ud.procesar_pyg(contenidos_hist)
        for anio_encontrado, filas in resultado.items():
            pyg_historico.setdefault(anio_encontrado, filas)
    return pyg_historico


def cargar_datos(modo='local', anio=2026):
    """
    Retorna un dict[str, pandas.DataFrame] con todos los datos financieros
    procesados, listos para análisis. No tiene efectos secundarios: no
    escribe archivos, no toca git.

    modo: 'local' (default, usa datos/) o 'sharepoint' (usa Graph API).
    anio: año en curso a extraer de Facturación (por defecto el año en curso).
          PyG en cambio devuelve TODOS los años disponibles (año en curso +
          históricos en datos/<AAAA>/), con columna 'anio' para filtrar.
    """
    ud = _importar_update_dashboard()

    contenidos = ud.cargar_sharepoint() if modo == 'sharepoint' else ud.cargar_local()
    if not contenidos:
        raise RuntimeError(f"No se encontraron archivos fuente (modo={modo}).")
    if modo == 'local':
        contenidos = _completar_xls_faltantes(ud, contenidos)

    pyg_raw = ud.procesar_pyg(contenidos)
    fac_raw = ud.procesar_facturacion(contenidos)
    cxc_raw = ud.procesar_cxc(contenidos)
    saldo_data, monthly, mexico = ud.procesar_bancos(contenidos)
    asignaciones_raw = ud.leer_asignaciones(contenidos)

    fac_sin_emitir = fac_raw.pop('_sin_emitir', [])

    if modo == 'local':
        for anio_hist, filas in _cargar_pyg_historico(ud).items():
            pyg_raw.setdefault(anio_hist, filas)

    asignaciones_filas = [
        {'cliente': cliente, 'categoria': categoria, 'pct': pct}
        for cliente, splits in asignaciones_raw.items()
        for categoria, pct in splits.items()
    ]

    pyg_todos_anios = pd.concat(
        [pd.DataFrame(filas).assign(anio=a) for a, filas in sorted(pyg_raw.items())],
        ignore_index=True
    ) if pyg_raw else pd.DataFrame()

    return {
        'pyg': pyg_todos_anios,
        'facturacion': pd.DataFrame(fac_raw.get(anio, [])),
        'facturacion_sin_emitir': pd.DataFrame(fac_sin_emitir),
        'cxc': pd.DataFrame(cxc_raw),
        'bancos_saldo': pd.DataFrame(saldo_data),
        'bancos_monthly': pd.DataFrame(monthly),
        'bancos_mexico': pd.DataFrame(mexico),
        'asignaciones': pd.DataFrame(asignaciones_filas),
    }


if __name__ == '__main__':
    datos = cargar_datos()
    for nombre, df in datos.items():
        print(f"{nombre}: {df.shape[0]} filas, columnas = {list(df.columns)}")
