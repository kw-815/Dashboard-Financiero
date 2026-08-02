#!/usr/bin/env python3
"""
analisis/regenerar_pyg_2025.py

Herramienta de UNA SOLA VEZ (no forma parte del pipeline regular de
update_dashboard.py) para regenerar PYG_ALL_RAW[2025] en index.html a
partir de datos/2025/RESULTADOS 2025.xls, con los fixes de clasificación
ya aplicados (gasto_grupo SERVICOS/INTERNET, CC_MAP A4/01->A1).

CONTEXTO — por qué existe este script (ver CLAUDE.md, sección "Reconciliación
2025" para el detalle completo):

  2026-08-02: el usuario reportó que la cifra de "resultado antes de
  impuestos" 2025 publicada en el dashboard ($375,091.44, cuadrada
  manualmente por contabilidad con mucho esfuerzo en su momento) no
  coincidía con lo que el script recalculaba desde datos/2025/RESULTADOS
  2025.xls ($381,424.19 antes de los fixes de este script).

  Investigación: 15 movimientos de nómina de septiembre 2025 (asiento 124,
  "Rol de pagos Período: 202509") estaban etiquetados con Centro de Costo
  inválido: CC='A4' (9 filas, $4,562.02) y CC='01' (6 filas, $1,640.76).
  resolver_cc() los descartaba en silencio (ningún error, ninguna alerta —
  antes de que existiera el contador de "CC no reconocido" agregado en la
  auditoría de esta misma sesión).

  - CC='A4' resultó ser un CC REAL y activo: aparece también en marzo,
    abril y junio de 2026 (siempre en la cuenta "Bono Incentivo",
    5.2.01.005), nunca dado de alta en VALID_CC/CC_GRUPO. Confirmado con
    el usuario: pertenece a Comunicación (A1).
  - CC='01' es un typo puntual: aparece UNA sola vez en toda la data
    2025-2026, junto a un paquete completo de nómina que en todos los
    demás meses va bajo CC='A1'. Confirmado con el usuario: era A1.

  Ambos se agregaron a CC_MAP en update_dashboard.py (fix permanente,
  beneficia también a 2026 en adelante). Con ese fix, el resultado 2025
  recalculado da $375,221.41 — a $129.97 (0.03%) de la cifra cuadrada por
  contabilidad. Diferencia no explicada, probablemente redondeo de la
  conciliación manual original; no se investigó más a fondo por ser
  inmaterial.

  El usuario decidió explícitamente regenerar el 2025 publicado con esta
  cifra corregida (en vez de dejar la vieja $375,091.44), porque está más
  cerca de la realidad contable y de paso corrige la misma clasificación
  SERVICOS/INTERNET que ya se corrigió para 2026.

QUÉ HACE este script:
  1. Lee datos/2025/RESULTADOS 2025.xls
  2. Corre procesar_pyg() (de update_dashboard.py, sin modificarlo más
     allá del fix de CC_MAP ya aplicado)
  3. Reemplaza SOLO el bloque `2025:[...]` dentro de
     `const PYG_ALL_RAW = {...}` en index.html — no toca 2026 ni 2024,
     no toca FAC_ALL_RAW (no hay archivo fuente de Facturación 2025 en
     datos/2025/, ese año sigue con Facturación tal cual estaba)
  4. Imprime un resumen antes/después para verificación manual

NO se ejecuta como parte de `python3 update_dashboard.py` — es intencional
que este script no vuelva a correr solo. 2025 es un año cerrado; si hace
falta tocarlo de nuevo, correr esto a mano y revisar el diff con cuidado.

Uso:
    python3 analisis/regenerar_pyg_2025.py           # dry-run, solo imprime
    python3 analisis/regenerar_pyg_2025.py --escribir # aplica el cambio a index.html
"""

import argparse
import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
ARCHIVO_2025 = BASE_DIR / 'datos' / '2025' / 'RESULTADOS 2025.xls'
INDEX_PATH = BASE_DIR / 'index.html'


def _importar_update_dashboard():
    sys.path.insert(0, str(BASE_DIR))
    argv_original = sys.argv
    sys.argv = ['update_dashboard.py']
    try:
        import update_dashboard as ud
    finally:
        sys.argv = argv_original
    return ud


def calcular_pyg_2025(ud):
    if not ARCHIVO_2025.exists():
        raise FileNotFoundError(f"No se encontró: {ARCHIVO_2025}")
    contenidos = {'RESULTADOS.xls': ARCHIVO_2025.read_bytes()}
    resultado = ud.procesar_pyg(contenidos)
    filas_2025 = resultado.get(2025, [])
    if not filas_2025:
        raise RuntimeError(
            f"procesar_pyg() no devolvió filas para 2025 — revisar "
            f"{ARCHIVO_2025} y el año de las fechas en el archivo."
        )
    return filas_2025


def reemplazar_2025_en_html(html, filas_2025):
    """Reemplaza SOLO el array 2025 dentro de `const PYG_ALL_RAW = {...}`,
    dejando 2026 y 2024 intactos. Localiza el bloque igual que
    _extract_2025_from_html() en update_dashboard.py, pero para escribir
    en vez de extraer."""
    start = html.find('const PYG_ALL_RAW')
    if start == -1:
        raise RuntimeError("No se encontró 'const PYG_ALL_RAW' en index.html")
    p25 = html.find('  2025:', start)
    if p25 == -1:
        raise RuntimeError("No se encontró el bloque '2025:' dentro de PYG_ALL_RAW")
    arr_start = html.index('[', p25)
    depth = 0
    arr_end = None
    for i in range(arr_start, len(html)):
        if html[i] == '[':
            depth += 1
        elif html[i] == ']':
            depth -= 1
            if depth == 0:
                arr_end = i + 1
                break
    if arr_end is None:
        raise RuntimeError("No se pudo determinar el final del array 2025 (llaves desbalanceadas)")

    nuevo_array = json.dumps(filas_2025, ensure_ascii=False)
    return html[:arr_start] + nuevo_array + html[arr_end:]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--escribir', action='store_true',
                         help='Aplicar el cambio a index.html (sin esto, solo imprime el resumen)')
    args = parser.parse_args()

    ud = _importar_update_dashboard()
    filas_2025 = calcular_pyg_2025(ud)

    ing = round(sum(r['ing'] for r in filas_2025), 2)
    costo = round(sum(r['costo'] for r in filas_2025), 2)
    gastos = round(sum(r['gastos'] for r in filas_2025), 2)
    res = round(sum(r['res'] for r in filas_2025), 2)

    print(f"PYG 2025 recalculado desde {ARCHIVO_2025.name} ({len(filas_2025)} registros):")
    print(f"  Ingresos:  ${ing:>14,.2f}")
    print(f"  Costos:    ${costo:>14,.2f}")
    print(f"  Gastos:    ${gastos:>14,.2f}")
    print(f"  Resultado: ${res:>14,.2f}   (cifra contabilidad: $375,091.44, diff: ${res-375091.44:,.2f})")

    if not args.escribir:
        print("\n(dry-run — nada se escribió. Correr con --escribir para aplicar el cambio a index.html)")
        return

    html = INDEX_PATH.read_text(encoding='utf-8')
    html_nuevo = reemplazar_2025_en_html(html, filas_2025)
    INDEX_PATH.with_suffix('.html.bak').write_text(html, encoding='utf-8')
    INDEX_PATH.write_text(html_nuevo, encoding='utf-8')
    print(f"\n✅ index.html actualizado (backup en {INDEX_PATH.with_suffix('.html.bak').name})")


if __name__ == '__main__':
    main()
