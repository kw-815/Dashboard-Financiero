#!/usr/bin/env python3
"""
analisis/reporte.py
Reporte narrativo de inteligencia financiera: resumen ejecutivo, alertas de
riesgo que HOY NO cubre update_dashboard.py, y una proyección informativa
por tendencia. Solo lee datos (SharePoint o local) y usa las funciones ya
validadas de update_dashboard.py vía analisis/cargador.py.

No escribe index.html, no toca escenarios.json, no hace git add/commit/push.
El Markdown se imprime a stdout y, si --guardar, se guarda en
analisis/reportes/YYYY-MM.md (carpeta gitignored).

Uso:
    python3 analisis/reporte.py                # SharePoint
    python3 analisis/reporte.py --local        # datos/ local
    python3 analisis/reporte.py --local --guardar
"""

import argparse
import datetime
from pathlib import Path

import pandas as pd

from cargador import cargar_datos

MESES = ['', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun',
         'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']

REPORTES_DIR = Path(__file__).parent / 'reportes'

SUBGRUPOS_GASTO = ['nom', 'inst', 'viaj', 'gest', 'serv', 'otros']


def _fmt(v):
    return f"${v:,.2f}"


def seccion_resumen(pyg: pd.DataFrame, fac: pd.DataFrame, bancos_monthly: pd.DataFrame) -> str:
    out = ["## 1. Resumen ejecutivo\n"]

    if pyg.empty:
        out.append("_Sin datos de PyG._\n")
        return "\n".join(out)

    por_mes = pyg.groupby('month')[['ing', 'costo', 'gastos', 'res']].sum().sort_index()
    meses = list(por_mes.index)
    if not meses:
        out.append("_Sin meses con datos._\n")
        return "\n".join(out)

    ultimo = meses[-1]
    fila_actual = por_mes.loc[ultimo]
    out.append(f"**Último mes con datos: {MESES[ultimo]}**\n")
    out.append("| Concepto | Mes actual | Mes anterior | Variación |")
    out.append("|---|---|---|---|")

    if len(meses) >= 2:
        anterior = meses[-2]
        fila_ant = por_mes.loc[anterior]
        for campo, etiqueta in [('ing', 'Ingresos'), ('costo', 'Costos'),
                                 ('gastos', 'Gastos'), ('res', 'Resultado')]:
            actual_v, ant_v = fila_actual[campo], fila_ant[campo]
            var_pct = (actual_v - ant_v) / abs(ant_v) * 100 if ant_v else float('nan')
            var_str = f"{var_pct:+.1f}%" if var_pct == var_pct else "n/a"
            out.append(f"| {etiqueta} | {_fmt(actual_v)} | {_fmt(ant_v)} | {var_str} |")
    else:
        out.append(f"| Ingresos | {_fmt(fila_actual['ing'])} | — | — |")
        out.append(f"| Resultado | {_fmt(fila_actual['res'])} | — | — |")

    out.append("")

    if not fac.empty:
        top = (fac.groupby('cliente')['valor'].sum()
               .sort_values(ascending=False).head(5))
        total_fac = fac['valor'].sum()
        out.append(f"\n**Top 5 clientes por facturación YTD** (total: {_fmt(total_fac)})\n")
        out.append("| Cliente | Facturado | % del total |")
        out.append("|---|---|---|")
        for cliente, valor in top.items():
            pct = valor / total_fac * 100 if total_fac else 0
            out.append(f"| {cliente} | {_fmt(valor)} | {pct:.1f}% |")

    if not bancos_monthly.empty:
        bm = bancos_monthly.sort_values('month')
        if len(bm) >= 2:
            var_neta = bm.iloc[-1]['ing'] - bm.iloc[-1]['egr']
            out.append(f"\n**Flujo neto del último mes registrado:** {_fmt(var_neta)}\n")

    return "\n".join(out)


def seccion_riesgos(pyg: pd.DataFrame, fac: pd.DataFrame, cxc: pd.DataFrame,
                     hoy: datetime.date) -> str:
    out = ["## 2. Alertas de riesgo (complementarias a las de update_dashboard.py)\n"]
    alertas = []

    # Concentración de clientes
    if not fac.empty:
        total_fac = fac['valor'].sum()
        top_cliente = fac.groupby('cliente')['valor'].sum().sort_values(ascending=False)
        if not top_cliente.empty and total_fac:
            cliente, valor = top_cliente.index[0], top_cliente.iloc[0]
            pct = valor / total_fac * 100
            if pct > 25:
                alertas.append(
                    f"⚠️ Concentración de clientes: **{cliente}** representa "
                    f"{pct:.1f}% de la facturación YTD ({_fmt(valor)} de {_fmt(total_fac)})."
                )

    # Aging de CxC
    if not cxc.empty:
        cxc = cxc.copy()
        cxc['fven_dt'] = pd.to_datetime(cxc['fven_iso'], errors='coerce')
        vencidas = cxc[(cxc['estado'] == 'Vencida') & cxc['fven_dt'].notna()].copy()
        if not vencidas.empty:
            vencidas['dias_vencida'] = (pd.Timestamp(hoy) - vencidas['fven_dt']).dt.days
            buckets = [(0, 30), (30, 60), (60, 90), (90, float('inf'))]
            etiquetas = ['0-30 días', '31-60 días', '61-90 días', '90+ días']
            filas = []
            for (lo, hi), etq in zip(buckets, etiquetas):
                sub = vencidas[(vencidas['dias_vencida'] > lo) & (vencidas['dias_vencida'] <= hi)]
                if not sub.empty:
                    filas.append((etq, sub['valor'].sum(), len(sub)))
            if filas:
                out.append("**Aging de CxC vencida:**\n")
                out.append("| Antigüedad | Monto | # facturas |")
                out.append("|---|---|---|")
                for etq, monto, n in filas:
                    out.append(f"| {etq} | {_fmt(monto)} | {n} |")
                out.append("")
            monto_90 = vencidas[vencidas['dias_vencida'] > 90]['valor'].sum()
            if monto_90 > 0:
                alertas.append(
                    f"⚠️ CxC crítica: {_fmt(monto_90)} vencidos hace más de 90 días."
                )

    # Variación de gasto por subgrupo
    if not pyg.empty:
        por_mes = pyg.groupby('month')[SUBGRUPOS_GASTO].sum().sort_index()
        if len(por_mes) >= 2:
            actual, anterior = por_mes.iloc[-1], por_mes.iloc[-2]
            mes_actual = por_mes.index[-1]
            for sub in SUBGRUPOS_GASTO:
                if anterior[sub] and abs(actual[sub] - anterior[sub]) / abs(anterior[sub]) > 0.30:
                    var_pct = (actual[sub] - anterior[sub]) / abs(anterior[sub]) * 100
                    alertas.append(
                        f"⚠️ Gasto '{sub}' varió {var_pct:+.0f}% en {MESES[mes_actual]} "
                        f"vs mes anterior ({_fmt(anterior[sub])} → {_fmt(actual[sub])})."
                    )

        # Tendencia de margen neto (3 meses consecutivos a la baja)
        por_mes_res = pyg.groupby('month')[['ing', 'res']].sum().sort_index()
        por_mes_res['margen'] = por_mes_res.apply(
            lambda r: (r['res'] / r['ing'] * 100) if r['ing'] else float('nan'), axis=1)
        margenes = por_mes_res['margen'].dropna()
        if len(margenes) >= 3:
            ult3 = margenes.iloc[-3:]
            if ult3.iloc[0] > ult3.iloc[1] > ult3.iloc[2]:
                alertas.append(
                    f"⚠️ Margen neto en caída 3 meses seguidos: "
                    f"{ult3.iloc[0]:.1f}% → {ult3.iloc[1]:.1f}% → {ult3.iloc[2]:.1f}%."
                )

    if alertas:
        for a in alertas:
            out.append(f"- {a}")
    else:
        out.append("_Sin alertas adicionales este corte._")

    return "\n".join(out)


def seccion_proyeccion(pyg: pd.DataFrame) -> str:
    out = ["\n## 3. Proyección por tendencia (informativo, no se escribe a escenarios.json)\n"]

    if pyg.empty:
        out.append("_Sin datos de PyG._")
        return "\n".join(out)

    por_mes = pyg.groupby('month')[['ing', 'costo', 'gastos', 'res']].sum().sort_index()
    if len(por_mes) < 1:
        out.append("_Sin meses con datos._")
        return "\n".join(out)

    ultimo = por_mes.iloc[-1]
    ultimo_mes = por_mes.index[-1]
    promedio_3m = por_mes.iloc[-3:].mean() if len(por_mes) >= 2 else ultimo

    out.append(
        "El módulo Proyección del dashboard público asume que los meses futuros "
        f"repiten el último mes real ({MESES[ultimo_mes]}) sin cambios. "
        "Esta sección compara ese supuesto contra un promedio móvil de los últimos "
        f"{min(3, len(por_mes))} meses, para detectar si el 'run-rate plano' está "
        "sobre o subestimando la tendencia real.\n"
    )
    out.append("| Concepto | Run-rate plano (último mes) | Promedio móvil (3m) | Diferencia |")
    out.append("|---|---|---|---|")
    for campo, etiqueta in [('ing', 'Ingresos'), ('res', 'Resultado')]:
        plano = ultimo[campo]
        movil = promedio_3m[campo]
        diff_pct = (movil - plano) / abs(plano) * 100 if plano else float('nan')
        diff_str = f"{diff_pct:+.1f}%" if diff_pct == diff_pct else "n/a"
        out.append(f"| {etiqueta} | {_fmt(plano)} | {_fmt(movil)} | {diff_str} |")

    return "\n".join(out)


def generar_reporte(modo='local', guardar=False):
    datos = cargar_datos(modo=modo)
    hoy = datetime.date.today()

    # pyg ahora trae todos los años disponibles (columna 'anio'); las
    # secciones comparan mes a mes dentro de un mismo año, así que se filtra
    # al año más reciente con datos.
    pyg = datos['pyg']
    pyg_actual = pyg[pyg['anio'] == pyg['anio'].max()] if not pyg.empty else pyg

    partes = [
        f"# Reporte de inteligencia financiera — {hoy.isoformat()}\n",
        seccion_resumen(pyg_actual, datos['facturacion'], datos['bancos_monthly']),
        seccion_riesgos(pyg_actual, datos['facturacion'], datos['cxc'], hoy),
        seccion_proyeccion(pyg_actual),
    ]
    reporte_md = "\n\n".join(partes)
    print(reporte_md)

    if guardar:
        REPORTES_DIR.mkdir(exist_ok=True)
        destino = REPORTES_DIR / f"{hoy.strftime('%Y-%m')}.md"
        destino.write_text(reporte_md, encoding='utf-8')
        print(f"\n\n(guardado en {destino})")

    return reporte_md


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Reporte de inteligencia financiera')
    parser.add_argument('--local', action='store_true', help='Usar datos/ local en vez de SharePoint')
    parser.add_argument('--guardar', action='store_true', help='Guardar el reporte en analisis/reportes/')
    args = parser.parse_args()

    generar_reporte(modo='local' if args.local else 'sharepoint', guardar=args.guardar)
