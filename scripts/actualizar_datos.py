"""
╔══════════════════════════════════════════════════════════════════════╗
║         COMCEC · INTERCEC — ACTUALIZADOR DE DATOS PARA DASHBOARDS   ║
╠══════════════════════════════════════════════════════════════════════╣
║  Ejecutar después de actualizar cualquier archivo Excel fuente.      ║
║  Genera los archivos .js y .json en la carpeta /data/.               ║
║                                                                      ║
║  Uso:  python scripts/actualizar_datos.py                            ║
╚══════════════════════════════════════════════════════════════════════╝
"""

import os
import json
import math
from datetime import datetime, date

# ── Dependencias externas ─────────────────────────────────────────────
try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("ERROR: Instala las dependencias con:  pip install pandas openpyxl numpy")
    raise SystemExit(1)

# ── Rutas ─────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

EXCEL = {
    "presupuesto":   "COMCEC - PRESUPUESTO_SEGUIMIENTO_alejandro.xlsx",
    "avances":       "Control Avances de INF SEM -- INF MES - INTERCEC.xlsx",
    "actividades":   "COMCEC - XML.xml-v00-Activities.xlsx",
    "subcontratos":  "COMCEC - Gastos_Obra_civilco.xlsx",
    "fondo_rot":     "COMCEC - MOV_FONDO_ROT.xlsx",
    "workplan":      "COMCEC - Control WorkPlan.xlsx",
    "trimble":       "COMCEC - CecFenix_FaseI__Files.xlsx",
}

FECHA_CORTE = datetime.now().strftime("%d-%b-%Y")

# ─────────────────────────────────────────────────────────────────────
# Utilidades
# ─────────────────────────────────────────────────────────────────────

def _ser(obj):
    """Serializa cualquier valor a tipos JSON-compatibles."""
    if obj is None:
        return None
    try:
        if pd.isna(obj):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(obj, (datetime, date, pd.Timestamp)):
        return obj.strftime("%Y-%m-%d")
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating, float)):
        if math.isnan(obj) or math.isinf(obj):
            return None
        return round(float(obj), 4)
    if isinstance(obj, (np.ndarray, list)):
        return [_ser(x) for x in obj]
    return str(obj).strip()


def _clean(row_dict):
    return {str(k).strip().replace("\n", " "): _ser(v) for k, v in row_dict.items()}


def _save(base_name, var_name, payload):
    """Guarda el dataset como .json y como .js (para uso offline con file://)."""
    json_path = os.path.join(DATA_DIR, f"{base_name}.json")
    js_path   = os.path.join(DATA_DIR, f"{base_name}.js")
    dumped    = json.dumps(payload, ensure_ascii=False)
    with open(json_path, "w", encoding="utf-8") as f:
        f.write(json.dumps(payload, ensure_ascii=False, indent=2))
    with open(js_path, "w", encoding="utf-8") as f:
        f.write(f"window.{var_name} = {dumped};\n")
    size_kb = os.path.getsize(js_path) // 1024
    print(f"   OK  {base_name}.js  ({size_kb} KB)")


def _xls(key, **kwargs):
    """Lee un Excel por clave del diccionario EXCEL."""
    path = os.path.join(BASE_DIR, EXCEL[key])
    if not os.path.exists(path):
        raise FileNotFoundError(f"No se encontró: {EXCEL[key]}")
    return pd.read_excel(path, **kwargs)


def _float(val):
    try:
        v = float(val)
        return 0.0 if math.isnan(v) or math.isinf(v) else v
    except (TypeError, ValueError):
        return 0.0


# ─────────────────────────────────────────────────────────────────────
# 1. PRESUPUESTO
# ─────────────────────────────────────────────────────────────────────
def procesar_presupuesto():
    print("\n[1/7] Presupuesto...")
    xl_path = os.path.join(BASE_DIR, EXCEL["presupuesto"])
    xl = pd.ExcelFile(xl_path)

    df_all = pd.read_excel(xl_path, sheet_name="Control presupuestal financiero")

    # ── Filtrar solo filas de Centros de Costo (CC013-x.xx) — igual que el BI ──
    df = df_all[df_all["Centro de Costo"].astype(str).str.match(r"CC\d+")].copy()

    def _c(col): return df[col].fillna(0)

    total_asig  = _float(_c("Asignación Presupuestal").sum())
    total_cont  = _float(_c("Valor Contratado").sum())          # Tot. Vlr Contratado (BI)
    total_pag   = _float(_c("Total Pagado - Valor").sum())      # Tot. Vlr Pagado (BI)
    total_pend  = _float(_c("Pendiente por Pagar - Valor").sum())

    pct_cont_asig   = round(total_cont / total_asig * 100 if total_asig else 0, 2)  # 93.76%
    pct_pag_cont    = round(total_pag  / total_cont * 100 if total_cont else 0, 2)  # 62.40%

    # ── Agrupado por Tipo de costo para gráficas ──
    por_tipo = (
        df.groupby("Tipo de costo")
          .agg({
              "Asignación Presupuestal": "sum",
              "Valor Contratado":        "sum",
              "Total Pagado - Valor":    "sum",
          })
          .reset_index()
    )
    tipo_data = [
        {
            "tipo":       str(r["Tipo de costo"]),
            "asignado":   _ser(r["Asignación Presupuestal"]),
            "contratado": _ser(r["Valor Contratado"]),
            "pagado":     _ser(r["Total Pagado - Valor"]),
        }
        for _, r in por_tipo.iterrows()
    ]

    # ── Histórico de cortes semanales ──
    corte_sheets = sorted([s for s in xl.sheet_names if s.startswith("Control presup-")], reverse=True)
    historico = []
    for s in corte_sheets[:15]:
        try:
            dc = pd.read_excel(xl_path, sheet_name=s)
            dcc = dc[dc["Centro de Costo"].astype(str).str.match(r"CC\d+")].copy()
            ca  = _float(dcc["Asignación Presupuestal"].fillna(0).sum())
            cco = _float(dcc["Valor Contratado"].fillna(0).sum())
            cp  = _float(dcc["Total Pagado - Valor"].fillna(0).sum())
            historico.append({
                "corte":        s.replace("Control presup-", ""),
                "asignado":     ca,
                "contratado":   cco,
                "pagado":       cp,
                "pct_cont":     round(cco / ca * 100 if ca else 0, 2),
                "pct_pag_cont": round(cp  / cco * 100 if cco else 0, 2),
            })
        except Exception:
            pass
    historico.reverse()

    # ── Filas de detalle por CC — columnas exactas del BI ──
    filas_cc = []
    for _, r in df.iterrows():
        asig  = _float(r.get("Asignación Presupuestal"))
        cont  = _float(r.get("Valor Contratado"))
        pag   = _float(r.get("Total Pagado - Valor"))
        pend  = _float(r.get("Pendiente por Pagar - Valor"))
        dif   = round(asig - cont, 2)   # Diferencia (Asignado - Contratado)
        filas_cc.append({
            "cc":           str(r.get("Centro de Costo") or ""),
            "cuenta":       str(r.get("Cuentas de Control") or ""),
            "tipo":         str(r.get("Tipo de costo") or ""),
            "asignado":     asig,
            "contratado":   cont,
            "diferencia":   dif,
            "pagado":       pag,
            "pendiente":    pend,
            "pct_cont_asig":  round(cont / asig * 100 if asig else 0, 2),
            "pct_pag_cont":   round(pag  / cont * 100 if cont else 0, 2),
        })

    _save("presupuesto", "DATA_PRESUPUESTO", {
        "kpis": {
            "total_asignado":       total_asig,
            "total_contratado":     total_cont,
            "total_pagado":         total_pag,
            "total_pendiente":      total_pend,
            "pct_cont_asig":        pct_cont_asig,
            "pct_pag_cont":         pct_pag_cont,
            "fecha_corte":          FECHA_CORTE,
        },
        "por_tipo":         tipo_data,
        "historico_cortes": historico,
        "detalles":         filas_cc,
    })


# ─────────────────────────────────────────────────────────────────────
# 2. AVANCES % — CURVA S
# ─────────────────────────────────────────────────────────────────────
def procesar_avances():
    print("\n[2/7] Avances % · Curva S...")
    df = _xls("avances", sheet_name="SEMANALES")
    curvas = []
    for _, row in df.iterrows():
        corte   = row.get("CORTE")
        inf_sem = row.get("inf sem")
        if pd.isna(corte) and pd.isna(inf_sem):
            continue
        
        # Desfases y SPI
        spi_val = row.get("PROYECTO - SPI %ejec/%prog")
        if pd.isna(spi_val):
            p_val = _float(row.get("PLAN EJECUCIÓN % ACUMULADA SEMANAL PROJECT"))
            r_val = _float(row.get("REAL EJECUCIÓN % ACUMULADA SEMANAL PROJECT"))
            spi_val = round(r_val / p_val, 2) if p_val > 0 else 1.0

        curvas.append({
            "semana":                  _ser(inf_sem),
            "corte":                   _ser(corte),
            "prog_revisado":           _ser(row.get("PROG revisado")),
            "ejec_revisado":           _ser(row.get("EJEC revisado")),
            "plan_acumulado":          _ser(row.get("PLAN EJECUCIÓN % ACUMULADA SEMANAL PROJECT")),
            "real_acumulado":          _ser(row.get("REAL EJECUCIÓN % ACUMULADA SEMANAL PROJECT")),
            "desfase_proyecto":        _ser(row.get("PROYECTO % DESFASE")),
            "spi":                     _ser(spi_val),
            
            # Subcapítulos
            "prog_preliminares":       _ser(row.get("PROGR\n1.1\nPRELIMINARES")),
            "ejec_preliminares":       _ser(row.get("EJEC\n1.1\nPRELIMINARES")),
            "desfase_preliminares":    _ser(row.get("Preliminares % DESFASE")),
            
            "prog_subcontrataciones":  _ser(row.get("PROG\n1.2\nSUBCONTRATACIONES")),
            "ejec_subcontrataciones":  _ser(row.get("EJEC\n1.2\nSUBCONTRATACIONES")),
            "desfase_subcontrataciones": _ser(row.get("Sub contra\ntaciones \n% DESFASE")),
            
            "prog_construccion":       _ser(row.get("PROG\n1.3\nCONSTRUCC")),
            "ejec_construccion":       _ser(row.get("EJEC\n1.3\nCONSTRUCC")),
            "desfase_construccion":    _ser(row.get("delta EN CONSTRUCCION") or row.get("CONSTRUCCION % DESFASE")),

            # Contratistas principales
            "coveco_prog":             _ser(row.get("DINACOL - % PROG")),
            "coveco_ejec":             _ser(row.get("DINACOL - % EJEC")),
            "coveco_desfase":          _ser(row.get("DINACOL - % DESFASE")),

            "ingecol_prog":            _ser(row.get("INGECOL - % PROG")),
            "ingecol_ejec":            _ser(row.get("INGECOL - % EJEC")),
            "ingecol_desfase":         _ser(row.get("INGECOL - % DESFASE")),

            "mosel_prog":              _ser(row.get("MOSEL - % PROG")),
            "mosel_ejec":              _ser(row.get("MOSEL- % EJEC")),
            "mosel_desfase":           _ser(row.get("MOSEL- % DESFASE")),

            "ch_pereira_prog":         _ser(row.get("CH PEREIRA - % PROG")),
            "ch_pereira_ejec":         _ser(row.get("CH PEREIRA - % EJEC")),
            "ch_pereira_desfase":      _ser(row.get("CH PEREIRA - % DESFASE")),
        })
    curvas = [c for c in curvas if c["corte"] or c["semana"]]
    ultimo = curvas[-1] if curvas else {}
    _save("avances_curvas", "DATA_AVANCES", {
        "kpis": {
            "ultimo_avance_real": ultimo.get("real_acumulado", 0),
            "ultimo_avance_plan": ultimo.get("plan_acumulado", 0),
            "desviacion":        round(
                (_float(ultimo.get("real_acumulado") or 0) - _float(ultimo.get("plan_acumulado") or 0)),
                4
            ),
            "total_semanas": len(curvas),
            "corte":         ultimo.get("corte") or FECHA_CORTE,
        },
        "curvas_semanales": curvas,
    })


# ─────────────────────────────────────────────────────────────────────
# 3. ACTIVIDADES PDT  (retorna lista para que Hitos la reutilice)
# ─────────────────────────────────────────────────────────────────────
def procesar_actividades():
    print("\n[3/7] Actividades PDT...")
    df = _xls("actividades", sheet_name="ACTIVITIES")
    actividades = []
    tot_comp = tot_prog = tot_not = 0
    for _, row in df.iterrows():
        act_id   = _ser(row.get("Activity ID"))
        act_name = _ser(row.get("Activity Name"))
        if not act_id and not act_name:
            continue
        raw_st = str(_ser(row.get("Activity Status")) or "")
        if "Complete" in raw_st:
            status = "Completed";   tot_comp += 1
        elif "Progress" in raw_st:
            status = "In Progress"; tot_prog += 1
        else:
            status = "Not Started"; tot_not  += 1
        pct = _ser(row.get("Physical % Complete"))
        if pct is None:
            pct = 100 if status == "Completed" else (50 if status == "In Progress" else 0)
        
        # Paquetes clave identificados en el Power BI
        paquetes = []
        if pd.notna(row.get("CARRIL")): paquetes.append("CARRIL")
        if pd.notna(row.get("PILOTAJE")): paquetes.append("PILOTAJE")
        if pd.notna(row.get("CIMENTACION Y ESTRUCTURA")): paquetes.append("CIMENTACION Y ESTRUCTURA")
        if pd.notna(row.get("MAMPOSTERIA I")): paquetes.append("MAMPOSTERIA I")
        if pd.notna(row.get("MAMPOSTERIA II")): paquetes.append("MAMPOSTERIA II")
        if pd.notna(row.get("TERRACEO")) or pd.notna(row.get("TERRACEO2")): paquetes.append("TERRACEO")
        if pd.notna(row.get("RED ELEC EXT")): paquetes.append("RED ELEC EXT")
        if pd.notna(row.get("SUB CONTRATACIONES 1.2.1")): paquetes.append("SUB CONTRATACIONES 1.2.1")
        if pd.notna(row.get("SUB CONTRATACIONES 1.2.2")): paquetes.append("SUB CONTRATACIONES 1.2.2")

        actividades.append({
            "id":             act_id,
            "name":           act_name,
            "status":         status,
            "wbs_name":       _ser(row.get("WBS Name")),
            "wbs_code":       _ser(row.get("WBS Code")),
            "edt_code":       _ser(row.get("EDT CODE")),
            "planned_start":  _ser(row.get("Planned Start")),
            "planned_finish": _ser(row.get("Planned Finish")),
            "duration":       _ser(row.get("Original Duration")),
            "total_float":    _ser(row.get("Total Float")),
            "pct_complete":   pct,
            "paquetes":       paquetes,
        })

    wbs = {}
    for a in actividades:
        w = a["wbs_name"] or "General"
        if w not in wbs:
            wbs[w] = {"total": 0, "completed": 0, "in_progress": 0, "not_started": 0}
        wbs[w]["total"] += 1
        wbs[w][a["status"].lower().replace(" ", "_")] += 1

    wbs_list = [{"wbs": k, **v} for k, v in sorted(wbs.items(), key=lambda x: x[1]["total"], reverse=True)]
    _save("actividades_pdt", "DATA_ACTIVIDADES", {
        "kpis": {
            "total_actividades": len(actividades),
            "completadas":       tot_comp,
            "en_progreso":       tot_prog,
            "sin_iniciar":       tot_not,
            "pac_pct":           round(tot_comp / len(actividades) * 100, 1) if actividades else 0,
        },
        "wbs_summary": wbs_list[:15],
        "actividades": actividades,
    })
    return actividades  # reutilizado por hitos


# ─────────────────────────────────────────────────────────────────────
# 4. HITOS PDT
# ─────────────────────────────────────────────────────────────────────
def procesar_hitos(actividades):
    print("\n[4/7] Hitos PDT...")
    hitos = [
        a for a in actividades
        if "HITO" in (a["name"] or "").upper()
        or "ENTREGA" in (a["name"] or "").upper()
        or a["duration"] == 0
    ][:60]
    _save("hitos_pdt", "DATA_HITOS", {
        "kpis": {
            "total_hitos": len(hitos),
            "cumplidos":   sum(1 for h in hitos if h["status"] == "Completed"),
        },
        "hitos": hitos,
    })


# ─────────────────────────────────────────────────────────────────────
# 5. SUBCONTRATOS
# ─────────────────────────────────────────────────────────────────────
def procesar_subcontratos():
    print("\n[5/7] Subcontratos...")
    df = _xls("subcontratos", sheet_name="CONTRATOS Y ORD")
    subs = []
    tot_comp = tot_pag = 0.0
    estados_set = set()
    tipos_set = set()
    ccostos_set = set()

    for _, row in df.iterrows():
        n_cont      = _ser(row.get("Nº CONTRATO"))
        contratista = _ser(row.get("EMPRESA CONTRATISTA"))
        if not n_cont and not contratista:
            continue
        vc = _float(row.get("VALOR COMPROMISO ACTUALIZADO"))
        vp = _float(row.get("VALOR OBLIGADO PAGADO"))
        tot_comp += vc
        tot_pag  += vp

        est  = _ser(row.get("ESTADO")) or "EN EJECUCIÓN"
        tipo = _ser(row.get("TIPO CONTRATACIÓN")) or "CONTRATO"
        cc   = _ser(row.get(" CCOSTOS")) or ""
        det_cc = _ser(row.get("Detalle - CCOSTOS")) or ""

        if est: estados_set.add(est)
        if tipo: tipos_set.add(tipo)
        if cc: ccostos_set.add(cc)

        subs.append({
            "codigo":           n_cont,
            "contratista":      contratista,
            "objeto":           _ser(row.get("OBJETO DEL CONTRATO")),
            "estado":           est,
            "tipo":             tipo,
            "ccosto":           cc,
            "detalle_ccosto":   det_cc,
            "fecha_inicio":     _ser(row.get("FECHA DE INICIO")),
            "fecha_fin":        _ser(row.get("FECHA DE FINALIZACIÓN")),
            "valor_compromiso": vc,
            "valor_pagado":     vp,
            "pct_ejecutado":    _ser(row.get("% EJECUTADO")),
            "pct_facturado":    _ser(row.get("% FACTURADO")),
            "saldo_por_pagar":  round(vc - vp, 2),
            "pct_pagado":       round(vp / vc * 100, 1) if vc else 0,
        })

    _save("subcontratos", "DATA_SUBCONTRATOS", {
        "kpis": {
            "total_contratos":       len(subs),
            "total_compromiso":      tot_comp,
            "total_obligado_pagado": tot_pag,
            "saldo_total":           round(tot_comp - tot_pag, 2),
            "pct_ejecucion_pagos":   round(tot_pag / tot_comp * 100, 1) if tot_comp else 0,
            "fecha_corte":           "20/09/2026",
        },
        "filtros": {
            "estados": sorted(list(estados_set)),
            "tipos":   sorted(list(tipos_set)),
            "ccostos": sorted(list(ccostos_set)),
        },
        "subcontratos": subs,
    })


# ─────────────────────────────────────────────────────────────────────
# 6. FONDO ROTATORIO
# ─────────────────────────────────────────────────────────────────────
def procesar_fondo_rotatorio():
    print("\n[6/7] Fondo Rotatorio...")
    df = _xls("fondo_rot", sheet_name="MOV. FONDO ROT.")
    movs = []
    tot_fact = tot_neto = tot_ant = 0.0
    estados_pago_set = set()
    estados_rev_set = set()
    fuentes_set = set()
    ccostos_set = set()

    for _, row in df.iterrows():
        reg     = _ser(row.get("Registro No."))
        tercero = _ser(row.get("TERCERO"))
        if not reg and not tercero:
            continue
        vf = _float(row.get("VALOR TOTAL FACTURA"))
        vn = _float(row.get("VALOR NETO FR"))
        va = _float(row.get("ANTICIPO"))
        tot_fact += vf; tot_neto += vn; tot_ant += va

        ep = _ser(row.get("ESTADO DE PAGO")) or "Revisado"
        er = _ser(row.get("ESTADO REVISIÓN")) or "Sin estado"
        fp = _ser(row.get("FUENTE DE PAGO")) or _ser(row.get("FUENTE RECURSOS")) or ""
        cc = _ser(row.get(" CENTRO DE COSTOS")) or ""

        if ep: estados_pago_set.add(ep)
        if er: estados_rev_set.add(er)
        if fp: fuentes_set.add(fp)
        if cc: ccostos_set.add(cc)

        movs.append({
            "registro_no":     reg,
            "contrato_orden":  _ser(row.get("CONT/ORD")),
            "tercero":         tercero,
            "ccosto":          cc,
            "detalle_ccosto":  _ser(row.get("Detalle - CC013")),
            "doc_reembolso":   _ser(row.get("Doc. REEMBOLSO")),
            "descripcion":     _ser(row.get("DESCRIPCIÓN")),
            "factura":         _ser(row.get("FACT/DOC EQ.")),
            "valor_factura":   vf,
            "valor_neto_fr":   vn,
            "anticipo":        va,
            "estado_pago":     ep,
            "estado_revision": er,
            "fuente_pago":     fp,
            "estado_fr":       _ser(row.get("ESTADO FR")),
            "fecha_pago":      _ser(row.get("FECHA DE PAGO")),
            "fecha_radicado":  _ser(row.get("FECHA Rad. Cert.")),
        })

    por_estado = {}
    for m in movs:
        e = m["estado_pago"] or "Revisado"
        por_estado[e] = por_estado.get(e, 0) + m["valor_neto_fr"]

    _save("fondo_rotatorio", "DATA_FONDO_ROTATORIO", {
        "kpis": {
            "total_registros":  len(movs),
            "total_facturado":  tot_fact,
            "total_neto_fr":    tot_neto,
            "total_anticipos":  tot_ant,
            "fecha_corte":      "20/09/2026",
        },
        "filtros": {
            "estados_pago":     sorted(list(estados_pago_set)),
            "estados_revision": sorted(list(estados_rev_set)),
            "fuentes_pago":     sorted(list(fuentes_set)),
            "ccostos":          sorted(list(ccostos_set)),
        },
        "por_estado_pago": [{"estado": k, "valor": v} for k, v in por_estado.items()],
        "movimientos":     movs,
    })


# ─────────────────────────────────────────────────────────────────────
# 7. WORKPLAN
# ─────────────────────────────────────────────────────────────────────
def procesar_workplan():
    print("\n[7/7] WorkPlan...")
    xl_path = os.path.join(BASE_DIR, EXCEL["workplan"])
    xl = pd.ExcelFile(xl_path)
    
    # Buscar hoja de corte más reciente o '20sep26'
    target_sheet = xl.sheet_names[0]
    for s in ['20sep26', '30ago26', '19jul26', '12jul26']:
        if s in xl.sheet_names:
            target_sheet = s
            break

    df = pd.read_excel(xl_path, sheet_name=target_sheet)

    def parse_cc_edt(cc_val, edt_val):
        cc_s = str(cc_val).strip() if pd.notna(cc_val) else ""
        if cc_s.startswith('CC013-1.'):
            sub = cc_s.replace('CC013-1.', '')
            parts = sub.split('.')
            e2 = int(parts[0]) if len(parts) > 0 and parts[0].isdigit() else None
            e3 = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else None
            e4 = int(parts[2]) if len(parts) > 2 and parts[2].isdigit() else None
            return 1, e2, e3, e4
        elif cc_s == 'CC013-1':
            return 1, None, None, None
        else:
            return None, None, None, None

    items = []
    for _, row in df.iterrows():
        cc  = _ser(row.get("CENTRO DE COSTOS"))
        edt = _ser(row.get("EDT"))
        if not cc and not edt:
            continue
        
        e1, e2, e3, e4 = parse_cc_edt(cc, edt)
        
        vb  = _float(row.get("VALOR PRESUPUESTO BASE"))
        vpy = _float(row.get("VALOR PROYECCIÓN PRESUPUESTO", row.get("VALOR PROYECCIN PRESUPUESTO", row.get("VALOR PROYECCI\u00d3N PRESUPUESTO"))))
        vtc = _float(row.get("VALOR TOTAL CONTRATADO"))
        va  = _float(row.get("ADICIONES Y SUSTRACCIONES A LOS CENTROS DE COSTOS"))
        vc  = _float(row.get("VALOR CONTRATO", row.get("VALOR CONTRATOS")))
        vo  = _float(row.get("VALOR ÓRDENES", row.get("VALOR \u00d3RDENES", row.get("VALOR ORDENES"))))

        items.append({
            "ccosto":                  cc,
            "edt":                     edt,
            "edt1":                    e1,
            "edt2":                    e2,
            "edt3":                    e3,
            "edt4":                    e4,
            "actividad":               _ser(row.get("ACTIVIDAD")),
            "unidad":                  _ser(row.get(" UNIDAD")),
            "cantidad":                _ser(row.get("CANTIDAD ")),
            "valor_base":              vb,
            "valor_proyeccion":        vpy,
            "valor_total_contratado":  vtc,
            "adiciones_sustracciones": va,
            "valor_contrato":          vc,
            "valor_ordenes":           vo,
        })

    # Filtrar capítulos EDT 2 para los totales maestros (excluyendo la fila resumen CC013-1 y los subniveles)
    # En Power BI, el total general corresponde a los 16 capítulos EDT.2 (CC013-1.1 al CC013-1.16)
    capitulos = [i for i in items if i["edt2"] is not None and i["edt3"] is None and i["edt4"] is None]
    
    tot_base  = sum(c["valor_base"] for c in capitulos)
    tot_proy  = sum(c["valor_proyeccion"] for c in capitulos)
    tot_cont  = sum(c["valor_total_contratado"] for c in capitulos)
    tot_adic  = sum(c["adiciones_sustracciones"] for c in capitulos)
    tot_vcont = sum(c["valor_contrato"] for c in capitulos)
    tot_vord  = sum(c["valor_ordenes"] for c in capitulos)

    _save("workplan", "DATA_WORKPLAN", {
        "kpis": {
            "total_items":                len(items),
            "total_presupuesto_base":     tot_base,
            "total_proyeccion":           tot_proy,
            "total_contratado":           tot_cont,
            "total_adiciones_sustracciones": tot_adic,
            "total_valor_contratos":      tot_vcont,
            "total_valor_ordenes":        tot_vord,
            "fecha_corte":                "20/09/2026",
        },
        "capitulos": capitulos,
        "items": items,
    })


# ─────────────────────────────────────────────────────────────────────
# 8. ARCHIVOS TRIMBLE (CecFenix)
# ─────────────────────────────────────────────────────────────────────
def procesar_trimble():
    print("\n[+] Archivos Trimble (CecFenix)...")
    df = _xls("trimble", sheet_name=0)
    files = []
    for _, row in df.iterrows():
        fname = str(row.get("File name") or "").strip()
        if not fname or fname == "nan":
            continue
        files.append({
            "name":        fname,
            "location":    str(row.get("File location")  or "").strip(),
            "location2":   str(row.get("File location2") or "").strip(),
            "size":        str(row.get("Size")           or "").strip(),
            "created_by":  str(row.get("Created by")     or "").strip(),
            "created_on":  str(row.get("Created on")     or "")[:10],
            "modified_by": str(row.get("Last modified by") or "").strip(),
            "modified_on": str(row.get("Last modified on") or "")[:10],
        })

    def _agg(key):
        counts = {}
        for f in files:
            u = f[key]
            if u and u != "nan":
                counts[u] = counts.get(u, 0) + 1
        return sorted([{"user": k, "count": v} for k, v in counts.items()], key=lambda x: x["count"], reverse=True)

    loc2 = {}
    for f in files:
        l = f["location2"]
        if l and l != "nan":
            loc2[l] = loc2.get(l, 0) + 1
    top_loc2 = sorted([{"location": k, "count": v} for k, v in loc2.items()], key=lambda x: x["count"], reverse=True)[:100]

    # Rango de fechas real
    dates_mod = [f["modified_on"] for f in files if f["modified_on"] and f["modified_on"] != "nan"]
    fecha_min = min(dates_mod) if dates_mod else "—"
    fecha_max = max(dates_mod) if dates_mod else "—"

    _save("archivos_trimble", "DATA_TRIMBLE", {
        "kpis": {
            "total_documentos":    len(files),
            "fecha_min_mod":       fecha_min,
            "fecha_max_mod":       fecha_max,
            "fecha_actualizacion": FECHA_CORTE,
        },
        "created_by_summary":  _agg("created_by"),
        "modified_by_summary": _agg("modified_by"),
        "locations_summary":   top_loc2,
        "sample_files":        files[:3000],
    })


# ─────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 60)
    print("  COMCEC · INTERCEC — Actualizador de Datos")
    print(f"  Fecha de corte: {FECHA_CORTE}")
    print("=" * 60)

    errores = []

    def run(fn, *args):
        try:
            return fn(*args)
        except Exception as e:
            print(f"   ERROR en {fn.__name__}: {e}")
            errores.append((fn.__name__, str(e)))
            return None

    run(procesar_presupuesto)
    run(procesar_avances)
    acts = run(procesar_actividades)
    if acts:
        run(procesar_hitos, acts)
    run(procesar_subcontratos)
    run(procesar_fondo_rotatorio)
    run(procesar_workplan)
    run(procesar_trimble)

    print("\n" + "=" * 60)
    if errores:
        print(f"  Completado con {len(errores)} error(es):")
        for nombre, msg in errores:
            print(f"    - {nombre}: {msg}")
    else:
        print("  [OK] Todos los datasets actualizados correctamente.")
    print("=" * 60)
