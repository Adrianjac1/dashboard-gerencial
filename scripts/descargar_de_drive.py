"""
╔══════════════════════════════════════════════════════════════════════╗
║     DESCARGADOR DE ARCHIVOS GOOGLE DRIVE PARA GITHUB ACTIONS        ║
╠══════════════════════════════════════════════════════════════════════╣
║  Descarga automáticamente los 8 Excels fuente y el PBIX desde Drive ║
║  utilizando gdown de forma directa y confiable.                      ║
╚══════════════════════════════════════════════════════════════════════╝
"""

import os
import sys

# Mapeo exacto de nombres de archivo a IDs en Google Drive
DRIVE_FILES = {
    # 1. Archivos Trimble (CecFenix)
    "COMCEC - CecFenix_FaseI__Files.xlsx": os.environ.get(
        "ID_TRIMBLE", "1ZajqJ40G5JmFEyRBAxwx4s7XhLXlkKpw"
    ),
    # 2. WorkPlan
    "COMCEC - Control WorkPlan.xlsx": os.environ.get(
        "ID_WORKPLAN", "1qvzNJySR5bB3E80km1CVXamSIGMItmAU"
    ),
    # 3. Subcontratos (Gastos Obra)
    "COMCEC - Gastos_Obra_civilco.xlsx": os.environ.get(
        "ID_SUBCONTRATOS", "1Ik25sGcBOlvi0mCOOm0phDWl17C5crX9"
    ),
    # 4. Fondo Rotatorio
    "COMCEC - MOV_FONDO_ROT.xlsx": os.environ.get(
        "ID_FONDO_ROT", "1SG1tGq0xD108VgSZd99gcoqlGc8M9LVK"
    ),
    # 5. Presupuesto
    "COMCEC - PRESUPUESTO_SEGUIMIENTO_alejandro.xlsx": os.environ.get(
        "ID_PRESUPUESTO", "1sIKoAcdsY2fjyp79BX4JfFk5DGH0phch"
    ),
    # 6. Registro de Lluvias en Obra
    "COMCEC - Registro de LLUVIAS en obra ahr.xlsx": os.environ.get(
        "ID_LLUVIAS", "1kUCr_P02cjkUmMO_f6BgwnS3P0vvspBE"
    ),
    # 7. Actividades PDT
    "COMCEC - XML.xml-v00-Activities.xlsx": os.environ.get(
        "ID_ACTIVIDADES", "1eLAMf90oPm9n21Q7DVB4vTV0hdx3DlYj"
    ),
    # 8. Curvas de Avance
    "Control Avances de INF SEM -- INF MES - INTERCEC.xlsx": os.environ.get(
        "ID_AVANCES", "1C6ij0XvWB4-WiHD5g6rbtptEqPG26f3g"
    ),
    # 9. Power BI Gerencial (Opcional, para respaldo)
    "COMCEC - Gerencial_02.pbix": os.environ.get(
        "ID_PBIX", "1ALEdrVWv9v4IZXzkNCZI5DGf4L4h5DNo"
    ),
}

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def descargar():
    try:
        import gdown
    except ImportError:
        print("ERROR: Instala gdown con: pip install gdown")
        sys.exit(1)

    print("=" * 60)
    print("  Iniciando descarga de archivos desde Google Drive...")
    print("=" * 60)

    exitos = 0
    fallos = 0

    for filename, file_id in DRIVE_FILES.items():
        if not file_id:
            print(f"  [AVISO] Omitiendo {filename} (ID no configurado)")
            continue

        dest_path = os.path.join(BASE_DIR, filename)
        url = f"https://drive.google.com/uc?id={file_id}"
        print(f"  -> Descargando: {filename}...")
        try:
            gdown.download(url, dest_path, quiet=True)
            if os.path.exists(dest_path) and os.path.getsize(dest_path) > 1000:
                size_kb = os.path.getsize(dest_path) // 1024
                print(f"     [OK] Descargado ({size_kb} KB)")
                exitos += 1
            else:
                print(f"     [!] Advertencia: archivo descargado vacío o con error de permisos.")
                fallos += 1
        except Exception as e:
            print(f"     [ERROR] Falló al descargar {filename}: {e}")
            fallos += 1

    print("=" * 60)
    print(f"  Descarga completada: {exitos} exitosos, {fallos} fallidos.")
    print("=" * 60)
    if fallos > 0:
        sys.exit(1)


if __name__ == "__main__":
    descargar()
