import os
from utils import (
    read_urls,
    ensure_folder,
    ensure_reports_folder
)
from downloader import download_instagram_video
    # returns: (ok, mp4_path_or_error, meta_txt_path)

from analyzer import analyze_video_with_qwen
    # returns: texto generado por Qwen

from reporter import write_report

# NUEVO MÓDULO (tu versión renombrada)
from excel_filler import generate_excel_from_multiple_txt


def main():

    # ---------------------------------------------------------
    # 0. Preparación de carpetas
    # ---------------------------------------------------------
    ensure_folder("recipes_videos")
    reports_dir = ensure_reports_folder()

    # ---------------------------------------------------------
    # 1. Leer URLs
    # ---------------------------------------------------------
    urls = read_urls()
    report_entries = []

    # Lista donde guardamos los QWEN.txt generados EN ORDEN
    generated_qwen_files = []

    # ---------------------------------------------------------
    # 2. Procesar cada URL en orden
    # ---------------------------------------------------------
    for url in urls:
        print(f"\nProcessing: {url}")

        # 2.1 Descargar vídeo
        ok, mp4_or_error, meta_txt = download_instagram_video(url)
        if not ok:
            report_entries.append(f"[ERROR] {url} → Download failed: {mp4_or_error}")
            continue

        mp4_path = mp4_or_error

        # 2.2 Analizar vídeo con Qwen
        try:
            generated_text = analyze_video_with_qwen(mp4_path, url)
        except Exception as e:
            report_entries.append(f"[ERROR] {url} → Qwen failed: {str(e)}")
            continue

        # Guardar el texto generado en un archivo .txt
        qwen_txt_path = mp4_path.replace(".mp4", "_QWEN.txt")
        with open(qwen_txt_path, "w", encoding="utf-8") as f:
            f.write(generated_text)

        # Guardamos el QWEN.txt en orden EXACTO
        generated_qwen_files.append(qwen_txt_path)

        report_entries.append(f"[OK] {url} → {qwen_txt_path}")

    # ---------------------------------------------------------
    # 3. Escribir summary/report
    # ---------------------------------------------------------
    write_report(report_entries, reports_dir)

    # ---------------------------------------------------------
    # 4. (DESACTIVADO) Generar Excel final
    # ---------------------------------------------------------
    if generated_qwen_files:
        print("\n=== GENERATING FINAL EXCEL ===")
        TEMPLATE_PATH = "plantilla.xlsx" 
    
        output_excel = generate_excel_from_multiple_txt(
            template_xlsx_path=TEMPLATE_PATH,
            qwen_txt_paths=generated_qwen_files
        )
    
        print(f"\n[OK] Excel final generado → {output_excel}")
    
        # ---------------------------------------------------------
        # 5. EJECUTAR SCRIPT CONFIDENCIAL SOLO SI EL EXCEL EXISTE
        # ---------------------------------------------------------
        import subprocess
    
        if os.path.exists(output_excel):
            print("\n=== EJECUTANDO SCRIPT CONFIDENCIAL ===")
    
            CONFIDENTIAL_SCRIPT = "ejecutar_BSH_automatico.py"
    
            subprocess.run(
                ["python3", CONFIDENTIAL_SCRIPT, output_excel],
                check=True
            )
    
            print("\n[OK] Script confidencial ejecutado correctamente.")
        else:
            print("\n[ERROR] El Excel no se ha generado. No se ejecutará el script confidencial.")
    
    else:
        print("\n[INFO] No se generaron QWEN.txt en esta ejecución.")


if __name__ == "__main__":
    main()
