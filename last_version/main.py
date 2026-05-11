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

from excel_filler import generate_excel_from_multiple_txt

from transformers import (
    Qwen3VLForConditionalGeneration,
    AutoProcessor,
    TextIteratorStreamer,
    BitsAndBytesConfig
)
import torch

# ============================================================
# CONFIGURACIÓN DEL MODELO
# ============================================================

USE_32B = True

if USE_32B:
    os.environ["CUDA_VISIBLE_DEVICES"] = "0,1"
    os.environ["PYTORCH_CUDA_ALLOC_CONF"] = "expandable_segments:True"
else:
    os.environ["CUDA_VISIBLE_DEVICES"] = "0"

# ============================================================
# HIPERPARÁMETROS DE GENERACIÓN
# ============================================================

# Forzamos greedy decoding para máxima obediencia
GEN_KWARGS = dict(
    max_new_tokens=1500,
    do_sample=False,
    top_p=1.0,
    repetition_penalty=1.0
)

# ============================================================
# CARGA DEL MODELO
# ============================================================

if USE_32B:
    MODEL_PATH = "/media/raid/santiagojn/downloaded_models/Qwen3-VL-32B-Instruct"

    model = Qwen3VLForConditionalGeneration.from_pretrained(
        MODEL_PATH,
        torch_dtype=torch.float16,
        device_map="auto",
        max_memory={0: "26GiB", 1: "46GiB", "cpu": "100GiB"}
    )
else:
    MODEL_PATH = "/media/raid/santiagojn/downloaded_models/Qwen3-VL-8B-Instruct"

    model = Qwen3VLForConditionalGeneration.from_pretrained(
        MODEL_PATH,
        torch_dtype=torch.float16,
        device_map="auto",
        local_files_only=True
    )

model.eval()
torch.set_grad_enabled(False)

processor = AutoProcessor.from_pretrained(MODEL_PATH)

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

            # prompt = """ Analiza el vídeo y devuelve sin texto adicional antes o después los ingredientes y los pasos de una receta. 
            # Para cada acción realizada, debe ser un paso independiente, sin agrupar acciones. Mantén el orden cronológico exacto del vídeo. Si se menciona un tiempo de cocción, inclúyelo dentro del paso correspondiente.
            # No inventes información que no se vea o mencione claramente en el vídeo. """


            generated_text = analyze_video_with_qwen(mp4_path,url,model,processor,GEN_KWARGS)


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
    # 4. Generar Excel final
    # ---------------------------------------------------------
    if generated_qwen_files:
        print("\n=== GENERATING FINAL EXCEL ===")
        TEMPLATE_PATH = "plantilla.xlsx" 
    
        output_excel = generate_excel_from_multiple_txt(
            template_xlsx_path=TEMPLATE_PATH,
            qwen_txt_paths=generated_qwen_files,
            model=model,
            processor=processor,
            gen_kwargs=GEN_KWARGS, 
            url_list=urls
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








# import os
# from utils import (
#     read_urls,
#     ensure_folder,
#     ensure_reports_folder
# )
# from downloader import download_instagram_video
#     # returns: (ok, mp4_path_or_error, meta_txt_path)

# from analyzer import analyze_video_with_qwen
#     # returns: texto generado por Qwen

# from reporter import write_report

# from excel_filler import generate_excel_from_multiple_txt


# def main():

#     # ---------------------------------------------------------
#     # 0. Preparación de carpetas
#     # ---------------------------------------------------------
#     ensure_folder("recipes_videos")
#     reports_dir = ensure_reports_folder()

#     # ---------------------------------------------------------
#     # 1. Leer URLs
#     # ---------------------------------------------------------
#     urls = read_urls()
#     report_entries = []

#     # Lista donde guardamos los QWEN.txt generados EN ORDEN
#     generated_qwen_files = []

#     # ---------------------------------------------------------
#     # 2. Procesar cada URL en orden
#     # ---------------------------------------------------------
#     for url in urls:
#         print(f"\nProcessing: {url}")

#         # 2.1 Descargar vídeo
#         ok, mp4_or_error, meta_txt = download_instagram_video(url)
#         if not ok:
#             report_entries.append(f"[ERROR] {url} → Download failed: {mp4_or_error}")
#             continue

#         mp4_path = mp4_or_error

#         # 2.2 Analizar vídeo con Qwen
#         try:

#             # prompt = """ Analiza el vídeo y devuelve sin texto adicional antes o después los ingredientes y los pasos de una receta. 
#             # Para cada acción realizada, debe ser un paso independiente, sin agrupar acciones. Mantén el orden cronológico exacto del vídeo. Si se menciona un tiempo de cocción, inclúyelo dentro del paso correspondiente.
#             # No inventes información que no se vea o mencione claramente en el vídeo. """

# #             prompt = """
# #                     Analiza el vídeo y devuelve ÚNICAMENTE el siguiente formato, sin texto adicional antes o después:
                    

# #                     TÍTULO:
# #                     (título deducido o vacío)

# #                     INGREDIENTES:
# #                     - ingrediente 1
# #                     - ingrediente 2
# #                     - ...

# #                     PASOS:
# #                     1. acción 1
# #                     2. acción 2
# #                     3. acción 3
# #                     ...

# #                     Instrucciones estrictas:
# #                     - Cada acción debe ser un paso independiente.
# #                     - No agrupes acciones.
# #                     - Mantén el orden cronológico exacto del vídeo.
# #                     - Si se menciona un tiempo de cocción, inclúyelo dentro del paso correspondiente.
# #                     - No añadas comentarios, explicaciones ni texto fuera de las secciones indicadas.
# #                     - No inventes información que no se vea o mencione claramente.
# # """


#             # generated_text = analyze_video_with_qwen(mp4_path, prompt, url)
#             generated_text = analyze_video_with_qwen(
#                 video_path=mp4_path,
#                 original_url=url
#             )

#         except Exception as e:
#             report_entries.append(f"[ERROR] {url} → Qwen failed: {str(e)}")
#             continue

#         # Guardar el texto generado en un archivo .txt
#         qwen_txt_path = mp4_path.replace(".mp4", "_QWEN.txt")
#         with open(qwen_txt_path, "w", encoding="utf-8") as f:
#             f.write(generated_text)

#         # Guardamos el QWEN.txt en orden EXACTO
#         generated_qwen_files.append(qwen_txt_path)

#         report_entries.append(f"[OK] {url} → {qwen_txt_path}")

#     # ---------------------------------------------------------
#     # 3. Escribir summary/report
#     # ---------------------------------------------------------
#     write_report(report_entries, reports_dir)

#     # ---------------------------------------------------------
#     # 4. Generar Excel final
#     # ---------------------------------------------------------
#     if generated_qwen_files:
#         print("\n=== GENERATING FINAL EXCEL ===")
#         TEMPLATE_PATH = "plantilla.xlsx" 
    
#         output_excel = generate_excel_from_multiple_txt(
#             template_xlsx_path=TEMPLATE_PATH,
#             qwen_txt_paths=generated_qwen_files
#         )
    
#         print(f"\n[OK] Excel final generado → {output_excel}")
    
#         # ---------------------------------------------------------
#         # 5. EJECUTAR SCRIPT CONFIDENCIAL SOLO SI EL EXCEL EXISTE
#         # ---------------------------------------------------------
#         import subprocess
    
#         if os.path.exists(output_excel):
#             print("\n=== EJECUTANDO SCRIPT CONFIDENCIAL ===")
    
#             CONFIDENTIAL_SCRIPT = "ejecutar_BSH_automatico.py"
    
#             subprocess.run(
#                 ["python3", CONFIDENTIAL_SCRIPT, output_excel],
#                 check=True
#             )
    
#             print("\n[OK] Script confidencial ejecutado correctamente.")
#         else:
#             print("\n[ERROR] El Excel no se ha generado. No se ejecutará el script confidencial.")
    
#     else:
#         print("\n[INFO] No se generaron QWEN.txt en esta ejecución.")


# if __name__ == "__main__":
#     main()
