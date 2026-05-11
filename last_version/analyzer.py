import os
import torch
from threading import Thread
from transformers import TextIteratorStreamer

# ============================================================
# PROMPT ESTRICTO
# ============================================================

STRICT_PROMPT = """
Quiero que analices el vídeo y devuelvas la información en este formato:

TÍTULO:
(título deducido o vacío)

INGREDIENTES:
- ingrediente 1
- ingrediente 2
- ...

PASOS:
1. acción 1
2. acción 2
3. acción 3
...

Reglas:
- Sigue el orden del vídeo.
- Cada acción debe ser un paso independiente.
- No inventes información que no aparezca claramente.
- No añadas explicaciones fuera del formato.
- Si no estás seguro de algo, déjalo vacío.
""".strip()


# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

@torch.inference_mode()
def analyze_video_with_qwen(video_path: str,
                            original_url: str,
                            model,
                            processor,
                            gen_kwargs) -> str:
    """
    Procesa un vídeo con Qwen usando el modelo cargado en main.py.
    Devuelve texto limpio para el Excel filler.
    """

    # Construcción del mensaje
    messages = [
        {
            "role": "system",
            "content": STRICT_PROMPT
        },
        {
            "role": "user",
            "content": [
                {"type": "video", "video": video_path}
            ]
        }
    ]

    # Preparamos inputs
    inputs = processor.apply_chat_template(
        messages,
        tokenize=True,
        add_generation_prompt=True,
        return_dict=True,
        return_tensors="pt"
    ).to(model.device)

    # Streamer para generación progresiva
    streamer = TextIteratorStreamer(
        processor.tokenizer,
        skip_prompt=True,
        skip_special_tokens=True,
    )

    # Mezclamos inputs + kwargs de generación
    generation_kwargs = dict(
        **inputs,
        streamer=streamer,
        **gen_kwargs
    )

    # Lanzamos generación en un hilo
    thread = Thread(target=model.generate, kwargs=generation_kwargs)
    thread.start()

    output_text = ""
    for new_text in streamer:
        output_text += new_text

    thread.join()

    # ============================================================
    # LIMPIEZA PARA EL EXCEL FILLER
    # ============================================================

    cleaned = output_text

    # 1. Recortar todo lo que esté antes de "TÍTULO:"
    if "TÍTULO:" in cleaned:
        cleaned = cleaned[cleaned.index("TÍTULO:"):]

    # 2. Recortar después de PASOS
    import re
    match = re.search(r"(PASOS:\s*(?:\d+\..*\n?)*)", cleaned, re.DOTALL)
    if match:
        cleaned = cleaned[:match.end()]

    # 3. Añadir URL al principio
    cleaned = f"URL: {original_url}\n\n" + cleaned

    return cleaned.strip()


# import os
# import torch
# from threading import Thread
# from transformers import (
#     Qwen3VLForConditionalGeneration,
#     AutoProcessor,
#     TextIteratorStreamer,
#     BitsAndBytesConfig
# )



# # ============================================================
# # CONFIGURACIÓN DEL MODELO
# # ============================================================

# # Cambia esto a True si quieres usar el modelo de 32B
# USE_32B = True   # TRUE = MODELO GRANDE (32B) → REQUIERE 2 GPUs Y MEMORIA EXPANDIBLE

# # ============================================================
# # CONFIGURACIÓN DE GPU Y MEMORIA
# # ============================================================

# if USE_32B:
#     # Modelo grande → requiere 2 GPUs y memoria expandible
#     os.environ["CUDA_VISIBLE_DEVICES"] = "0,1"
#     os.environ["PYTORCH_CUDA_ALLOC_CONF"] = "expandable_segments:True"
# else:
#     # Modelo pequeño → 1 GPU
#     os.environ["CUDA_VISIBLE_DEVICES"] = "0"

# # ============================================================
# # HIPERPARÁMETROS DE GENERACIÓN
# # ============================================================

# os.environ["GREEDY"] = "false"
# os.environ["TOP_P"] = "0.8"
# os.environ["TOP_K"] = "20"
# os.environ["TEMPERATURE"] = "0.7"
# os.environ["REPETITION_PENALTY"] = "1.0"
# os.environ["PRESENCE_PENALTY"] = "1.5"
# os.environ["OUT_SEQ_LENGTH"] = "16384"


# # ============================================================
# # CARGA DEL MODELO
# # ============================================================

# if USE_32B:
#     MODEL_PATH = "/media/raid/santiagojn/downloaded_models/Qwen3-VL-32B-Instruct"

#     model = Qwen3VLForConditionalGeneration.from_pretrained(
#         MODEL_PATH,
#         torch_dtype=torch.float16,
#         device_map="auto",
#         max_memory={
#             0: "26GiB",
#             1: "46GiB",
#             "cpu": "100GiB"
#         }
#     )

# else:
#     MODEL_PATH = "/media/raid/santiagojn/downloaded_models/Qwen3-VL-8B-Instruct"

#     model = Qwen3VLForConditionalGeneration.from_pretrained(
#         MODEL_PATH,
#         torch_dtype=torch.float16,
#         device_map="auto",
#         local_files_only=True
#     )

# model.eval()
# torch.set_grad_enabled(False)

# # ============================================================
# # PROCESSOR
# # ============================================================

# processor = AutoProcessor.from_pretrained(MODEL_PATH)

# RECIPE_PROMPT = """
# Analiza el vídeo y devuelve ÚNICAMENTE el siguiente formato, sin texto adicional antes o después:

# TÍTULO:
# (título deducido o vacío)

# INGREDIENTES:
# - ingrediente 1
# - ingrediente 2
# - ...

# PASOS:
# 1. acción 1
# 2. acción 2
# 3. acción 3
# ...

# Instrucciones estrictas:
# - Cada acción debe ser un paso independiente.
# - No agrupes acciones.
# - Mantén el orden cronológico exacto del vídeo.
# - Si se menciona un tiempo de cocción, inclúyelo dentro del paso correspondiente.
# - No añadas comentarios, explicaciones ni texto fuera de las secciones indicadas.
# - No inventes información que no se vea o mencione claramente.
# """

# # ============================================================
# # FUNCIÓN PRINCIPAL DE ANÁLISIS
# # ============================================================
# # def analyze_video_with_qwen(video_path: str, prompt: str, original_url: str) -> str:
# def analyze_video_with_qwen(video_path: str, original_url: str) -> str:
#     """
#     Procesa un vídeo con Qwen y devuelve el texto generado,
#     limpiado para que el Excel filler pueda procesarlo sin errores.
#     """

#     messages = [
#         {
#             "role": "user",
#             "content": [
#                 {"type": "video", "video": video_path},
#                 {"type": "text", "text": RECIPE_PROMPT},
#             ],
#         }
#     ]

#     inputs = processor.apply_chat_template(
#         messages,
#         tokenize=True,
#         add_generation_prompt=True,
#         return_dict=True,
#         return_tensors="pt",
#     ).to(model.device)

#     streamer = TextIteratorStreamer(
#         processor.tokenizer,
#         skip_prompt=True,
#         skip_special_tokens=True,
#     )

#     generation_kwargs = dict(
#         **inputs,
#         max_new_tokens=1500,
#         do_sample=True,
#         streamer=streamer,
#     )

#     thread = Thread(target=model.generate, kwargs=generation_kwargs)
#     thread.start()

#     output_text = ""
#     for new_text in streamer:
#         output_text += new_text

#     thread.join()

#     # ============================================================
#     # LIMPIEZA DEL TEXTO PARA QUE EL EXCEL FILLER NO FALLE
#     # ============================================================

#     cleaned = output_text

#     # 1. Recortar todo lo que esté antes de "TÍTULO:"
#     if "TÍTULO:" in cleaned:
#         cleaned = cleaned[cleaned.index("TÍTULO:"):]

#     # 2. Recortar todo lo que esté después de la sección PASOS
#     import re
#     match = re.search(r"(PASOS:\s*(?:\d+\..*\n?)*)", cleaned, re.DOTALL)
#     if match:
#         cleaned = cleaned[:match.end()]

#     # 3. Añadir la URL al principio para que el Excel filler siga funcionando
#     cleaned = f"URL: {original_url}\n\n" + cleaned

#     return cleaned.strip()



# # def analyze_video_with_qwen(video_path: str, prompt: str) -> str:
# #     """
# #     Procesa un vídeo con Qwen y devuelve el texto generado.
# #     """

# #     messages = [
# #         {
# #             "role": "user",
# #             "content": [
# #                 {"type": "video", "video": video_path},
# #                 {"type": "text", "text": prompt},
# #             ],
# #         }
# #     ]

# #     inputs = processor.apply_chat_template(
# #         messages,
# #         tokenize=True,
# #         add_generation_prompt=True,
# #         return_dict=True,
# #         return_tensors="pt",
# #     ).to(model.device)

# #     streamer = TextIteratorStreamer(
# #         processor.tokenizer,
# #         skip_prompt=True,
# #         skip_special_tokens=True,
# #     )

# #     generation_kwargs = dict(
# #         **inputs,
# #         max_new_tokens=1500,
# #         do_sample=True,
# #         streamer=streamer,
# #     )

# #     thread = Thread(target=model.generate, kwargs=generation_kwargs)
# #     thread.start()

# #     output_text = ""
# #     for new_text in streamer:
# #         output_text += new_text

# #     thread.join()
# #     return output_text


# # # _MODEL = None
# # # _PROCESSOR = None

# # # def _load_qwen_once(model_path: str):
# # #     """
# # #     Loads the Qwen model and processor only once.
# # #     """
# # #     global _MODEL, _PROCESSOR

# # #     if _MODEL is not None and _PROCESSOR is not None:
# # #         return _MODEL, _PROCESSOR

# # #     model = Qwen3VLForConditionalGeneration.from_pretrained(
# # #         model_path,
# # #         torch_dtype=torch.float16,
# # #         device_map="auto",
# # #     )
# # #     model.eval()

# # #     processor = AutoProcessor.from_pretrained(model_path)

# # #     _MODEL, _PROCESSOR = model, processor
# # #     return _MODEL, _PROCESSOR



# # # def analyze_video_with_qwen(
# # #     # Sección ajustes modelo Qwen 8B
# # #     video_path: str,
# # #     original_url: str,
# # #     model_path: str = "/media/raid/santiagojn/downloaded_models/Qwen3-VL-8B-Instruct",
# # #     max_new_tokens: int = 1500,
# # #     # Sección ajustes modelo Qwen 32B
    
# # # ):
# # #     """
# # #     Analyzes a cooking video using Qwen and generates a .txt file containing:
# # #         - Title of the recipe (interpreted by Qwen)
# # #         - Ingredients
# # #         - Steps

# # #     Returns:
# # #         (True, qwen_txt_path) on success
# # #         (False, error_message) on failure
# # #     """

# # #     if not os.path.exists(video_path):
# # #         return False, f"Video file not found: {video_path}"

# # #     try:
# # #         model, processor = _load_qwen_once(model_path)

# # #         prompt_text = (
# # #             "Watch the video and extract the recipe information. "
# # #             "Return the following sections clearly:\n\n"
# # #             "TITLE:\n"
# # #             "INGREDIENTS:\n"
# # #             "STEPS:\n"
# # #         )

# # #         messages = [
# # #             {
# # #                 "role": "user",
# # #                 "content": [
# # #                     {"type": "video", "video": video_path},
# # #                     {"type": "text", "text": prompt_text},
# # #                 ],
# # #             }
# # #         ]

# # #         inputs = processor.apply_chat_template(
# # #             messages,
# # #             tokenize=True,
# # #             add_generation_prompt=True,
# # #             return_dict=True,
# # #             return_tensors="pt",
# # #         )
# # #         inputs = inputs.to(model.device)

# # #         streamer = TextIteratorStreamer(
# # #             processor.tokenizer,
# # #             skip_prompt=True,
# # #             skip_special_tokens=True,
# # #         )

# # #         generation_kwargs = dict(
# # #             **inputs,
# # #             max_new_tokens=max_new_tokens,
# # #             do_sample=True,
# # #             streamer=streamer,
# # #         )

# # #         # Output file path
# # #         base, _ = os.path.splitext(video_path)
# # #         qwen_txt_path = base + "_QWEN.txt"

# # #         # Write header
# # #         with open(qwen_txt_path, "w", encoding="utf-8") as f:
# # #             f.write(f"URL: {original_url}\n")
# # #             f.write("\nQWEN_ANALYSIS:\n")
# # #             f.write("-" * 60 + "\n")

# # #         # Run generation
# # #         generation_thread = Thread(
# # #             target=model.generate,
# # #             kwargs=generation_kwargs,
# # #             daemon=True,
# # #         )
# # #         generation_thread.start()

# # #         # Stream output to file
# # #         with open(qwen_txt_path, "a", encoding="utf-8") as f:
# # #             for new_text in streamer:
# # #                 f.write(new_text)

# # #         generation_thread.join()

# # #         return True, qwen_txt_path

# # #     except Exception as e:
# # #         return False, str(e)














