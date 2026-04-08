import os
import torch
from threading import Thread
from transformers import (
    Qwen3VLForConditionalGeneration,
    AutoProcessor,
    TextIteratorStreamer,
    BitsAndBytesConfig
)
from qwen_vl_utils import process_vision_info



# ============================================================
# CONFIGURACIÓN DEL MODELO
# ============================================================

# Cambia esto a True si quieres usar el modelo de 32B
USE_32B = True   # TRUE = MODELO GRANDE (32B) → REQUIERE 2 GPUs Y MEMORIA EXPANDIBLE

# ============================================================
# CONFIGURACIÓN DE GPU Y MEMORIA
# ============================================================

if USE_32B:
    # Modelo grande → requiere 2 GPUs y memoria expandible
    os.environ["CUDA_VISIBLE_DEVICES"] = "0,1"
    os.environ["PYTORCH_CUDA_ALLOC_CONF"] = "expandable_segments:True"
else:
    # Modelo pequeño → 1 GPU
    os.environ["CUDA_VISIBLE_DEVICES"] = "0"

# ============================================================
# HIPERPARÁMETROS DE GENERACIÓN
# ============================================================

os.environ["GREEDY"] = "false"
os.environ["TOP_P"] = "0.8"
os.environ["TOP_K"] = "20"
os.environ["TEMPERATURE"] = "0.7"
os.environ["REPETITION_PENALTY"] = "1.0"
os.environ["PRESENCE_PENALTY"] = "1.5"
os.environ["OUT_SEQ_LENGTH"] = "16384"


# ============================================================
# CARGA DEL MODELO
# ============================================================

if USE_32B:
    MODEL_PATH = "/media/raid/santiagojn/downloaded_models/Qwen3-VL-32B-Instruct"

    model = Qwen3VLForConditionalGeneration.from_pretrained(
        MODEL_PATH,
        torch_dtype=torch.float16,
        device_map="auto",
        max_memory={
            0: "26GiB",
            1: "46GiB",
            "cpu": "100GiB"
        }
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

# ============================================================
# PROCESSOR
# ============================================================

processor = AutoProcessor.from_pretrained(MODEL_PATH)


# ============================================================
# FUNCIÓN PRINCIPAL DE ANÁLISIS
# ============================================================

def analyze_video_with_qwen(video_path: str, prompt: str) -> str:
    """
    Procesa un vídeo con Qwen y devuelve el texto generado.
    """

    messages = [
        {
            "role": "user",
            "content": [
                {"type": "video", "video": video_path},
                {"type": "text", "text": prompt},
            ],
        }
    ]

    inputs = processor.apply_chat_template(
        messages,
        tokenize=True,
        add_generation_prompt=True,
        return_dict=True,
        return_tensors="pt",
    ).to(model.device)

    streamer = TextIteratorStreamer(
        processor.tokenizer,
        skip_prompt=True,
        skip_special_tokens=True,
    )

    generation_kwargs = dict(
        **inputs,
        max_new_tokens=1500,
        do_sample=True,
        streamer=streamer,
    )

    thread = Thread(target=model.generate, kwargs=generation_kwargs)
    thread.start()

    output_text = ""
    for new_text in streamer:
        output_text += new_text

    thread.join()
    return output_text


# _MODEL = None
# _PROCESSOR = None

# def _load_qwen_once(model_path: str):
#     """
#     Loads the Qwen model and processor only once.
#     """
#     global _MODEL, _PROCESSOR

#     if _MODEL is not None and _PROCESSOR is not None:
#         return _MODEL, _PROCESSOR

#     model = Qwen3VLForConditionalGeneration.from_pretrained(
#         model_path,
#         torch_dtype=torch.float16,
#         device_map="auto",
#     )
#     model.eval()

#     processor = AutoProcessor.from_pretrained(model_path)

#     _MODEL, _PROCESSOR = model, processor
#     return _MODEL, _PROCESSOR



# def analyze_video_with_qwen(
#     # Sección ajustes modelo Qwen 8B
#     video_path: str,
#     original_url: str,
#     model_path: str = "/media/raid/santiagojn/downloaded_models/Qwen3-VL-8B-Instruct",
#     max_new_tokens: int = 1500,
#     # Sección ajustes modelo Qwen 32B
    
# ):
#     """
#     Analyzes a cooking video using Qwen and generates a .txt file containing:
#         - Title of the recipe (interpreted by Qwen)
#         - Ingredients
#         - Steps

#     Returns:
#         (True, qwen_txt_path) on success
#         (False, error_message) on failure
#     """

#     if not os.path.exists(video_path):
#         return False, f"Video file not found: {video_path}"

#     try:
#         model, processor = _load_qwen_once(model_path)

#         prompt_text = (
#             "Watch the video and extract the recipe information. "
#             "Return the following sections clearly:\n\n"
#             "TITLE:\n"
#             "INGREDIENTS:\n"
#             "STEPS:\n"
#         )

#         messages = [
#             {
#                 "role": "user",
#                 "content": [
#                     {"type": "video", "video": video_path},
#                     {"type": "text", "text": prompt_text},
#                 ],
#             }
#         ]

#         inputs = processor.apply_chat_template(
#             messages,
#             tokenize=True,
#             add_generation_prompt=True,
#             return_dict=True,
#             return_tensors="pt",
#         )
#         inputs = inputs.to(model.device)

#         streamer = TextIteratorStreamer(
#             processor.tokenizer,
#             skip_prompt=True,
#             skip_special_tokens=True,
#         )

#         generation_kwargs = dict(
#             **inputs,
#             max_new_tokens=max_new_tokens,
#             do_sample=True,
#             streamer=streamer,
#         )

#         # Output file path
#         base, _ = os.path.splitext(video_path)
#         qwen_txt_path = base + "_QWEN.txt"

#         # Write header
#         with open(qwen_txt_path, "w", encoding="utf-8") as f:
#             f.write(f"URL: {original_url}\n")
#             f.write("\nQWEN_ANALYSIS:\n")
#             f.write("-" * 60 + "\n")

#         # Run generation
#         generation_thread = Thread(
#             target=model.generate,
#             kwargs=generation_kwargs,
#             daemon=True,
#         )
#         generation_thread.start()

#         # Stream output to file
#         with open(qwen_txt_path, "a", encoding="utf-8") as f:
#             for new_text in streamer:
#                 f.write(new_text)

#         generation_thread.join()

#         return True, qwen_txt_path

#     except Exception as e:
#         return False, str(e)














