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
