# Copyright 2024 Alibaba Group Holding Limited.
# Licensed under the Apache License, Version 2.0

import base64
from PIL import Image
import io


def load_image(image):
    if isinstance(image, str):
        return Image.open(image).convert("RGB")
    return image


def encode_image(image):
    buffered = io.BytesIO()
    image.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode("utf-8")


def process_vision_info(messages):
    """
    Convierte imágenes y vídeos en el formato que Qwen-VL espera.
    """
    images = []
    videos = []

    for msg in messages:
        if "content" not in msg:
            continue

        for item in msg["content"]:
            if item["type"] == "image":
                img = load_image(item["image"])
                images.append(encode_image(img))

            elif item["type"] == "video":
                # Qwen-VL espera la ruta del vídeo directamente
                videos.append(item["video"])

    return {"images": images, "videos": videos}
