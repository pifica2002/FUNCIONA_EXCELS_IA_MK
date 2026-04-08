import os
import re
import json
from datetime import datetime
from typing import Any, Dict, List

import torch
from openpyxl import load_workbook
from transformers import Qwen3VLForConditionalGeneration, AutoProcessor


# ============================================================
# Default configuration
# ============================================================

DEFAULT_MODEL_PATH = "/media/raid/santiagojn/downloaded_models/Qwen3-VL-8B-Instruct/"
DEFAULT_INSTRUCTIONS_TXT = "instrucciones.txt"
DEFAULT_OUTPUT_PREFIX = "output"

DEFAULT_HEADER_ROW = 4
DEFAULT_START_ROW = 5
DEFAULT_START_COL = 1
DEFAULT_NUM_COLS = 14
DEFAULT_CUDA_VISIBLE_DEVICES = "0"


# ============================================================
# Utility functions
# ============================================================

def read_text_file(path: str) -> str:
    """Read a UTF-8 text file."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return f.read().strip()


def extract_json_from_text(text: str) -> Any:
    """Extract valid JSON from model output."""
    code_block = re.search(r"```json\s*(.*?)\s*```", text, re.DOTALL | re.IGNORECASE)
    if code_block:
        return json.loads(code_block.group(1))

    array_match = re.search(r"(\[\s*{.*}\s*\])", text, re.DOTALL)
    if array_match:
        return json.loads(array_match.group(1))

    object_match = re.search(r"(\{\s*\".*\}\s*)", text, re.DOTALL)
    if object_match:
        return json.loads(object_match.group(1))

    first = text.find("{")
    last = text.rfind("}")
    if first != -1 and last != -1:
        return json.loads(text[first:last + 1])

    raise ValueError("No valid JSON detected in model output.")


def read_headers_A_to_N(template_path: str, header_row: int) -> List[str]:
    """Read header row (A..N) from FIRST worksheet."""
    wb = load_workbook(template_path)
    ws = wb.worksheets[0]

    headers = []
    for col in range(DEFAULT_START_COL, DEFAULT_START_COL + DEFAULT_NUM_COLS):
        v = ws.cell(row=header_row, column=col).value
        headers.append("" if v is None else str(v).strip())

    return headers


def normalize_table_output(data: Any, default_start_row: int) -> Dict[str, Any]:
    """Ensure model JSON matches expected schema."""
    if not isinstance(data, dict) or "rows" not in data:
        raise ValueError("Invalid JSON structure. Expected object with key 'rows'.")

    start_row = int(data.get("start_row", default_start_row))
    rows = data["rows"]

    normalized = []

    for r in rows:
        if not isinstance(r, list):
            raise ValueError("Each row must be a list.")

        if len(r) < DEFAULT_NUM_COLS:
            r += [""] * (DEFAULT_NUM_COLS - len(r))
        elif len(r) > DEFAULT_NUM_COLS:
            r = r[:DEFAULT_NUM_COLS]

        normalized.append(r)

    return {"start_row": start_row, "rows": normalized}


def apply_table_to_first_sheet_inplace(ws,
                                       start_row: int,
                                       rows: List[List[Any]]) -> int:
    """
    Escribe filas en la primera hoja, empezando en start_row.
    Devuelve la siguiente fila libre después de escribir.
    """
    for i, row_vals in enumerate(rows):
        for j, value in enumerate(row_vals):
            ws.cell(row=start_row + i,
                    column=DEFAULT_START_COL + j).value = value
    return start_row + len(rows)


def build_messages(instructions: str, user_input: str,
                   headers: List[str], start_row: int):

    # ⚠️ Esta es EXACTAMENTE tu versión que funcionaba
    headers_str = "\n".join(
        [f"{chr(65+i)}: {headers[i]}" for i in range(DEFAULT_NUM_COLS)]
    )

    system_prompt = (
        "You are a data extraction engine.\n"
        "Output ONLY valid JSON.\n"
        "Structure:\n"
        "{\n"
        f'  \"start_row\": {start_row},\n'
        '  \"rows\": [[\"A\",\"B\",\"C\",\"D\",\"E\",\"F\",\"G\",\"H\",\"I\",\"J\",\"K\",\"L\",\"M\",\"N\"]]\n'
        "}\n"
        "Exactly 14 columns per row.\n"
        "No explanations."
    )

    user_prompt = (
        "COLUMNS (A..N):\n"
        f"{headers_str}\n\n"
        "INSTRUCTIONS:\n"
        f"{instructions}\n\n"
        "INPUT DATA:\n"
        f"{user_input}"
    )

    return [
        {"role": "system", "content": [{"type": "text", "text": system_prompt}]},
        {"role": "user", "content": [{"type": "text", "text": user_prompt}]},
    ]


@torch.inference_mode()
def generate_response(model, processor, messages):
    """Generate model response."""
    inputs = processor.apply_chat_template(
        messages,
        tokenize=True,
        add_generation_prompt=True,
        return_dict=True,
        return_tensors="pt",
    ).to(model.device)

    output_ids = model.generate(
        **inputs,
        max_new_tokens=2048,
        do_sample=False
    )

    generated_ids = output_ids[0][inputs["input_ids"].shape[-1]:]
    return processor.tokenizer.decode(
        generated_ids,
        skip_special_tokens=True
    ).strip()


# ============================================================
# FUNCIÓN ORIGINAL (un solo TXT → un Excel)
# ============================================================

def generate_filled_excel(template_xlsx_path: str,
                          qwen_txt_path: str) -> str:
    """
    Generate a filled Excel file using:
    - instrucciones.txt
    - Qwen analysis TXT file from Step 2

    Returns:
        Path to generated Excel file.
    """

    base_dir = os.path.dirname(os.path.abspath(__file__))

    instructions_path = os.path.join(base_dir, DEFAULT_INSTRUCTIONS_TXT)

    instructions = read_text_file(instructions_path)
    user_input = read_text_file(qwen_txt_path)

    headers = read_headers_A_to_N(template_xlsx_path, DEFAULT_HEADER_ROW)

    os.environ["CUDA_VISIBLE_DEVICES"] = DEFAULT_CUDA_VISIBLE_DEVICES

    model = Qwen3VLForConditionalGeneration.from_pretrained(
        DEFAULT_MODEL_PATH,
        torch_dtype=torch.float16,
        device_map="auto",
        local_files_only=True,
    )
    model.eval()

    processor = AutoProcessor.from_pretrained(DEFAULT_MODEL_PATH)

    messages = build_messages(
        instructions,
        user_input,
        headers,
        DEFAULT_START_ROW
    )

    raw_output = generate_response(model, processor, messages)

    parsed_json = extract_json_from_text(raw_output)
    table = normalize_table_output(parsed_json, DEFAULT_START_ROW)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(
        base_dir,
        f"{DEFAULT_OUTPUT_PREFIX}_{timestamp}.xlsx"
    )

    # Aquí se abría y escribía en un solo Excel
    wb = load_workbook(template_xlsx_path)
    ws = wb.worksheets[0]

    apply_table_to_first_sheet_inplace(
        ws,
        table["start_row"],
        table["rows"],
    )

    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(output_path)

    return output_path


# ============================================================
# NUEVA FUNCIÓN: varios TXT → UN SOLO Excel
# ============================================================

def generate_excel_from_multiple_txt(template_xlsx_path: str,
                                     qwen_txt_paths: List[str]) -> str:
    """
    Igual que generate_filled_excel, pero:
    - Reutiliza el mismo modelo y processor
    - Procesa varios QWEN.txt en ORDEN
    - Escribe todas las filas en un ÚNICO Excel
    """

    base_dir = os.path.dirname(os.path.abspath(__file__))
    instructions_path = os.path.join(base_dir, DEFAULT_INSTRUCTIONS_TXT)

    instructions = read_text_file(instructions_path)
    headers = read_headers_A_to_N(template_xlsx_path, DEFAULT_HEADER_ROW)

    os.environ["CUDA_VISIBLE_DEVICES"] = DEFAULT_CUDA_VISIBLE_DEVICES

    model = Qwen3VLForConditionalGeneration.from_pretrained(
        DEFAULT_MODEL_PATH,
        torch_dtype=torch.float16,
        device_map="auto",
        local_files_only=True,
    )
    model.eval()

    processor = AutoProcessor.from_pretrained(DEFAULT_MODEL_PATH)

    # Abrimos la plantilla UNA sola vez
    wb = load_workbook(template_xlsx_path)
    ws = wb.worksheets[0]

    current_row = DEFAULT_START_ROW

    for qwen_txt_path in qwen_txt_paths:
        user_input = read_text_file(qwen_txt_path)

        messages = build_messages(
            instructions,
            user_input,
            headers,
            current_row  # ← start_row se va actualizando
        )

        raw_output = generate_response(model, processor, messages)
        parsed_json = extract_json_from_text(raw_output)
        table = normalize_table_output(parsed_json, current_row)

        current_row = apply_table_to_first_sheet_inplace(
            ws,
            table["start_row"],
            table["rows"],
        )

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(
        base_dir,
        f"{DEFAULT_OUTPUT_PREFIX}_{timestamp}.xlsx"
    )

    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(output_path)

    return output_path
