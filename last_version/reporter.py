import os
from datetime import datetime

def write_report(entries, reports_dir):
    # Crear nombre único para el archivo
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(reports_dir, f"report_{timestamp}.txt")

    # Escribir el archivo
    with open(output_path, "w", encoding="utf-8") as f:
        for line in entries:
            f.write(line + "\n")

    return output_path
