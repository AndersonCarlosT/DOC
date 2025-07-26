from docx import Document
import re

def extraer_observaciones(docx_file):
    doc = Document(docx_file)
    full_text = "\n".join([para.text for para in doc.paragraphs if para.text.strip() != ""])

    # Separar por observaciones (usamos regex para capturar el número y título)
    bloques = re.split(r"\n?(\d+\.\s+OBSERVACIÓN\s+\d{2,3}:.+?)\n", full_text)

    observaciones = []

    # El primer elemento puede ser texto antes de la primera observación, lo ignoramos
    for i in range(1, len(bloques), 2):
        titulo_completo = bloques[i].strip()
        contenido = bloques[i+1].strip()

        # Extraer número y título
        match = re.match(r"(\d+)\.\s+OBSERVACIÓN\s+(\d{2,3}):\s+(.*)", titulo_completo)
        if not match:
            continue
        num_obs = match.group(2)
        titulo = match.group(3).strip()

        # Separar en secciones
        partes = re.split(r"\n*Sustento\n*", contenido, maxsplit=1, flags=re.IGNORECASE)
        obs_text = partes[0].strip()
        sustento_text = ""
        solicitud_text = ""

        if len(partes) > 1:
            partes2 = re.split(r"\n*Solicitud\n*", partes[1], maxsplit=1, flags=re.IGNORECASE)
            sustento_text = partes2[0].strip()
            if len(partes2) > 1:
                solicitud_text = partes2[1].strip()

        observaciones.append({
            "N° Obs": num_obs,
            "Título": titulo,
            "Observación": obs_text,
            "Sustento": sustento_text,
            "Solicitud": solicitud_text
        })

    return observaciones
