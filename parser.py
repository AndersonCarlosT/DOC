from docx import Document
import re

def extraer_observaciones(docx_file):
    doc = Document(docx_file)
    full_text = "\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip() != ""])

    # Separamos por observación
    bloques = re.split(r"(?=\d+\.\s+OBSERVACIÓN\s+\d{2,3}:)", full_text)

    observaciones = []

    for bloque in bloques:
        bloque = bloque.strip()
        if not bloque:
            continue

        # Dividimos en dos partes: encabezado (hasta primer doble salto) y el resto
        partes = re.split(r"\n{2,}", bloque, maxsplit=1)
        encabezado = partes[0].replace("\n", " ").strip()
        cuerpo = partes[1].strip() if len(partes) > 1 else ""

        # Extraer número y título desde encabezado
        match = re.match(r"(\d+)\.\s+OBSERVACIÓN\s+(\d{2,3}):(.*)", encabezado)
        if not match:
            continue

        num_obs = match.group(2).strip()
        titulo = match.group(3).strip()

        # Separar cuerpo en secciones: Observación / Sustento / Solicitud
        obs_text, sustento_text, solicitud_text = "", "", ""

        partes_cuerpo = re.split(r"\n*Sustento\n*", cuerpo, maxsplit=1, flags=re.IGNORECASE)
        obs_text = partes_cuerpo[0].strip()

        if len(partes_cuerpo) > 1:
            partes_solicitud = re.split(r"\n*Solicitud\n*", partes_cuerpo[1], maxsplit=1, flags=re.IGNORECASE)
            sustento_text = partes_solicitud[0].strip()
            if len(partes_solicitud) > 1:
                solicitud_text = partes_solicitud[1].strip()

        observaciones.append({
            "N° Obs": num_obs,
            "Título": titulo,
            "Observación": obs_text,
            "Sustento": sustento_text,
            "Solicitud": solicitud_text
        })

    return observaciones
