from docx import Document
import re

def extraer_observaciones(docx_file):
    doc = Document(docx_file)
    full_text = "\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip() != ""])

    # Unimos todo el texto como bloques separados por observaciones
    bloques = re.split(r"(?=\d+\.\s+OBSERVACIÓN\s+\d{2,3}:)", full_text)

    observaciones = []

    for bloque in bloques:
        bloque = bloque.strip()
        if not bloque:
            continue

        # Separar encabezado (número + título) del contenido
        match = re.match(r"(\d+)\.\s+OBSERVACIÓN\s+(\d{2,3}):(.*?)(?=\n\n|Sustento|Solicitud)", bloque, flags=re.DOTALL)
        if not match:
            continue

        num_obs = match.group(2).strip()
        titulo = match.group(3).replace("\n", " ").strip()  # Unimos líneas del título

        # Remover encabezado del bloque para quedarnos con el resto
        resto = bloque[match.end():].strip()

        # Extraer secciones
        partes = re.split(r"\n*Sustento\n*", resto, maxsplit=1, flags=re.IGNORECASE)
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
