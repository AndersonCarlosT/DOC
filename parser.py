from docx import Document
import re

def limpiar_texto(texto):
    # Reemplaza espacio + tabulación por un espacio
    texto = re.sub(r"[ \t]+", " ", texto)
    # Reemplaza múltiples espacios por uno
    texto = re.sub(r" {2,}", " ", texto)
    return texto.strip()

def extraer_observaciones(docx_file):
    doc = Document(docx_file)
    full_text = "\n".join([limpiar_texto(para.text) for para in doc.paragraphs if para.text.strip() != ""])

    bloques = re.split(r"\n?(\d+\.\s+OBSERVACIÓN\s+\d{2,3}:.+?)\n", full_text)

    observaciones = []

    for i in range(1, len(bloques), 2):
        titulo_completo = bloques[i].strip()
        contenido = bloques[i+1].strip()

        match = re.match(r"(\d+)\.\s+OBSERVACIÓN\s+(\d{2,3}):\s+(.*)", titulo_completo)
        if not match:
            continue
        num_obs = match.group(2)
        titulo = limpiar_texto(match.group(3).strip())

        partes = re.split(r"\n*Sustento\n*", contenido, maxsplit=1, flags=re.IGNORECASE)
        obs_text = limpiar_texto(partes[0])
        sustento_text = ""
        solicitud_text = ""

        if len(partes) > 1:
            partes2 = re.split(r"\n*Solicitud\n*", partes[1], maxsplit=1, flags=re.IGNORECASE)
            sustento_text = limpiar_texto(partes2[0])
            if len(partes2) > 1:
                solicitud_text = limpiar_texto(partes2[1])

        observaciones.append({
            "N° Obs": num_obs,
            "Título": titulo,
            "Observación": obs_text,
            "Sustento": sustento_text,
            "Solicitud": solicitud_text
        })

    return observaciones
