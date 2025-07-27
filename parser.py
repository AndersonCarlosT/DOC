from docx import Document
import re

def extraer_observaciones(docx_file):
    doc = Document(docx_file)
    
    # Obtenemos todos los párrafos no vacíos
    parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    texto_completo = "\n".join(parrafos)

    # Separar bloques por observaciones
    bloques = re.split(r"(?=\d+\.\s+OBSERVACIÓN\s+\d{2,3}:)", texto_completo)

    observaciones = []

    for bloque in bloques:
        bloque = bloque.strip()
        if not bloque:
            continue

        # Encontrar el encabezado (primera parte hasta la primera línea vacía)
        lineas = bloque.split("\n")
        encabezado = []
        cuerpo_inicio = 0

        for i, linea in enumerate(lineas):
            if linea.strip() == "":
                cuerpo_inicio = i + 1
                break
            encabezado.append(linea.strip())

        encabezado_texto = " ".join(encabezado).strip()

        # Extraer número y título desde encabezado
        match = re.match(r"(\d+)\.\s+OBSERVACIÓN\s+(\d{2,3}):\s+(.*)", encabezado_texto)
        if not match:
            continue

        num_obs = match.group(2).strip()
        titulo = match.group(3).strip()

        # Unir el cuerpo desde la línea siguiente
        cuerpo = "\n".join(lineas[cuerpo_inicio:]).strip()

        # Dividir en secciones
        obs_text = sustento_text = solicitud_text = ""

        partes = re.split(r"\n*Sustento\n*", cuerpo, maxsplit=1, flags=re.IGNORECASE)
        obs_text = partes[0].strip()
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
