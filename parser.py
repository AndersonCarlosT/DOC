from docx import Document
import re

def limpiar_texto(texto):
    # Elimina tabulaciones y múltiples espacios
    texto = texto.replace('\t', ' ')
    texto = re.sub(r'[ ]{2,}', ' ', texto)
    texto = re.sub(r'\n+', '\n', texto)
    
    # Unir líneas “colgadas” que probablemente están partidas
    lineas = texto.split('\n')
    nuevas_lineas = []
    buffer = ""
    for i, linea in enumerate(lineas):
        stripped = linea.strip()
        if len(stripped) < 60 and not stripped.endswith('.') and not stripped.endswith(':'):
            # Si la línea parece colgada y no está vacía, unirla con la anterior
            if nuevas_lineas:
                nuevas_lineas[-1] += ' ' + stripped
            else:
                buffer = stripped
        else:
            nuevas_lineas.append(stripped)
    return "\n".join(nuevas_lineas).strip()

def extraer_observaciones(docx_file):
    doc = Document(docx_file)

    observaciones = []
    current_obs = None
    buffer = ""

    for para in doc.paragraphs:
        text = para.text.strip()

        # Detectar título en negrita que contiene "OBSERVACIÓN"
        if any(run.bold and "OBSERVACIÓN" in run.text for run in para.runs):
            if current_obs:  # guardar la observación anterior
                current_obs["Texto"] = buffer.strip()
                observaciones.append(current_obs)

            # Capturar número y título completo (puede estar dividido en varias líneas)
            numero = ""
            titulo = ""

            # Unir los runs en negrita para formar el título
            full_title = "".join([run.text for run in para.runs if run.bold])
            match = re.match(r"(\d+)\.\s+OBSERVACIÓN\s+(\d{2,3}):\s*(.*)", full_title)
            if match:
                numero = match.group(2)
                titulo = match.group(3)
            else:
                numero = ""
                titulo = full_title

            current_obs = {
                "N° Obs": numero,
                "Título": titulo
            }
            buffer = ""
        else:
            buffer += "\n" + text

    # Guardar última observación
    if current_obs:
        current_obs["Texto"] = buffer.strip()
        observaciones.append(current_obs)

    # Separar texto en secciones
    resultado = []
    for obs in observaciones:
        texto = limpiar_texto(obs["Texto"])

        partes = re.split(r"\n*Sustento\n*", texto, maxsplit=1, flags=re.IGNORECASE)
        obs_text = partes[0].strip()
        sustento_text = ""
        solicitud_text = ""

        if len(partes) > 1:
            partes2 = re.split(r"\n*Solicitud\n*", partes[1], maxsplit=1, flags=re.IGNORECASE)
            sustento_text = partes2[0].strip()
            if len(partes2) > 1:
                solicitud_text = partes2[1].strip()

        resultado.append({
            "N° Obs": obs["N° Obs"],
            "Título": obs["Título"],
            "Observación": obs_text,
            "Sustento": sustento_text,
            "Solicitud": solicitud_text
        })

    return resultado
