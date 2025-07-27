from docx import Document
import re

def limpiar_texto(texto):
    # Reemplaza tabulaciones con espacios y limpia múltiples espacios
    texto = texto.replace('\t', ' ')
    texto = re.sub(r'[ ]{2,}', ' ', texto)
    texto = re.sub(r'\n+', '\n', texto)  # Elimina saltos de línea repetidos
    return texto.strip()

def extraer_observaciones(docx_file):
    doc = Document(docx_file)

    # Etapa 1: Construimos una lista de observaciones con sus bloques de texto
    observaciones = []
    obs_actual = {}
    recogiendo = False
    buffer_texto = ""

    for para in doc.paragraphs:
        text = para.text.strip()

        # Buscar si hay título en negrita que comienza con "1. OBSERVACIÓN XX:"
        for run in para.runs:
            if run.bold and "OBSERVACIÓN" in run.text:
                if obs_actual:
                    obs_actual["Texto"] = buffer_texto.strip()
                    observaciones.append(obs_actual)

                titulo_completo = para.text.strip()
                match = re.match(r"(\d+)\.\s+OBSERVACIÓN\s+(\d{2,3}):\s*(.*)", titulo_completo)
                if match:
                    obs_actual = {
                        "N° Obs": match.group(2),
                        "Título": match.group(3)
                    }
                else:
                    obs_actual = {"N° Obs": "", "Título": titulo_completo}
                buffer_texto = ""
                recogiendo = True
                break

        if recogiendo and not any(run.bold and "OBSERVACIÓN" in run.text for run in para.runs):
            buffer_texto += "\n" + para.text

    # Guardar la última observación
    if obs_actual:
        obs_actual["Texto"] = buffer_texto.strip()
        observaciones.append(obs_actual)

    # Etapa 2: Separar el bloque de texto en secciones
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
