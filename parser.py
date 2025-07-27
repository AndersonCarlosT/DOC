from docx import Document
import re

def extraer_observaciones(docx_file):
    doc = Document(docx_file)
    observaciones = []
    i = 0
    paragraphs = doc.paragraphs

    while i < len(paragraphs):
        para = paragraphs[i].text.strip()

        # Buscar el inicio de una observación
        match = re.match(r"(\d+)\.\s+OBSERVACIÓN\s+(\d{2,3}):", para, re.IGNORECASE)
        if match:
            num_obs = match.group(2)

            # Extraer todo lo que esté en negrita como título
            titulo_parts = []
            while i < len(paragraphs):
                for run in paragraphs[i].runs:
                    if run.bold:
                        titulo_parts.append(run.text.strip())
                # Terminar si ya no hay más negrita (evita ir más allá del título)
                if any(run.bold for run in paragraphs[i].runs):
                    i += 1
                else:
                    break

            titulo = " ".join(titulo_parts).strip()

            # Extraer observación, sustento, solicitud
            obs_text = []
            sustento_text = []
            solicitud_text = []

            seccion = "observacion"
            while i < len(paragraphs):
                text = paragraphs[i].text.strip()

                if re.match(r"^Sustento$", text, re.IGNORECASE):
                    seccion = "sustento"
                elif re.match(r"^Solicitud$", text, re.IGNORECASE):
                    seccion = "solicitud"
                elif re.match(r"^\d+\.\s+OBSERVACIÓN\s+\d{2,3}:", text, re.IGNORECASE):
                    break  # próxima observación, salimos
                else:
                    if seccion == "observacion":
                        obs_text.append(text)
                    elif seccion == "sustento":
                        sustento_text.append(text)
                    elif seccion == "solicitud":
                        solicitud_text.append(text)

                i += 1

            observaciones.append({
                "N° Obs": num_obs,
                "Título": titulo,
                "Observación": " ".join(obs_text).strip(),
                "Sustento": " ".join(sustento_text).strip(),
                "Solicitud": " ".join(solicitud_text).strip()
            })
        else:
            i += 1  # seguir buscando

    return observaciones
