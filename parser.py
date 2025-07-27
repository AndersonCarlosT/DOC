from docx import Document

def limpiar_texto(texto):
    return texto.replace("\t", " ").replace("\xa0", " ").strip()

def es_negrita(parrafo):
    return any(run.bold for run in parrafo.runs if run.text.strip())

def parsear_docx(docx_file):
    doc = Document(docx_file)
    observaciones = []
    num_obs = None
    titulo = ""
    seccion_actual = None
    obs_text, sustento_text, solicitud_text = [], [], []

    for parrafo in doc.paragraphs:
        texto = limpiar_texto(parrafo.text)

        if not texto:
            continue  # saltar líneas vacías

        if texto.lower().startswith("n°") or texto.lower().startswith("nº") or texto.lower().startswith("n.o"):
            if num_obs:
                observaciones.append({
                    "N° Obs": num_obs,
                    "Título": limpiar_texto(titulo),
                    "Observación": limpiar_texto("\n\n".join(obs_text)),
                    "Sustento": limpiar_texto("\n\n".join(sustento_text)),
                    "Solicitud": limpiar_texto("\n\n".join(solicitud_text))
                })
                titulo, obs_text, sustento_text, solicitud_text = "", [], [], []

            num_obs = texto.split()[-1]  # último token como número
            seccion_actual = "titulo"
            continue

        if es_negrita(parrafo) and not titulo:
            titulo += " " + texto
            continue

        if texto.lower().startswith("observación"):
            seccion_actual = "observacion"
            continue
        elif texto.lower().startswith("sustento"):
            seccion_actual = "sustento"
            continue
        elif texto.lower().startswith("solicitud"):
            seccion_actual = "solicitud"
            continue

        if seccion_actual == "observacion":
            obs_text.append(texto)
        elif seccion_actual == "sustento":
            sustento_text.append(texto)
        elif seccion_actual == "solicitud":
            solicitud_text.append(texto)
        elif seccion_actual == "titulo":
            titulo += " " + texto

    if num_obs:
        observaciones.append({
            "N° Obs": num_obs,
            "Título": limpiar_texto(titulo),
            "Observación": limpiar_texto("\n\n".join(obs_text)),
            "Sustento": limpiar_texto("\n\n".join(sustento_text)),
            "Solicitud": limpiar_texto("\n\n".join(solicitud_text))
        })

    return observaciones
