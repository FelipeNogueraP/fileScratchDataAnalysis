import pandas as pd
import docx2txt


def load_data(xlsx_path, tipologias_path, criterios_path):
    df = pd.read_excel(xlsx_path)
    df_tipologias_text = docx2txt.process(tipologias_path)
    df_criterios_text = docx2txt.process(criterios_path)

    # Extract tipologias
    tipologias = [line.strip() for line in df_tipologias_text.split(
        "\n") if line.startswith("TIPOLOGÍA")]
    return df, tipologias, df_criterios_text


def check_keywords(text):
    keywords = [
        "Aciertos", "Acceso", "Citación", "Citar", "Clave", "Cite", "Comparar", "Comprobar", "Confirmar",
        "Conocer", "Copia del material", "Constatar", "Corroborar", "Cotejar", "Cuadernillo", "Examinar",
        "Exhibición", "Exponer", "Inspeccionar", "Material", "Mirar", "Mostrar", "Revisar", "Validar",
        "ver", "Verificar", "Visualizar", "Revisión manual", "Preguntas que respondió bien", "Preguntas que respondió mal"
    ]
    return any(keyword.lower() in text.lower() for keyword in keywords)


def process_row(row, tipologias):
    detalle = str(row["detalle"])

    # Check for VERDADERO
    if row["acceso_solicitud_pruebas"] == "VERDADERO":
        for col in row.index[7:-1]:
            row[col] = ""
        return row

    # Check for keywords
    if check_keywords(detalle):
        row["Tipología 1"] = 1
        row["observaciones"] = "FALSO POSITIVO"
        return row

    # Check for FALSO with ANEXO
    if "ANEXO" in detalle:
        for col in row.index[7:-1]:
            row[col] = 0
        row["observaciones"] = "VER ANEXO"
        return row

    # Check for tipologias
    matched_tipologia = next(
        (tipologia for tipologia in tipologias if tipologia.lower() in detalle.lower()), None)
    if matched_tipologia:
        col_name = matched_tipologia.split(" ")[1]
        row[f"Tipología {col_name}"] = 1
    else:
        row["RECLAMACIONES PARTICULARES"] = detalle
        for col in row.index[7:-1]:
            row[col] = 0

    # Update "ESTADO ANALISTA"
    row["ESTADO ANALISTA"] = "OK"
    return row


def process_data(df, tipologias):
    df_luisa = df[df["analista"] == "Luisa Figueroa"].copy()
    df_luisa_processed = df_luisa.apply(
        lambda row: process_row(row, tipologias), axis=1)
    return df_luisa_processed


def save_to_excel(df, output_path):
    df.to_excel(output_path, index=False)


if __name__ == "__main__":
    # Using the functions
    df, tipologias, criterios_text = load_data(
        "ruta_al_archivo.xlsx", "ruta_al_archivo_tipologias.docx", "ruta_al_archivo_criterios.docx")
    df_processed = process_data(df, tipologias)
    save_to_excel(df_processed, "ruta_donde_guardar.xlsx")
