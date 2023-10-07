import pandas as pd
import docx2txt
import os
from listadoPalabras import PALABRAS_CLAVE


class VerdaderoHandler:
    """VERDADEROS (los que SÍ solicitan acceso al material de la prueba):
    1. Leer en detalle las tipologías.
    2. Si en la casilla de acceso a la prueba está en estado VERDADERO no se debe tabular. (dejar
    en blanco todas las casillas).
    3. Identificar si hay un tema nuevo o diferente a las tipologías creadas, en caso positivo, se
    deberá incluir en la columna “RECLAMACIONES PARTICULARES” lo solicitado por el aspirante en forma textual."""

    @staticmethod
    def handle_verdadero(row):
        detalle_lower = row['detalle'].lower()
        if row['acceso_solicitud_pruebas'] == 'VERDADERO':
            for col in row.index[row.index.get_loc('auditor') + 1:]:
                row[col] = ''
            for palabra in PALABRAS_CLAVE:
                if palabra in detalle_lower:
                    row['RECLAMACIONES PARTICULARES'] = row['detalle']
                    break
        return row


class FalsosHandler:
    def __init__(self, row):
        self.row = row

    def is_falso_condicionado(self):
        return 'ANEXO' in self.row['asunto'] or 'ANEXO' in self.row['detalle']

    def is_falso_positivo(self):
        detalle_lower = self.row['detalle'].lower()
        return any(palabra in detalle_lower for palabra in PALABRAS_CLAVE)

    def handle(self):
        if self.is_falso_condicionado():
            # El código original para falso condicionado
            for col in self.row.index:
                if col not in ['id_reclamacion', 'nombre', 'identificacion', 'nombre2', 'apellido', 'email', 'id_inscripcion', 'nro_opec', 'denominacion', 'nivel', 'grado', 'estado_reclamacion', 'fecha_reclamacion', 'estado_inicial', 'detalle', 'asunto', 'con_anexo', 'descripcion', 'acceso_solicitud_pruebas', 'analista', 'auditor']:
                    self.row[col] = 0
            self.row['observaciones'] = 'VER ANEXO'
        elif self.is_falso_positivo():
            # El código original para falso positivo
            self.row['Tipologia 1'] = 1
            self.row['observaciones'] = 'FALSO POSITIVO'
            for col in self.row.index:
                if col not in ['id_reclamacion', 'nombre', 'identificacion', 'nombre2', 'apellido', 'email', 'id_inscripcion', 'nro_opec', 'denominacion', 'nivel', 'grado', 'estado_reclamacion', 'fecha_reclamacion', 'estado_inicial', 'detalle', 'asunto', 'con_anexo', 'descripcion', 'acceso_solicitud_pruebas', 'analista', 'auditor', 'Tipologia 1', 'observaciones']:
                    self.row[col] = 0
        return self.row


def load_data(xlsx_path, tipologias_path, criterios_path):
    df = pd.read_excel(xlsx_path)
    df_tipologias_text = docx2txt.process(tipologias_path)
    df_criterios_text = docx2txt.process(criterios_path)

    # Extract tipologias
    tipologias = [line.strip() for line in df_tipologias_text.split(
        "\n") if line.startswith("TIPOLOGÍA")]
    return df, tipologias, df_criterios_text


def process_row(row, tipologias):
    if row["acceso_solicitud_pruebas"] == "VERDADERO":
        row = VerdaderoHandler.handle_verdadero(row)
    elif row["acceso_solicitud_pruebas"] == "FALSO":
        handler = FalsosHandler(row)
        row = handler.handle()
    return row


def save_to_excel(df, path):
    if os.path.exists(path):
        base, ext = os.path.splitext(path)
        counter = 1
        while os.path.exists(path):
            path = base + f"_{counter}" + ext
            counter += 1
    df.to_excel(path, index=False)


def process_data_v2(df_original, tipologias):
    df_processed = df_original.copy()

    df_luisa = df_processed[df_processed["analista"]
                            == "Luisa Figueroa"].copy()
    df_luisa_processed = df_luisa.apply(
        lambda row: process_row(row, tipologias), axis=1)
    df_processed.update(df_luisa_processed)

    return df_processed


if __name__ == "__main__":
    df, tipologias, criterios_text = load_data(
        "/Users/Felipe/Downloads/inicial.xlsx",
        "/Users/Felipe/Downloads/Consolidado_Tipologias_Reclamaciones.docx",
        "/Users/Felipe/Downloads/CRITERIOS_DE_ANALISIS.docx")
    df_processed = process_data_v2(df, tipologias)
    save_to_excel(df_processed, "/Users/Felipe/Downloads/resultado.xlsx")
    print("Realizado con exito")
