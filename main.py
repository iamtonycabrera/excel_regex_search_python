import argparse
import json
import openpyxl
import re


def main():
    ### ARGPARSE ###
    parser = argparse.ArgumentParser(
        description="Procesar un archivo Excel según la configuración proporcionada en un archivo JSON."
    )
    # Argumento de entrada. Path a fichero en formato JSON con la configuración del script
    parser.add_argument(
        "-c",
        "--config-file",
        required=True,
        help="Ruta del archivo JSON de configuración",
    )
    # Argumento de salida. Path del fichero de resultado del script
    parser.add_argument(
        "-o", "--output-file", help="Ruta del archivo '.xlsx' de salida analizado"
    )
    # Anlaizar los argumentos proporcionados al script
    args = parser.parse_args()

    # Cargar configuración desde el JSON
    # Abrir JSON y leer (r)
    with open(args.config_file, "r") as config_file:
        config = json.load(config_file)

    # Extraer parámetros de configuración del JSON
    input_file = config.get("input")
    sheet_name = config.get("sheet")
    search_pattern = config.get("search")

    # Validar que la estructura del JSON está correcta
    if not (input_file and sheet_name and search_pattern):
        print("Estructura del JSON con errores")
        return

    # Nombre archivo salida. Si no se especifica, usar nombre del script con extensión .xlsx
    output_file = args.output_file or "%(prog)s.xlsx"
