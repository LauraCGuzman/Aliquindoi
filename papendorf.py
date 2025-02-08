input_text_file = "C:/Users/camp_la/Documents/Laura/OPAC/proyects/roll2sol/datos_matlab/details.txt"

import pandas as pd
from datetime import datetime
import numpy as np


def load_papendorf_iv_data(config_params):
    """
    Function to load Papendorf IV data from a text file specified in the config_params.

    Parameters:
        config_params (dict): Configuration dictionary containing the path to the data file.

    Returns:
        pd.DataFrame: A DataFrame containing the imported data.
    """

    # Define the file path
    file_path = config_params['SoilingAnalysis']['txtFileWithIrradianceData']

    # Define column names and types
    column_names = [
        "Zeitstempel", "Gertenummer", "Seriennummer", "Kennlinie",
        "LeerlaufspannungMesswert1V", "LeerlaufspannungMesswert2V", "LeerlaufspannungMesswert3V",
        "LeerlaufspannungMesswert4V", "LeerlaufspannungMesswert5V", "LeerlaufspannungMesswert6V",
        "KurzschlussstromMesswert1A", "KurzschlussstromMesswert2A", "KurzschlussstromMesswert3A",
        "KurzschlussstromMesswert4A", "KurzschlussstromMesswert5A", "KurzschlussstromMesswert6A",
        "MPPSpannungMesswert1V", "MPPSpannungMesswert2V", "MPPSpannungMesswert3V",
        "MPPSpannungMesswert4V", "MPPSpannungMesswert5V", "MPPSpannungMesswert6V",
        "MPPStromMesswert1A", "MPPStromMesswert2A", "MPPStromMesswert3A", "MPPStromMesswert4A",
        "MPPStromMesswert5A", "MPPStromMesswert6A", "Bestrahlungsstrke1Wm", "Bestrahlungsstrke2Wm",
        "Bestrahlungsstrke3Wm", "Bestrahlungsstrke4Wm", "Bestrahlungsstrke5Wm", "Bestrahlungsstrke6Wm",
        "ModultemperaturMesswert1C", "ModultemperaturMesswert2C", "ModultemperaturMesswert3C",
        "ModultemperaturMesswert4C", "ModultemperaturMesswert5C", "ModultemperaturMesswert6C",
        "SensortemperaturMesswert1C", "SensortemperaturMesswert2C", "SensortemperaturMesswert3C",
        "SensortemperaturMesswert4C", "SensortemperaturMesswert5C", "SensortemperaturMesswert6C",
        "Status", "KhlkrpertemperaturC", "TemperaturC"
    ]

    # Define data types
    column_types = {
        "Zeitstempel": "datetime64[ns]",
        "Gertenummer": "float64",
        "Seriennummer": "category",
        "Kennlinie": "string",
        # Define other column types as needed, setting numeric fields to 'float64'
    }

    # Read the data file
    try:
        data = pd.read_csv(
            file_path,
            sep="\t",
            skiprows=1,
            names=column_names,
            dtype=column_types,
            parse_dates=["Zeitstempel"],
            date_parser=lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S'),
            na_values=["", "NA", "null"],
            engine="python"
        )
    except Exception as e:
        raise RuntimeError(f"Error loading data from file {file_path}: {e}")

    return data

def extract_iv_curve_data(papendorf_iv_data):
    """
    Extract IV curve data from the provided DataFrame.

    Parameters:
        papendorf_iv_data (pd.DataFrame): Input data containing IV curve information.

    Returns:
        list of dict: A list where each entry contains the extracted IV curve data
                      for a specific timestamp.
    """
    iv_curves = []

    for index, row in papendorf_iv_data.iterrows():
        hex_val_long = str(row["Kennlinie"])
        current_values = []
        voltage_values = []

        # Extract voltage (V) and current (I) from the hex string
        for j in range(0, len(hex_val_long), 8):
            hex_val = hex_val_long[j:j+8]
            if len(hex_val) == 8:
                # Voltage (first 4 bytes)
                voltage = int(hex_val[2:4] + hex_val[0:2], 16) / 100
                voltage_values.append(voltage)

                # Current (last 4 bytes)
                current = int(hex_val[6:8] + hex_val[4:6], 16) / 1000
                current_values.append(current)

        # Create a dictionary for the IV curve
        iv_curve = {
            "I": np.array(current_values),
            "V": np.array(voltage_values),
            "timeStamp": row["Zeitstempel"],
            "deviceNumber": row["Gertenummer"]
        }

        iv_curves.append(iv_curve)

    return iv_curves

config_params = {
    'SoilingAnalysis': {
        'txtFileWithIrradianceData': input_text_file} }
data = load_papendorf_iv_data(config_params)
print(data.head())

iv_curves = extract_iv_curve_data(data)
print(iv_curves)