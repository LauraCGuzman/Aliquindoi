class Config:
    def __init__(self):
        # read config file and save all the parameters
        from configparser import ConfigParser
        import pandas as pd
        import os
        #read config file
        config = ConfigParser()
        # 1. Construir la ruta absoluta
        base_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(base_dir, "config.ini")

        # Imprimimos la ruta para depurar
        print(f"Intentando leer configuración en: {config_path}")

        # 2. Leer directamente el archivo (SIN BUCLE FOR)
        if os.path.exists(config_path):
            config.read(config_path, encoding='utf-8')  # Importante: encoding
        else:
            # Si no existe, lanzamos error aquí para no seguir ejecutando
            raise FileNotFoundError(f"No se encuentra el archivo config.ini en: {config_path}")

        # 3. Comprobar si se cargaron secciones
        if not config.sections():
            raise ValueError("El archivo config.ini existe pero está vacío o no tiene cabeceras entre corchetes []")

        # 4. Cargar la sección específica
        if "plantillas" in config:
            config_plantillas_nombre = config["plantillas"]
            self.path_plantillas = {key: str(value) for key, value in config_plantillas_nombre.items()}
        else:
            raise KeyError("El archivo config.ini no tiene la sección [plantillas]")

        config_plantillas_nombre = config["plantillas"]
        self.path_plantillas = {key: str(value) for key, value in config_plantillas_nombre.items()}

        # Create a dictionary for each template specified in path_plantillas
        self.celdas_plantillas = {}
        for plantillas_name, plantillas_path in self.path_plantillas.items():
            # Read the corresponding section
            config_plantillas = config[plantillas_name]
            # Create a dictionary with keys and values of the section
            plantillas_datos = {key: value for key, value in config_plantillas.items()}
            # Add them to sensors
            self.celdas_plantillas[plantillas_name] = plantillas_datos
