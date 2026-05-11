from pathlib import Path
import pandas as pd


class LectorNominaExcel:
    """
    Clase encargada de leer el archivo Excel de nómina.

    Esta clase permite:
    - Ubicar el archivo Excel dentro de la carpeta BD del proyecto.
    - Mostrar las hojas disponibles del archivo Excel.
    - Seleccionar una hoja por número.
    - Detectar automáticamente la fila donde empiezan los encabezados.
    - Extraer únicamente las columnas necesarias.
    - Limpiar filas vacías, totales y notas del archivo.
    - Convertir columnas numéricas a formato decimal.
    - Devolver un DataFrame de pandas listo para ser usado en el main.
    """

    def __init__(self):
        """
        Inicializa la clase LectorNominaExcel.

        Construye la ruta del archivo Excel tomando como base la estructura
        del proyecto. También carga el archivo Excel, inicializa la hoja
        seleccionada y define las columnas necesarias que se van a extraer
        del archivo de nómina.
        """

        ruta_base = Path(__file__).resolve().parent.parent

        self.ruta_excel = ruta_base / "BD" / "NOMINA 2026.xlsx"

        self.excel = pd.ExcelFile(self.ruta_excel)

        self.hoja_seleccionada = None

        self.columnas_necesarias = [
            "NOMBRE",
            "SALUD Y PENSION",
            "DIAS LABORADOS",
            "DESCUENTO PRESTAMOS",
            "PRESTAMOS A AMC O TPTES.",
            "SUBTOTAL"
            
        ]

    def mostrar_hojas(self):
        """
        Muestra en pantalla las hojas disponibles en el archivo Excel.

        Las hojas se presentan como una lista numerada para que el usuario
        pueda seleccionar fácilmente una hoja escribiendo su número.
        """

        print("\nHojas disponibles en el archivo:")

        for i, hoja in enumerate(self.excel.sheet_names, start=1):
            print(f"{i}. {hoja}")

    def seleccionar_hoja_por_numero(self):
        """
        Permite seleccionar una hoja del archivo Excel mediante su número.

        Primero muestra la lista de hojas disponibles. Después solicita al
        usuario que escriba el número de la hoja que desea leer. Si el número
        ingresado es válido, guarda el nombre de la hoja seleccionada en
        self.hoja_seleccionada.

        Returns:
            str: Nombre de la hoja seleccionada.

        Raises:
            ValueError: Si el usuario no escribe un número válido.
            ValueError: Si el número seleccionado no existe en la lista.
        """

        self.mostrar_hojas()

        try:
            opcion = int(input("\nEscribe el número de la hoja que deseas leer: "))
        except ValueError:
            raise ValueError("Debes escribir un número válido.")

        if opcion < 1 or opcion > len(self.excel.sheet_names):
            raise ValueError("El número seleccionado no existe en la lista.")

        self.hoja_seleccionada = self.excel.sheet_names[opcion - 1]

        print(f"\nHoja seleccionada: {self.hoja_seleccionada}")

        return self.hoja_seleccionada

    def buscar_fila_encabezado(self):
        """
        Busca la fila donde se encuentran los encabezados reales de la tabla.

        Lee la hoja seleccionada sin asumir encabezados y recorre cada fila
        hasta encontrar una celda con el texto 'NOMBRE'. Esa fila se toma como
        la fila de encabezados de la tabla de nómina.

        Returns:
            int: Índice de la fila donde se encuentran los encabezados.

        Raises:
            ValueError: Si no se ha seleccionado una hoja previamente.
            ValueError: Si no se encuentra la columna 'NOMBRE' en la hoja.
        """

        if self.hoja_seleccionada is None:
            raise ValueError("Primero debes seleccionar una hoja.")

        df_raw = pd.read_excel(
            self.ruta_excel,
            sheet_name=self.hoja_seleccionada,
            header=None
        )

        for index, row in df_raw.iterrows():
            valores = row.astype(str).str.upper().str.strip().tolist()

            if "NOMBRE" in valores:
                return index

        raise ValueError("No se encontró la columna 'NOMBRE' en la hoja seleccionada.")

    def leer_hoja(self):
        """
        Lee la hoja seleccionada usando la fila correcta como encabezado.

        Primero identifica la fila donde están los encabezados reales mediante
        el método buscar_fila_encabezado(). Luego lee nuevamente la hoja usando
        esa fila como encabezado.

        También normaliza los nombres de las columnas convirtiéndolos a
        mayúsculas y eliminando espacios sobrantes. Si la columna
        '*PRESTAMOS A AMC O TPTES.' viene con asterisco, se renombra para que
        quede como 'PRESTAMOS A AMC O TPTES.'.

        Returns:
            pandas.DataFrame: DataFrame completo de la hoja seleccionada.
        """

        fila_encabezado = self.buscar_fila_encabezado()

        df = pd.read_excel(
            self.ruta_excel,
            sheet_name=self.hoja_seleccionada,
            header=fila_encabezado
        )

        df.columns = (
            df.columns
            .astype(str)
            .str.upper()
            .str.strip()
        )

        if "*PRESTAMOS A AMC O TPTES." in df.columns:
            df = df.rename(
                columns={
                    "*PRESTAMOS A AMC O TPTES.": "PRESTAMOS A AMC O TPTES."
                }
            )

        return df

    def obtener_dataframe(self):
        """
        Obtiene el DataFrame final con las columnas necesarias de la nómina.

        Si todavía no se ha seleccionado una hoja, solicita al usuario que
        seleccione una hoja por número. Luego lee la hoja, valida que existan
        las columnas requeridas y devuelve únicamente esas columnas.

        También realiza las siguientes limpiezas:
        - Elimina filas completamente vacías.
        - Elimina filas sin nombre.
        - Corta la información antes de la fila TOTALES.
        - Elimina filas de notas o textos inferiores del archivo.
        - Convierte columnas numéricas a formato float.
        - Reemplaza valores vacíos por 0 en columnas numéricas.

        Returns:
            pandas.DataFrame: DataFrame filtrado con las columnas:
                - NOMBRE
                - SUBTOTAL
                - DIAS LABORADOS
                - DESCUENTO PRESTAMOS
                - PRESTAMOS A AMC O TPTES.
                - SALUD Y PENSION

        Raises:
            ValueError: Si faltan una o más columnas necesarias en el Excel.
        """

        if self.hoja_seleccionada is None:
            self.seleccionar_hoja_por_numero()

        df = self.leer_hoja()

        columnas_faltantes = [
            col for col in self.columnas_necesarias
            if col not in df.columns
        ]

        if columnas_faltantes:
            print("\nNo se encontraron estas columnas:")

            for col in columnas_faltantes:
                print(f"- {col}")

            print("\nColumnas disponibles en la hoja:")

            for col in df.columns:
                print(f"- {col}")

            raise ValueError("Faltan columnas necesarias en el Excel.")

        df_filtrado = df[self.columnas_necesarias].copy()

        df_filtrado = df_filtrado.dropna(how="all")

        df_filtrado = df_filtrado.dropna(subset=["NOMBRE"])

        df_filtrado["NOMBRE"] = (
            df_filtrado["NOMBRE"]
            .astype(str)
            .str.strip()
        )

        filas_totales = df_filtrado[
            df_filtrado["NOMBRE"]
            .str.upper()
            .str.contains("TOTALES", na=False)
        ]

        if not filas_totales.empty:
            indice_totales = filas_totales.index[0]
            df_filtrado = df_filtrado.loc[:indice_totales - 1]

        palabras_a_excluir = [
            "NOTA",
            "SMMLV",
            "TRANSPORTE",
            "QUEDAN",
            "CUÁL",
            "CUAL",
            "PORCENTAJE"
        ]

        patron_exclusion = "|".join(palabras_a_excluir)

        df_filtrado = df_filtrado[
            ~df_filtrado["NOMBRE"]
            .str.upper()
            .str.contains(patron_exclusion, na=False)
        ]

        columnas_numericas = [
            "SUBTOTAL",
            "DIAS LABORADOS",
            "DESCUENTO PRESTAMOS",
            "PRESTAMOS A AMC O TPTES.",
            "SALUD Y PENSION"
        ]

        for columna in columnas_numericas:
            df_filtrado[columna] = pd.to_numeric(
                df_filtrado[columna],
                errors="coerce"
            )

        df_filtrado["SUBTOTAL"] = (
            df_filtrado["SUBTOTAL"]
            .fillna(0)
            .astype(float)
            .round(2)
        )

        df_filtrado["DIAS LABORADOS"] = (
            df_filtrado["DIAS LABORADOS"]
            .fillna(0)
            .astype(float)
            .round(2)
        )

        df_filtrado["DESCUENTO PRESTAMOS"] = (
            df_filtrado["DESCUENTO PRESTAMOS"]
            .fillna(0)
            .astype(float)
            .round(2)
        )

        df_filtrado["PRESTAMOS A AMC O TPTES."] = (
            df_filtrado["PRESTAMOS A AMC O TPTES."]
            .fillna(0)
            .astype(float)
            .round(2)
        )

        df_filtrado["SALUD Y PENSION"] = (
            df_filtrado["SALUD Y PENSION"]
            .fillna(0)
            .astype(float)
            .round(2)
        )

        df_filtrado = df_filtrado[
            ~(
                (df_filtrado["SUBTOTAL"] == 0) &
                (df_filtrado["DIAS LABORADOS"] == 0)
            )
        ]

        return df_filtrado