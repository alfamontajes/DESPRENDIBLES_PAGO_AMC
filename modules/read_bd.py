import pandas as pd


class LectorNominaExcel:
    """
    Clase encargada de leer un archivo Excel de nómina.

    Permite mostrar las hojas disponibles, seleccionar una hoja por número,
    detectar automáticamente la fila donde empiezan los encabezados y devolver
    un DataFrame de pandas con las columnas necesarias para el procesamiento
    de la nómina.
    """

    def __init__(self):
        """
        Inicializa la clase LectorNominaExcel.

        Define la ruta del archivo Excel, carga el archivo para acceder a sus
        hojas, inicializa la hoja seleccionada como None y define las columnas
        que se van a extraer del archivo.
        """

        self.ruta_excel = r"BD/NOMINA 2026.xlsx"

        self.excel = pd.ExcelFile(self.ruta_excel)

        self.hoja_seleccionada = None

        self.columnas_necesarias = [
            "NOMBRE",
            "DIAS LABORADOS",
            "DESCUENTO PRESTAMOS",
            "SUBTOTAL"
        ]

    def mostrar_hojas(self):
        """
        Muestra en pantalla las hojas disponibles en el archivo Excel.

        Las hojas se presentan en forma de lista numerada para que el usuario
        pueda seleccionar fácilmente una hoja escribiendo su número.
        """

        print("\nHojas disponibles en el archivo:")

        for i, hoja in enumerate(self.excel.sheet_names, start=1):
            print(f"{i}. {hoja}")

    def seleccionar_hoja_por_numero(self):
        """
        Permite seleccionar una hoja del archivo Excel mediante su número.

        Primero muestra la lista de hojas disponibles, luego solicita al usuario
        que escriba el número de la hoja que desea leer. Si el número ingresado
        es válido, guarda el nombre de la hoja seleccionada en el atributo
        self.hoja_seleccionada.

        Returns:
            str: Nombre de la hoja seleccionada.

        Raises:
            ValueError: Si el usuario no escribe un número válido.
            ValueError: Si el número seleccionado no existe en la lista de hojas.
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

        Lee la hoja seleccionada sin asumir encabezados y recorre cada fila hasta
        encontrar una celda con el texto 'NOMBRE'. Esa fila se toma como la fila
        de encabezados de la tabla de nómina.

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
        esa fila como encabezado y normaliza los nombres de las columnas,
        convirtiéndolos a mayúsculas y eliminando espacios sobrantes.

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

        return df

    def obtener_dataframe(self):
        """
        Obtiene el DataFrame final con las columnas necesarias de la nómina.

        Si todavía no se ha seleccionado una hoja, permite seleccionar una.
        Luego lee la hoja, valida las columnas necesarias, corta la información
        antes de la fila TOTALES y elimina filas vacías o que no correspondan
        a empleados.

        Returns:
            pandas.DataFrame: DataFrame filtrado solo con los empleados.
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

        # Eliminar filas completamente vacías
        df_filtrado = df_filtrado.dropna(how="all")

        # Eliminar filas donde no haya nombre
        df_filtrado = df_filtrado.dropna(subset=["NOMBRE"])

        # Limpiar la columna NOMBRE
        df_filtrado["NOMBRE"] = (
            df_filtrado["NOMBRE"]
            .astype(str)
            .str.strip()
        )

        # Buscar la fila donde aparece TOTALES
        filas_totales = df_filtrado[
            df_filtrado["NOMBRE"].str.upper().str.contains("TOTALES", na=False)
        ]

        # Si encuentra TOTALES, corta el DataFrame antes de esa fila
        if not filas_totales.empty:
            indice_totales = filas_totales.index[0]
            df_filtrado = df_filtrado.loc[:indice_totales - 1]

        # Eliminar posibles filas de notas o textos que no sean empleados
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
            ~df_filtrado["NOMBRE"].str.upper().str.contains(patron_exclusion, na=False)
        ]

        # Convertir columnas numéricas
        columnas_numericas = [
            "SUBTOTAL",
            "DIAS LABORADOS",
            "DESCUENTO PRESTAMOS"
        ]

        for columna in columnas_numericas:
            df_filtrado[columna] = pd.to_numeric(
                df_filtrado[columna],
                errors="coerce"
            )

        # Opcional: eliminar filas donde SUBTOTAL y DIAS LABORADOS estén vacíos
        df_filtrado = df_filtrado.dropna(
            subset=["SUBTOTAL", "DIAS LABORADOS"],
            how="all"
        )

        return df_filtrado