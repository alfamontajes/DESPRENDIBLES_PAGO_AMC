from modules.read_bd import LectorNominaExcel


def main():
    lector = LectorNominaExcel()

    df_nomina = lector.obtener_dataframe()

    print("\nDataFrame obtenido:")
    print(df_nomina)


if __name__ == "__main__":
    main()