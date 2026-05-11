from modules.read_bd import LectorNominaExcel


def main():
    lector = LectorNominaExcel()

    df_nomina = lector.obtener_dataframe()

    df_nomina["TOTAL"] = (
        df_nomina["SUBTOTAL"]
        + df_nomina["PRESTAMOS A AMC O TPTES."]
        - df_nomina["DESCUENTO PRESTAMOS"]
    )

    print("\nDataFrame obtenido:")
    print(df_nomina)


if __name__ == "__main__":
    main()