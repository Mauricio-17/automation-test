import pandas as pd
import numpy as np
from datetime import date, datetime

# ---- Recogiendo datos
data_frame = pd.read_csv("./data/retrieved_filtered_data.csv", parse_dates=["Creations"])

# ------ DATA
owners = data_frame["Owners"].unique()
iterable = ["Solange Cedron", "Angelica Manrique", "Vladir Colunga", "Maria Melendez", "Zoila Lozada",
            "Marcela Castillo", "Miguel Huarcaya"]
states_row = ["Cliente", "En oportunidad", "Caliente", "Templado", "Frio", "No contesta", "Perdido", "Cliente perdido",
              "Otros", "Visitante", "Total"]
sub_states_row = ["CLIENTE", "EN PROCESO", "EN PROCESO", "EN PROCESO", "POR TRABAJAR", "POR TRABAJAR", "PERDIDO",
                  "PERDIDO", "OTROS", "OTROS", " "]

# Get the current date
current_date = datetime.now()
current_date = datetime.strptime(datetime.strftime(current_date, "%Y-%m-%d"), "%Y-%m-%d")
current_month = current_date.month

data = {
    "Estados": states_row,
    "Sub estado": sub_states_row,
    "Solange Cedron": [],
    "Angelica Manrique": [],
    "Vladir Colunga": [],
    "Maria Melendez": [],
    "Zoila Lozada": [],
    "Marcela Castillo": [],
    "Miguel Huarcaya": []
}

if __name__ == '__main__':

    # Iterando cada columna - asesor
    for i in range(len(iterable)):
        owner = iterable[i]

        # Realizando consultas para obtener el total de leads por cada esta estado
        len_warm = len(data_frame[(data_frame["Owners"] == owner) & (data_frame["States"] == "Templado")])
        len_cold = len(data_frame[(data_frame["Owners"] == owner) & (data_frame["States"] == "Frio")])
        len_not_q = len(data_frame[(data_frame["Owners"] == owner) & (data_frame["States"] == "Por trabajar")])
        len_others = len(data_frame[(data_frame["Owners"] == owner) & (data_frame["States"] == "Otros")])
        len_lost = len(data_frame[(data_frame["Owners"] == owner) & (data_frame["States"] == "Perdido")])
        len_hot = len(data_frame[(data_frame["Owners"] == owner) & (data_frame["States"] == "Caliente")])
        len_in_deal = len(data_frame[(data_frame["Owners"] == owner) & (data_frame["States"] == "No contesta")])
        len_visitor = len(data_frame[(data_frame["Owners"] == owner) & (data_frame["States"] == "Visitante")])
        client = data_frame[(data_frame["Owners"] == owner) & (data_frame["States"] == "Cliente") & (
            data_frame["Creations"].dt.month == current_month)]
        len_client = len(client)
        len_client_lost = len(data_frame[(data_frame["Owners"] == owner) & (data_frame["States"] == "Cliente perdido")])

        # Obteniendo el total de leads agrupado por estados del asesor
        total = len_warm + len_cold + len_not_q + len_others + len_lost + len_hot + len_in_deal + len_visitor + \
                len_client + len_client_lost

        # Asignando la data a cada columna - asesor
        data[owner] = [len_client, len_in_deal, len_hot, len_warm, len_cold, len_not_q, len_lost, len_client_lost,
                       len_others, len_visitor, total]

    # Construyendo el DataFrame
    frame = pd.DataFrame(data=data)
    frame["Total"] = (frame["Solange Cedron"] + frame["Angelica Manrique"] + frame["Vladir Colunga"] + frame["Maria Melendez"] +
                      frame["Zoila Lozada"] + frame["Marcela Castillo"] + frame["Miguel Huarcaya"])

    sub_total = frame.copy()
    frame = frame.set_index("Estados")

    print(frame)

    # --------------------------------------- SEGUNDA TABLA

    # Filtrando columnas a utilizar
    sub_total = sub_total[["Estados", "Solange Cedron", "Angelica Manrique", "Vladir Colunga",
                                       "Maria Melendez", "Zoila Lozada", "Marcela Castillo", "Miguel Huarcaya"]]

    sub_total = sub_total.set_index("Estados")

    # Ordenando de acuerdo a los que tienen mayores leads asignados
    sub_total = sub_total.T
    sub_total = sub_total.sort_values(by="Total", ascending=False)
    sub_total = sub_total.T
    sub_total["Total"] = (sub_total["Marcela Castillo"] + sub_total["Zoila Lozada"] + sub_total["Solange Cedron"] + sub_total["Maria Melendez"] + sub_total["Vladir Colunga"] + sub_total["Miguel Huarcaya"] + sub_total["Angelica Manrique"])

    total = sub_total["Total"].iloc[-1]

    sub_total = sub_total.T # Invertir la matriz

    # Creando columnas como resultado de otras
    sub_total["CLIENTE"] = sub_total["Cliente"]
    sub_total["EN PROCESO"] = sub_total["En oportunidad"] + sub_total["Caliente"] + sub_total["Templado"]
    sub_total["POR TRABAJAR"] = sub_total["Frio"] + sub_total["No contesta"]
    sub_total["PERDIDO"] = sub_total["Perdido"] + sub_total["Cliente perdido"]
    sub_total["OTROS"] = sub_total["Otros"] + sub_total["Visitante"]
    sub_total = sub_total.T # Invertir la tabla

    # Estableciendo la columna PORCENTAJE para despues convertirla a fila
    sub_total["Porcentaje"] = round(sub_total["Total"] / total, 2) * 100

    # Inventir la tabla para filtrar columnas y despues reinvertirla
    sub_total = sub_total.T
    sub_total = sub_total[["CLIENTE", "EN PROCESO", "POR TRABAJAR", "PERDIDO", "OTROS", "Total"]]

    sub_total = sub_total.T

    # Estableciendo el formato de los pocentajes con su tipo de dato respectivo
    sub_total["Porcentaje"] = sub_total["Porcentaje"].astype(str)
    sub_total["Porcentaje"] = sub_total["Porcentaje"] + "%"
    sub_total = sub_total.T

    with pd.ExcelWriter('./data/report_test.xlsx') as writer:
        # Write each DataFrame to a different sheet in the same Excel file
        frame.to_excel(writer,
                     sheet_name='Primary')
        sub_total.to_excel(writer,
                     sheet_name='Secondary') 
