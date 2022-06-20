import pandas as pd
import datetime as dt
import time
import gspread
import gspread_dataframe as gdf

start = time.process_time()

data = pd.read_excel("C:/Users/migue/OneDrive/Documents/Cabify - Propuesta/Business_Case_Dataset.xlsx")
data["Arrived At Local Dttm"] = pd.to_datetime(data["Arrived At Local Dttm"])
data["Start At Local Dttm"] = pd.to_datetime(data["Start At Local Dttm"])

arrived_time = data["Arrived At Local Dttm"].to_list()
start_time = data["Start At Local Dttm"].to_list()

tiempo_en_llegar = []
for i in range(len(arrived_time)):
    minutos = (arrived_time[i] - start_time[i]).seconds / 60
    tiempo_en_llegar.append(minutos)

data["Tiempo en Llegar"] = tiempo_en_llegar

lat_inicio = data["Real Start Lat"].to_list()
lon_inicio = data["Real Start Lon"].to_list()

coordenadas_inicio = []
for i in range(len(lat_inicio)):
    latitud = lat_inicio[i].replace(",", ".")
    longitud = lon_inicio[i].replace(",", ".")
    coordenada = str(latitud) + "," + str(longitud)
    coordenadas_inicio.append(coordenada)

data["Coordenadas Inicio"] = coordenadas_inicio
data["Hora Comienzo"] = data["Start At Local Dttm"].dt.hour

horas = [k for k in range(0, 24)]
horas.append(0)
horas_2 = [k for k in range(0, 24, 2)]
horas_2.append(0)
horas_3 = [k for k in range(0, 24, 3)]    
horas_3.append(0)


columna_horas = data["Hora Comienzo"].to_list()


rango_horario_1 = []
for i in range(len(columna_horas)):
    for a in range(len(horas)):
        if columna_horas[i] == horas[a]:
            rango_horario_1.append(str(horas[a]) + " - " + str(horas[a + 1]))
            break
        else:
            pass

rango_horario_2 = []
for i in range(len(columna_horas)):
    for a in range(len(horas_2)):
        if a != len(horas_2) - 2:
            if columna_horas[i] < horas_2[a + 1]:
                rango_horario_2.append(str(horas_2[a]) + " - " + str(horas_2[a + 1]))
                break
            else:
                pass
        else:
            rango_horario_2.append(str(horas_2[a]) + " - " + str(horas_2[a + 1]))
            break

rango_horario_3 = []
for i in range(len(columna_horas)):
    for a in range(len(horas_3)):
        if a != len(horas_3) - 2:
            if columna_horas[i] < horas_3[a + 1]:
                rango_horario_3.append(str(horas_3[a]) + " - " + str(horas_3[a + 1]))
                break
            else:
                pass
        else:
            rango_horario_3.append(str(horas_3[a]) + " - " + str(horas_3[a + 1]))
            break
                

data["Rango Horario 1"] = rango_horario_1
data["Rango Horario 2"] = rango_horario_2
data["Rango Horario 3"] = rango_horario_3

data = data.reset_index(drop = False)

rangos_horarios = {1: ["Cada una hora", "Rango Horario 1"], 2: ["Cada dos horas", "Rango Horario 2"],
                   3: ["Cada tres horas", "Rango Horario 3"]}

frames = []
for i in range(len(data)):
    ds = data[data["index"] == i]
    for key in rangos_horarios.keys():
        ds.loc[i + key] = ds.loc[i]
        ds.loc[i + key, "Frecuencia Horaria"] = rangos_horarios[key][0]
        ds.loc[i + key, "Franja Horaria"] = ds.loc[i, rangos_horarios[key][1]]
    ds = ds.drop([i], axis = 0)
    ds = ds.drop(["Rango Horario 1", "Rango Horario 2", "Rango Horario 3", "index"],
                 axis = 1)
    ds = ds.reset_index(drop = False)
    ds["ID"] = ds["index"].astype(str) + ds["Journey Id"]
    ds = ds.drop("index", axis = 1)
    ds = ds.set_index("ID", drop = True)
    frames.append(ds)
    

data_final = pd.concat(frames, axis = 0)


gc = gspread.service_account(filename = "C:/Users/migue/OneDrive/Documents/Cabify - Propuesta/Sheets-Python.json")
sh = gc.open("Base Trayectos")
base = sh.get_worksheet(0)
base.clear()

gdf.set_with_dataframe(base, data_final)


end = time.process_time()

tiempo = round((end - start) / 60, 2)

print(f"\nEl script tarda {tiempo} minutos en correr")    