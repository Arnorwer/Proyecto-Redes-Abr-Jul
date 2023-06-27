import pandas as pd
import numpy as np
import math

#revisamos el excel con pandas
dat = "data_io_copia.xlsx"
datos = pd.ExcelFile(dat)
print(datos.sheet_names)
print()
#guardamos las hojas de cada uno para poder revisar las celdas
#para las celdas se usa el formato [columna][fila]
hoja1 = pd.read_excel(dat,"f_and_ouput", header=None)
hoja2 = pd.read_excel(dat, "V_fuente", header=None)
hoja3 = pd.read_excel(dat, "I_fuente", header=None)
hoja4 = pd.read_excel(dat, "Z", header=None)
hoja5 = pd.read_excel(dat, "VTH_AND_ZTH", header=None)
hoja6 = pd.read_excel(dat, "Sfuente", header=None)
hoja7 = pd.read_excel(dat, "S_Z", header=None)
hoja8 = pd.read_excel(dat, "Balance_S", header=None)
#print(hoja2)
#print(hoja2[3])

#almcenamos las paginas de la tabla de excel
f_and_ouput = pd.read_excel(datos, 'f_and_ouput')
V_fuente = pd.read_excel(datos, 'V_fuente')
I_fuente = pd.read_excel(datos, 'I_fuente')
Z = pd.read_excel(datos, 'Z')
VTH_AND_ZTH = pd.read_excel(datos, 'VTH_AND_ZTH')
Sfuente = pd.read_excel(datos, 'Sfuente')
S_Z = pd.read_excel(datos, 'S_Z')
Balance_S = pd.read_excel(datos, 'Balance_S')

#calculamos el valor de w para poder hacer el wt y obtener el angulo de desfase
if hoja1[1][0] == 60: #si la frecuncia es 60Hz, el valor de w es directamente 60
    w = 377
else: #si la frecuencia no es 60, entonces se calcula manualmente el valor de w y se redondea a su entero mas cercano
    w = round(2 * math.pi * hoja1[1][0], 4)

#listas de valores
lista_angulos = list()
lista_V_Xc = list()
lista_V_Xl = list()
lista_V_Xr = list()
lisa_V_Zc = list()
lista_V_Zl = list()
lista_V_Zr = list()
#le damos un valor incial
lista_angulos = [0]
lista_V_Xc = [0]
lista_V_Xl = [0]
lista_V_Xr = [0]
lista_V_Zc = [0]
lista_V_Zl = [0]
lista_V_Zr = [0]

#Operaciones de los elementos para V_fuente (Hoja2)
for n in range(1, len(V_fuente["Bus i"])+1):
    #calculo del angulo de desfase
    angulo = round(w * hoja2[3][n])
    #print(angulo)
    #guardamos el angulo en una lista en orden
    lista_angulos.append(angulo)
    #Conversiones de mH a H y uF a F
    converion_Cf = round(hoja2[25][n] * 10 ** -3, 4)
    conversion_Lf = round(hoja2[5][n] * 10 ** -3, 4)
    #calculo de la reactancia
    V_Xl = round(w * conversion_Lf, 4) #reactancia del inductor
    V_Xc = round(w * converion_Cf, 4) #reactancia del capacitor
    V_Xc = round(1 / V_Xc, 4) #reactancia del capacitor
    V_Xr = hoja2[4][n] #reactancia del resistor
    #guardamos los valores en listas
    lista_V_Xc.append(V_Xc)
    lista_V_Xl.append(V_Xl)
    lista_V_Xr.append(V_Xr)
    #calculo de impedancia
    V_Zl = np.complex_(V_Xl * 1j) #impedancia del inductor
    V_Zc = np.complex_(V_Xc * 1j) #impedancia del capacitor
    V_Zr = np.complex_(V_Xr * 1j) #impedancia del resistor
    #guardamos los valores en listas
    lista_V_Zc.append(V_Zc)
    lista_V_Zl.append(V_Zl)
    lista_V_Zr.append(V_Zr)

    #suma de las impedancias de una misma linea desde las listas
    V
    

#print(hoja2)
#calculo para el caso de elementos en un mismo nodo, es decir, paralelo
#for i in range(len(V_fuente["Bus i"])):
    #for k in range(i + 1, len(V_fuente["Bus i"])):
        #print()