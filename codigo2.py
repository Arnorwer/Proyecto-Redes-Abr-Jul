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
lista_V_Zc = list()
lista_V_Zl = list()
lista_V_Zr = list()
lista_Vrms = list()
lista_V_fasorial = list()
lista_VZeq = list()
#le damos un valor incial
lista_angulos = [0]
lista_V_Xc = [0]
lista_V_Xl = [0]
lista_V_Xr = [0]
lista_V_Zc = [0]
lista_V_Zl = [0]
lista_V_Zr = [0]
lista_Vrms = [0]
lista_V_fasorial = [0]
lista_VZeq = [0]

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
    V_Xr = round(hoja2[4][n], 4) #reactancia del resistor
    #guardamos los valores en listas
    lista_V_Xc.append(V_Xc)
    lista_V_Xl.append(V_Xl)
    lista_V_Xr.append(V_Xr)
    #calculo de impedancia
    V_Zl = (V_Xl) #impedancia del inductor
    V_Zc = (V_Xc) #impedancia del capacitor
    V_Zr = (V_Xr) #impedancia del resistor

    '''AGREGAR CONDICION PARA CUANDO ALGUNO DE LOS VALORES ES CERO O ESTA VACIO, 
    YA SEA PARA AGREGAR EN ESE LUGAR EL VALOR 1 O PARA QUE NO SE INCLUYA EN LA LISTA
    CREO QUE LO MEJOR SERIA QUE SE TOMARA COMO VALOR 1'''

    #guardamos los valores en listas
    lista_V_Zc.append(V_Zc)
    lista_V_Zl.append(V_Zl)
    lista_V_Zr.append(V_Zr)
    #calculo del voltaje rms
    Vrms = round(hoja2[2][n] / np.sqrt(2), 4)
    #guardamos en una lista
    lista_Vrms.append(Vrms)
    #calculo del voltaje en forma fasorial (CREO QUE ES ASI, IDK ??)
    V_fasorial = round(Vrms * np.cos(angulo) + Vrms * np.sin(angulo), 4)
    #guardamos en una lista
    lista_V_fasorial.append(V_fasorial)

'''print("Listas: ")
print()
print(lista_angulos)
print(lista_V_Xc)
print(lista_V_Xl)
print(lista_V_Xr)
print(lista_V_Zc)
print(lista_V_Zl)
print(lista_V_Zr)
print(lista_Vrms)
print(lista_V_fasorial)
#print(hoja2)'''

#ESTO SE DEBE CAMBIAR, SI LOS ELEMENTOS ESTAN EN EL MISMO NODO ES QUE ESTAN EN SERIE Y NO EN PARALELO
#SE DEBEN CALCULAR SON LAS FUENTES EN SERIE
#calculo para el caso de elementos en un mismo nodo, es decir, paralelo
for i in range(1, len(hoja2[0])):
    for k in range(i + 1, len(hoja2[0])):
        if hoja2[0][i] == hoja2[0][k]:
            #calculo de las impedancias en paralelo
            V_Zeq1 = round(1/lista_V_Zc[i] + 1/lista_V_Zc[k], 4)
            V_Zeq2 = round(1/lista_V_Zl[i] + 1/lista_V_Zl[k], 4)
            V_Zeq3 = round(1/lista_V_Zr[i] + 1/lista_V_Zr[k], 4)
            V_Zeq1 = round(1 / V_Zeq1, 4)
            V_Zeq2 = round(1 / V_Zeq2, 4)
            V_Zeq3 = round(1 / V_Zeq3,4)
            #conversion a imaginario
            V_Zeq1 = np.complex_(V_Zeq1 * -1j)
            V_Zeq2 = np.complex_(V_Zeq2 * 1j)
            #print(V_Zeq1)
            #print(V_Zeq2)
            #print(V_Zeq3)
            #calculamos el equivalante
            V_Zeq = V_Zeq2 + V_Zeq1 + V_Zeq3
            #print(V_Zeq)
            #sacamos los elementos anteriores de la lista
            #print(lista_V_Zc)
            lista_V_Zr.pop(i)
            lista_V_Zr.pop(k-1)
            lista_V_Zc.pop(i)
            lista_V_Zc.pop(k-1)
            lista_V_Zl.pop(i)
            lista_V_Zl.pop(k-1)
            #guardamos en una lista
            lista_VZeq.append(V_Zeq)

#convertimos en imaginario los elementos restantes de las listas
#print(lista_V_Zc)
lista_V_Zc = [n * -1j for n in lista_V_Zc]
#print(lista_V_Zc)
lista_V_Zl = [n * 1j for n in lista_V_Zl]
#print(lista_V_Zl)
#sumar todas las impedancias
Z_total = sum(lista_V_Zl) + sum(lista_V_Zc) + sum(lista_V_Zr) + sum(lista_VZeq)
#print(sum(lista_V_Zl))
#print(sum(lista_V_Zc))
#print(sum(lista_V_Zr))
#print(sum(lista_VZeq))
#print(Z_total)
#calculamos la corriente en su forma fasorial para el circuito
