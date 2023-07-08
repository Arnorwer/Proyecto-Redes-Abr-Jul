import pandas as pd
import numpy as np
import math, sys

#revisamos el excel con pandas
dat = "data_io_copia.xlsx"
datos = pd.ExcelFile(dat)
print(datos.sheet_names)
print()
#guardamos las hojas de cada uno para poder revisar las celdas
#para las celdas se usa el formato [columna][fila]
hoja1 = pd.read_excel(dat,"f_and_ouput", header=None)
hoja2 = pd.read_excel(dat, "V_fuente", header=None, na_filter=False)
hoja3 = pd.read_excel(dat, "I_fuente", header=None, na_filter=False)
hoja4 = pd.read_excel(dat, "Z", header=None, na_filter=False)
hoja5 = pd.read_excel(dat, "VTH_AND_ZTH", header=None)
hoja6 = pd.read_excel(dat, "Sfuente", header=None)
hoja7 = pd.read_excel(dat, "S_Z", header=None)
hoja8 = pd.read_excel(dat, "Balance_S", header=None)
#print(hoja2)
#print(hoja4)

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
#V_fuente (Hoja2)
lista_Vangulos = list()
lista_V_Zc = list()
lista_V_Zl = list()
lista_V_Zr = list()
lista_Vrms = list()
lista_V_fasorial = list()
lista_VZeq = list()
#le damos un valor incial
lista_Vangulos = [0]
lista_V_Zc = [0]
lista_V_Zl = [0]
lista_V_Zr = [0]
lista_Vrms = [0]
lista_V_fasorial = [0]
lista_VZeq = [0]

#I_fuente (Hoja3)
lista_Iangulos = list()
lista_I_Zc = list()
lista_I_Zl = list()
lista_I_Zr = list()
lista_Irms = list()
lista_I_fasorial = list()
lista_IZeq = list()
#le damos un valor incial
lista_Iangulos = [0]
lista_I_Zc = [0]
lista_I_Zl = [0]
lista_I_Zr = [0]
lista_Irms = [0]
lista_I_fasorial = [0]
lista_IZeq = [0]

#Z (Hoja4)
lista_Z_Zc = list()
lista_Z_Zl = list()
lista_Z_Zr = list()
lista_Z_Zeq = list()
#le damos un valor inicial
lista_Z_Zc = [0]
lista_Z_Zl = [0]
lista_Z_Zr = [0]
lista_Z_Zeq = [0]


#Operaciones de los elementos para V_fuente (Hoja2)
#Revisamos primero si los valores en las listas no presentan errores
for n in range(1, len(V_fuente["Bus i"])+1):
    if type(hoja2[4][n]) == str or hoja2[4][n] < 0:
        sys.exit("Error de datos en Rf")
    elif type(hoja2[5][n]) == str or hoja2[5][n] < 0:
        sys.exit("Error de datos en Lf")
    elif type(hoja2[25][n]) == str or hoja2[25][n] < 0:
        sys.exit("Error de datos Cf")

for n in range(1, len(V_fuente["Bus i"])+1):
    #calculo del angulo de desfase
    V_angulo = round(w * hoja2[3][n])
    #print(angulo)
    #guardamos el angulo en una lista en orden
    lista_Vangulos.append(V_angulo)
    #Conversiones de mH a H y uF a F
    converion_CfV = hoja2[25][n] * 10 ** -6
    conversion_LfV = hoja2[5][n] * 10 ** -3
    #calculo de la reactancia
    V_Xl = w * conversion_LfV #reactancia del inductor
    V_Xc = w * converion_CfV #reactancia del capacitor
    if V_Xc != 0:
        V_Xc = 1 / V_Xc #reactancia del capacitor
    V_Xr = hoja2[4][n] #reactancia del resistor
    #calculo de impedancia
    V_Zl = np.complex_(V_Xl * 1j) #impedancia del inductor
    V_Zl = np.round(V_Zl, 4)
    V_Zc = np.complex_(V_Xc * -1j) #impedancia del capacitor
    V_Zc = np.round(V_Zc, 4)
    V_Zr = V_Xr #impedancia del resistor
    #guardamos los valores en listas
    lista_V_Zc.append(V_Zc)
    lista_V_Zl.append(V_Zl)
    lista_V_Zr.append(V_Zr)
    #calculo del voltaje rms
    Vrms = round(hoja2[2][n] / np.sqrt(2), 4)
    #guardamos en una lista
    lista_Vrms.append(Vrms)
    #calculo del voltaje en forma fasorial (CREO QUE ES ASI, IDK ??)
    V_fasorial = Vrms * np.cos(V_angulo) + np.complex_(Vrms * np.sin(V_angulo) * 1j)
    V_fasorial = np.round(V_fasorial, 4)
    #guardamos en una lista
    lista_V_fasorial.append(V_fasorial)

'''print("Listas: ")
print()
print(lista_angulos)
print(lista_V_Zc)
print(lista_V_Zl)
print(lista_V_Zr)
print(lista_Vrms)
print(lista_V_fasorial)
#print(hoja2)'''

#calculo para el caso en que hay varias fuentes en un mismo nodo, es decir, en serie
for i in range(1, len(hoja2[0])):
    for k in range(i + 1, len(hoja2[0])):
        if hoja2[0][i] == hoja2[0][k]:
            #calculos el Vrms en fasores equivalente en ese nodo
            Veq_fasorial = np.round(lista_V_fasorial[i] + lista_V_fasorial[k], 4)
            #sacamos los voltajes del nodo de la lista
            lista_V_fasorial.pop(i)
            lista_V_fasorial.pop(k-1)
            #guardamos el nuevo voltaje en la lista, en la pocicion de i
            lista_V_fasorial.insert(i, Veq_fasorial)
            #print(lista_V_fasorial)

#Operaciones de los elementos para I_fuente (Hoja3)
#Revisamos primero si los valores en las listas no presentan errores
for n in range(1, len(I_fuente["Bus i"])+1):
    if type(hoja3[4][n]) == str or hoja3[4][n] < 0:
        sys.exit("Error de datos en RfI")
    elif type(hoja3[5][n]) == str or hoja3[5][n] < 0:
        sys.exit("Error de datos en LfI")
    elif type(hoja3[25][n]) == str or hoja3[25][n] < 0:
        sys.exit("Error de datos CfI")

for n in range(1, len(I_fuente["Bus i"])+1):
    #calculo del angulo de desfase
    I_angulo = round(w * hoja3[3][n])
    #print(angulo)
    #guardamos el angulo en una lista en orden
    lista_Iangulos.append(I_angulo)
    #Conversiones de mH a H y uF a F
    converion_CfI = hoja3[25][n] * 10 ** -6
    conversion_LfI = hoja3[5][n] * 10 ** -3
    #calculo de la reactancia
    I_Xl = w * conversion_LfI #reactancia del inductor
    I_Xc = w * converion_CfI #reactancia del capacitor
    if I_Xc != 0:
        I_Xc = 1 / I_Xc #reactancia del capacitor
    I_Xr = hoja3[4][n] #reactancia del resistor
    #calculo de impedancia
    I_Zl = np.complex_(I_Xl * 1j) #impedancia del inductor
    I_Zl = np.round(I_Zl, 4)
    I_Zc = np.complex_(I_Xc * -1j) #impedancia del capacitor
    I_Zc = np.round(I_Zc, 4)
    I_Zr = (I_Xr) #impedancia del resistor
    #guardamos los valores en listas
    lista_I_Zc.append(I_Zc)
    lista_I_Zl.append(I_Zl)
    lista_I_Zr.append(I_Zr)
    #calculo del corriente rms
    Irms = round(hoja3[2][n] / np.sqrt(2), 4)
    #guardamos en una lista
    lista_Irms.append(Irms)
    #calculo del corriente en forma fasorial (CREO QUE ES ASI, IDK ??)
    I_fasorial = Irms * np.cos(I_angulo) + np.complex_(Irms * np.sin(I_angulo) * 1j)
    I_fasorial = np.round(I_fasorial, 4)
    #guardamos en una lista
    lista_I_fasorial.append(I_fasorial)

'''print("Listas: ")
print()
print(lista_angulos)
print(lista_I_Zc)
print(lista_I_Zl)
print(lista_I_Zr)
print(lista_Irms)
print(lista_I_fasorial)
#print(hoja3)'''

#calculo para el caso en que hay varias fuentes en un mismo nodo, es decir, en serie
for i in range(1, len(hoja3[0])):
    for k in range(i + 1, len(hoja3[0])):
        if hoja3[0][i] == hoja3[0][k]:
            #calculos el Vfasorial equivalente en ese nodo
            Ieq_fasorial = np.round(lista_I_fasorial[i] + lista_I_fasorial[k], 4)
            #sacamos las corrientes del nodo de la lista
            lista_I_fasorial.pop(i)
            lista_I_fasorial.pop(k-1)
            #guardamos la nueva corriente en la lista, en la pocicion de i
            lista_I_fasorial.insert(i, Ieq_fasorial)
            #print(lista_I_fasorial)

#Operaciones de los elementos para Z (Hoja4)
#Revisamos primero si los valores en las listas no presentan errores
for n in range(1, len(Z["Bus i"])+1):
    if type(hoja4[3][n]) == str or hoja4[3][n] < 0:
        sys.exit("Error de datos en R")
    elif type(hoja4[4][n]) == str or hoja4[4][n] < 0:
        sys.exit("Error de datos en L")
    elif type(hoja4[25][n]) == str or hoja4[25][n] < 0:
        sys.exit("Error de datos C")

for n in range(1, len(Z["Bus i"])+1):
    #Conversiones de uH a H y uF a F
    converion_C = hoja4[25][n] * 10 ** -6
    conversion_L = hoja4[4][n] * 10 ** -6
    #calculo de la reactancia
    ZL = w * conversion_L #reactancia del inductor 
    ZC = w * converion_C #reactancia del capacitor
    if ZC != 0:
        ZC = 1 / ZC #reactancia del capacitor
    ZR = hoja4[3][n] #reactancia del resistor
    #calculo de impedancia
    Z_Zl = np.complex_(ZL * 1j) #impedancia del inductor
    Z_Zl = np.round(Z_Zl, 4)
    Z_Zc = np.complex_(ZC * -1j) #impedancia del capacitor
    Z_Zc = np.round(Z_Zc, 4)
    Z_Zr = (ZR) #impedancia del resistor
    #guardamos los valores en listas
    lista_Z_Zc.append(Z_Zc)
    lista_Z_Zl.append(Z_Zl)
    lista_Z_Zr.append(Z_Zr)

'''print("Listas: ")
print()
print(lista_Z_Zc)
print(lista_Z_Zl)
print(lista_Z_Zr)
#print(hoja4)'''
#calculo de las impedancias que esten en paralelo
for i in range(1, len(hoja4[0])):
    for k in range(i + 1, len(hoja4[0])):
        if hoja4[0][i] == hoja4[0][k] and hoja4[1][i] == hoja4[1][k]:
            #Para las inductancias capacitivas
            if lista_Z_Zc[i] != 0 and lista_Z_Zc[k] != 0:
                Z_Zeq1 = 1 / lista_Z_Zc[i] + 1 / lista_Z_Zc[k]
            elif lista_Z_Zc[i] != 0 and lista_Z_Zc[k] == 0:
                Z_Zeq1 = 1 / lista_Z_Zc[i]
            elif lista_Z_Zc[i] == 0 and lista_Z_Zc[k] != 0:
                Z_Zeq1 = 1 / lista_Z_Zc[k]
            elif lista_Z_Zc[i] == 0 and lista_Z_Zc[k] == 0:
                Z_Zeq1 = 0
            #Para las inductancias inductivas
            if lista_Z_Zl[i] != 0 and lista_Z_Zl[k] != 0:
                Z_Zeq2 = 1 / lista_Z_Zl[i] + 1 / lista_Z_Zl[k]
            elif lista_Z_Zl[i] != 0 and lista_Z_Zl[k] == 0:
                Z_Zeq2 = 1 / lista_Z_Zl[i]
            elif lista_Z_Zl[i] == 0 and lista_Z_Zl[k] != 0:
                Z_Zeq2 = 1 / lista_Z_Zl[k]
            elif lista_Z_Zl[i] == 0 and lista_Z_Zl[k] == 0:
                Z_Zeq2 = 0
            #Para las inductancias resistivas
            if lista_Z_Zr[i] != 0 and lista_Z_Zr[k] != 0:
                Z_Zeq3 = 1 / lista_Z_Zr[i] + 1 / lista_Z_Zr[k]
            elif lista_Z_Zr[i] != 0 and lista_Z_Zr[k] == 0:
                Z_Zeq3 = 1 / lista_Z_Zr[i]
            elif lista_Z_Zr[i] == 0 and lista_Z_Zr[k] != 0:
                Z_Zeq3 = 1 / lista_Z_Zr[k]
            elif lista_Z_Zr[i] == 0 and lista_Z_Zr[k] == 0:
                Z_Zeq3 = 0