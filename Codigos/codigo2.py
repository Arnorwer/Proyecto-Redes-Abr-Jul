import pandas as pd
import numpy as np
import math, sys, openpyxl

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
lista_V_fasorial = list()
lista_VZeq = list()
lista_VYeq = list()
lista_Vnodos = list()
#le damos un valor incial
lista_V_fasorial = [0]
lista_VZeq = [0]
lista_VYeq = [0]

#I_fuente (Hoja3)
lista_I_fasorial = list()
lista_IZeq = list()
lista_IYeq = list()
lista_Inodos = list()
#le damos un valor incial
lista_I_fasorial = [0]
lista_IZeq = [0]
lista_IYeq = [0]

#Z (Hoja4)
lista_Z_Zieq = list()
lista_Z_Zkeq = list()
lista_Z_Yieq = list()
lista_Z_Ykeq = list()
lista_Zinodos = list()
lista_Zknodos = list()
#le damos un valor inicial
lista_Z_Zieq = [0]
lista_Z_Yieq = [0]
lista_Z_Zkeq = [0]
lista_Z_Ykeq = [0]

#calcular la cantidad de nodos
for n in range(1, len(V_fuente["Bus i"])+1):
    #guardamos los nodos de la hoja2 en una lista
    nodosV = hoja2[0][n]
    lista_Vnodos.append(nodosV)
for n in range(1, len(I_fuente["Bus i"])+1):
    #guardamos los nodos de la hoja3 en una lista
    nodosI = hoja3[0][n]
    lista_Inodos.append(nodosI)
for n in range(1, len(Z["Bus i"])+1):
    #guardamos los nodos de la hoja4 en una lista
    nodosZ = hoja4[0][n]
    lista_Zinodos.append(nodosZ)
    nodosZ = hoja4[1][n]
    lista_Zknodos.append(nodosZ)

lista_nodos = list()
lista_nodos = lista_Vnodos + lista_Inodos + lista_Zinodos + lista_Zknodos
lista_nodos = pd.unique(lista_nodos)
lista_nodos = lista_nodos.tolist()
lista_nodos = sorted(lista_nodos)
dim = len(lista_nodos)
if lista_nodos[0] == 0:
    dim = dim - 1
#metemos valores en las listas
lista = list()
lista = [0]
for n in range(dim):
    lista_V_fasorial.append(np.complex_(0))
    lista_VZeq.append(np.complex_(0))
    lista_VYeq.append(np.complex_(0))

    lista_I_fasorial.append(np.complex_(0))
    lista_IZeq.append(np.complex_(0))
    lista_IYeq.append(np.complex_(0))

    lista_Z_Zieq.append(np.complex_(0))
    lista_Z_Yieq.append(np.complex_(0))
    lista_Z_Zkeq.append(np.complex_(0))
    lista_Z_Ykeq.append(np.complex_(0))

#Revisamos primero si los valores en las listas no presentan errores
for n in range(1, len(V_fuente["Bus i"])+1):
    if hoja2[4][n] == 0:
        print()
    elif hoja2[5][n] == 0:
        print()
    elif hoja2[6][n] == 0:
        print()

for n in range(1, len(V_fuente["Bus i"])+1):
    if type(hoja2[4][n]) == str or hoja2[4][n] < 0:
        sys.exit("Error de datos en Rf")
    elif type(hoja2[5][n]) == str or hoja2[5][n] < 0:
        sys.exit("Error de datos en Lf")
    elif type(hoja2[6][n]) == str or hoja2[6][n] < 0:
        sys.exit("Error de datos Cf")

for n in range(1, len(V_fuente["Bus i"])+1):
    #calculo del angulo de desfase
    V_angulo = round(w * hoja2[3][n])
    #Conversiones de mH a H y uF a F
    converion_CfV = hoja2[6][n] * 10 ** -6
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
    #calculamos la impedancia equivalante de los elementos en serie
    V_Zeq = np.round(V_Zr + V_Zl + V_Zc, 4)
    #calculo de la admitancia
    if V_Zl != 0:
        V_Yl = 1 / V_Zl
    else:
        V_Yl = 0
    if V_Zc != 0:
        V_Yc = 1 / V_Zc
    else:
        V_Yc = 0
    if V_Zr != 0:
        V_Yr = 1 / V_Zr
    else:
        V_Yr = 0
    #calculamos la admitancia equivalente de los elementos en serie
    V_Yeq = np.round(V_Yr + V_Yl + V_Yc, 4)
    #guardamos los valores en listas
    lista_VZeq[hoja2[0][n]] = V_Zeq
    lista_VYeq[hoja2[0][n]] = V_Yeq
    #calculo del voltaje rms
    Vrms = round(hoja2[2][n] / np.sqrt(2), 4)
    #calculo del voltaje en forma fasorial (CREO QUE ES ASI, IDK ??)
    V_fasorial = Vrms * np.cos(V_angulo) + np.complex_(Vrms * np.sin(V_angulo) * 1j)
    V_fasorial = np.round(V_fasorial, 4)
    #guardamos en una lista
    lista_V_fasorial[hoja2[0][n]] = V_fasorial

#calculo para el caso en que hay varias fuentes en un mismo nodo, es decir, en serie
for i in range(1, len(hoja2[0])):
    for k in range(i + 1, len(hoja2[0])):
        if hoja2[0][i] == hoja2[0][k]:
            #calculos el Vrms en fasores equivalente en ese nodo
            Veq_fasorial = np.round(lista_V_fasorial[i] + lista_V_fasorial[k], 4)
            #sacamos los voltajes del nodo de la lista
            lista_V_fasorial.pop(i)
            lista_V_fasorial[k-1] = np.complex_(0)
            #guardamos el nuevo voltaje en la lista, en la pocicion de i
            lista_V_fasorial.insert(i, Veq_fasorial)
            #sumamos las impedancias en serie
            V_Zeq = lista_VZeq[i] + lista_VZeq[k]
            #los eliminamos de las listas
            lista_VZeq.pop(i)
            lista_VZeq[k-1] = np.complex_(0)
            #guardamos la nueva impedancia
            lista_VZeq.insert(i, V_Zeq)

#Operaciones de los elementos para I_fuente (Hoja3)
#Revisamos primero si los valores en las listas no presentan errores
for n in range(1, len(I_fuente["Bus i"])+1):
    if hoja3[4][n] == 0:
        print()
    elif hoja3[5][n] == 0:
        print()
    elif hoja3[6][n] == 0:
        print()

for n in range(1, len(I_fuente["Bus i"])+1):
    if type(hoja3[4][n]) == str or hoja3[4][n] < 0:
        sys.exit("Error de datos en RfI")
    elif type(hoja3[5][n]) == str or hoja3[5][n] < 0:
        sys.exit("Error de datos en LfI")
    elif type(hoja3[6][n]) == str or hoja3[6][n] < 0:
        sys.exit("Error de datos CfI")

for n in range(1, len(I_fuente["Bus i"])+1):
    #calculo del angulo de desfase
    I_angulo = round(w * hoja3[3][n])
    #Conversiones de mH a H y uF a F
    converion_CfI = hoja3[6][n] * 10 ** -6
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
    #calculo de la impedancia equivalente en serie
    I_Zeq = np.round(I_Zr + I_Zl + I_Zc, 4)
    #calculo de la admitancia
    if I_Zl != 0:
        I_Yl = 1 / I_Zl
    else:
        I_Yl = 0
    if I_Zc != 0:
        I_Yc = 1 / I_Zc
    else:
        I_Yc = 0
    if I_Zr != 0:
        I_Yr = 1 / I_Zr
    else:
        I_Yr = 0
    #calculamos la admitancia equivalente de los elementos en serie
    I_Yeq = np.round(I_Yr + I_Yl + I_Yc, 4)
    #guardamos los valores en listas
    lista_IZeq[hoja3[0][n]] = I_Zeq
    lista_IYeq[hoja3[0][n]] = I_Yeq
    #calculo del corriente rms
    Irms = round(hoja3[2][n] / np.sqrt(2), 4)
    #calculo del corriente en forma fasorial (CREO QUE ES ASI, IDK ??)
    I_fasorial = Irms * np.cos(I_angulo) + np.complex_(Irms * np.sin(I_angulo) * 1j)
    I_fasorial = np.round(I_fasorial, 4)
    #guardamos en una lista
    lista_I_fasorial[hoja3[0][n]] = I_fasorial

#calculo para el caso en que hay varias fuentes en un mismo nodo, es decir, en serie
for i in range(1, len(hoja3[0])):
    for k in range(i + 1, len(hoja3[0])):
        if hoja3[0][i] == hoja3[0][k]:
            #calculos el Vfasorial equivalente en ese nodo
            Ieq_fasorial = np.round(lista_I_fasorial[i] + lista_I_fasorial[k], 4)
            #sacamos las corrientes del nodo de la lista
            lista_I_fasorial.pop(i)
            lista_I_fasorial[k-1] = np.complex_(0)
            #guardamos la nueva corriente en la lista, en la pocicion de i
            lista_I_fasorial.insert(i, Ieq_fasorial)
            #print(lista_I_fasorial)
            #sumamos las impedancias en serie
            I_Zeq = lista_IZeq[i] + lista_IZeq[k]
            #los eliminamos de las listas
            lista_IZeq.pop(i)
            lista_IZeq[k-1] = np.complex_(0)
            #guardamos la nueva impedancia
            lista_IZeq.insert(i, I_Zeq)

#Operaciones de los elementos para Z (Hoja4)
#Revisamos primero si los valores en las listas no presentan errores
for n in range(1, len(Z["Bus i"])+1):
    if hoja4[3][n] == 0:
        print()
    elif hoja4[4][n] == 0:
        print()
    elif hoja4[5][n] == 0:
        print()

for n in range(1, len(Z["Bus i"])+1):
    if type(hoja4[3][n]) == str or hoja4[3][n] < 0:
        sys.exit("Error de datos en R")
    elif type(hoja4[4][n]) == str or hoja4[4][n] < 0:
        sys.exit("Error de datos en L")
    elif type(hoja4[5][n]) == str or hoja4[5][n] < 0:
        sys.exit("Error de datos C")

for n in range(1, len(Z["Bus i"])+1):
    #Conversiones de uH a H y uF a F
    converion_C = hoja4[5][n] * 10 ** -6
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
    #calculo de las impedancias en serie
    Z_Zeq = np.round(Z_Zr + Z_Zl + Z_Zc, 4)
    #calculo de la admitancia
    if Z_Zl != 0:
        Z_Yl = 1 / Z_Zl
    else:
        Z_Yl = 0
    if Z_Zc != 0:
        Z_Yc = 1 / Z_Zc
    else:
        Z_Yc = 0
    if Z_Zr != 0:
        Z_Yr = 1 / Z_Zr
    else:
        Z_Yr = 0
    #calculamos la admitancia equivalente de los elementos en serie
    Z_Yeq = np.round(Z_Yr + Z_Yl + Z_Yc, 4)
    #guardamos los valores en listas
    lista_Z_Zieq[hoja4[0][n]] = Z_Zeq
    lista_Z_Yieq[hoja4[0][n]] = Z_Yeq
    lista_Z_Zkeq[hoja4[1][n]] = Z_Zeq
    lista_Z_Ykeq[hoja4[1][n]] = Z_Yeq

#calculo de las impedancias que esten en paralelo en Bus i
for i in range(1, len(hoja4[0])):
    for k in range(i + 1, len(hoja4[0])):
        if hoja4[0][i] == hoja4[0][k] and hoja4[1][i] == hoja4[1][k]:
            #Para las inductancias
            if lista_Z_Zieq[i] != 0 and lista_Z_Zieq[k] != 0:
                Z_Zeq = 1 / lista_Z_Zieq[i] + 1 / lista_Z_Zieq[k]
            elif lista_Z_Zieq[i] != 0 and lista_Z_Zieq[k] == 0:
                Z_Zeq = 1 / lista_Z_Zieq[i]
            elif lista_Z_Zieq[i] == 0 and lista_Z_Zieq[k] != 0:
                Z_Zeq = 1 / lista_Z_Zieq[k]
            elif lista_Z_Zieq[i] == 0 and lista_Z_Zieq[k] == 0:
                Z_Zeq = 0
            #Suma de la impedancia equivalente
            if Z_Zeq != 0:
                Z_Zeq = np.round(1 / Z_Zeq)
            #sacamos de las listas las impedancias en paralelo
            #sacamos los elementos anteriores de la listas
            lista_Z_Zieq.pop(i)
            lista_Z_Zieq[k-1] = np.complex_(0)
            #guardamos los datos en una lista
            lista_Z_Zieq.insert(i, Z_Zeq)
            #Para las admitancias
            Z_Yeq = np.round(lista_Z_Yieq[i] + lista_Z_Yieq[k], 4)
            #sacamos los elementos anteriores de la listas
            lista_Z_Yieq.pop(i)
            lista_Z_Yieq[k-1] = np.complex_(0)
            #guardamos los datos en una lista
            lista_Z_Yieq.insert(i, Z_Yeq)

#calculo de las impedancias que esten en paralelo en Bus j
for i in range(1, len(hoja4[1])):
    for k in range(i + 1, len(hoja4[1])):
        if hoja4[1][i] == hoja4[1][k] and hoja4[0][i] == hoja4[0][k]:
            #Para las inductancias
            if lista_Z_Zkeq[i] != 0 and lista_Z_Zkeq[k] != 0:
                Z_Zeq = 1 / lista_Z_Zkeq[i] + 1 / lista_Z_Zkeq[k]
            elif lista_Z_Zkeq[i] != 0 and lista_Z_Zkeq[k] == 0:
                Z_Zeq = 1 / lista_Z_Zkeq[i]
            elif lista_Z_Zkeq[i] == 0 and lista_Z_Zkeq[k] != 0:
                Z_Zeq = 1 / lista_Z_Zkeq[k]
            elif lista_Z_Zkeq[i] == 0 and lista_Z_Zkeq[k] == 0:
                Z_Zeq = 0
            #Suma de la impedancia equivalente
            if Z_Zeq != 0:
                Z_Zeq = np.round(1 / Z_Zeq)
            #sacamos de las listas las impedancias en paralelo
            #sacamos los elementos anteriores de la listas
            lista_Z_Zkeq.pop(i)
            lista_Z_Zkeq[k-1] = np.complex_(0)
            #guardamos los datos en una lista
            lista_Z_Zkeq.insert(i, Z_Zeq)
            #Para las admitancias
            Z_Yeq = np.round(lista_Z_Ykeq[i] + lista_Z_Ykeq[k], 4)
            #sacamos los elementos anteriores de la listas
            lista_Z_Ykeq.pop(i)
            lista_Z_Ykeq[k-1] = np.complex_(0)
            #guardamos los datos en una lista
            lista_Z_Ykeq.insert(i, Z_Yeq)

#crear la matriz
ybus_array = np.zeros((dim, dim), dtype = "complex_")

for i in range(dim):
    for j in range(dim):
        if i == j:
            ybus_array[i][j] = lista_VYeq[i+1] + lista_IYeq[i+1] + lista_Z_Yieq[i+1] + lista_Z_Ykeq[j+1]
        elif i != j:
            ybus_array[i][j] = lista_VYeq[i+1] + lista_IYeq[i+1] + lista_Z_Yieq[i+1] + lista_VYeq[j+1] + lista_IYeq[j+1] + lista_Z_Ykeq[j+1]

print(ybus_array)