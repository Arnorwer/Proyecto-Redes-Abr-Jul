import pandas as pd
import numpy as np
import math, sys, openpyxl, cmath

#revisamos el excel con pandas
dat = "data_io_copia.xlsx"
datos = pd.ExcelFile(dat)
#print(datos.sheet_names)
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
lista_Vnodos = list()
lista_IV = list()
#le damos un valor incial
lista_V_fasorial = [0]
lista_VZeq = [0]
lista_IV = [0]

#I_fuente (Hoja3)
lista_I_fasorial = list()
lista_IZeq = list()
lista_Inodos = list()
#le damos un valor incial
lista_I_fasorial = [0]
lista_IZeq = [0]

#Z (Hoja4)
lista_Z_Zieq = list()
lista_Z_Zkeq = list()
lista_Zinodos = list()
lista_Zknodos = list()
#le damos un valor inicial
lista_Z_Zieq = [0]
lista_Z_Zkeq = [0]

#S_fuente
lista_S_fuente = list()
lista_S_fuente = [0]

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
    lista_IV.append(0)

    lista_I_fasorial.append(np.complex_(0))
    lista_IZeq.append(np.complex_(0))

    lista_Z_Zieq.append(np.complex_(0))
    lista_Z_Zkeq.append(np.complex_(0))
    
    lista_S_fuente.append(np.complex_(0))
    
#Matriz principal
listas_MP = []
for i in range(dim):
    lista_MP = []
    listas_MP.append(lista_MP)
    
lista_nodo1 = list()

#Revisamos que entre las listas V_fuente e I_fuente no hayan nodos repetidos
for i in range(1, len(V_fuente["Bus i"])+1):
    for n in range(1, len(I_fuente["Bus i"])+1):
        if hoja2[0][i] == hoja3[0][n]:
            hoja2[1][i] = "Warning"
            hoja3[1][n] = "Warning"
            print(hoja2)
            print(hoja3)
            sys.exit("Error de datos en Fuentes")

#Revisamos primero si los valores en las listas no presentan errores

'''for n in range(1, len(V_fuente["Bus i"])+1):
    if hoja2[4][n] == 0:
        print("b")
    elif hoja2[5][n] == 0:
        print("a")
    elif hoja2[6][n] == 0:
        print("c")

for n in range(1, len(V_fuente["Bus i"])+1):
    if type(hoja2[4][n]) == str or hoja2[4][n] < 0:
        sys.exit("Error de datos en Rf")
    elif type(hoja2[5][n]) == str or hoja2[5][n] < 0:
        sys.exit("Error de datos en Lf")
    elif type(hoja2[6][n]) == str:
        sys.exit("Error de datos Cf")'''

for n in range(1, len(V_fuente["Bus i"])+1):
    #calculo del angulo de desfase
    V_angulo = w * hoja2[3][n]
    #Conversiones de mH a H y uF a F
    converion_CfV = hoja2[6][n] * 10 ** -3
    conversion_LfV = hoja2[5][n] * 10 ** -3
    #calculo de la reactancia
    V_Xl = w * conversion_LfV #reactancia del inductor
    V_Xc = w * converion_CfV #reactancia del capacitor
    if V_Xc != 0:
        V_Xc = 1 / V_Xc #reactancia del capacitor
    V_Xr = hoja2[4][n] #reactancia del resistor
    #calculo de impedancia
    V_Zl = np.complex_(V_Xl * 1j) #impedancia del inductor
    V_Zc = np.complex_(V_Xc * -1j) #impedancia del capacitor
    V_Zr = V_Xr #impedancia del resistor
    #calculamos la impedancia equivalante de los elementos en serie
    V_Zeq = np.round(V_Zr + V_Zl + V_Zc, 4)
    #guardamos los valores en listas
    lista_VZeq[hoja2[0][n]] = V_Zeq
    #calculo del voltaje rms
    Vrms = hoja2[2][n] / np.sqrt(2)
    #calculo del voltaje en forma fasorial
    V_fasorial = Vrms * np.cos(V_angulo) + np.complex_(Vrms * np.sin(V_angulo) * 1j)
    #guardamos en una lista
    lista_V_fasorial[hoja2[0][n]] = np.round(V_fasorial, 4)
    #calculamos las corrientes inyectadas
    if V_fasorial != 0 and V_Zeq != 0:
        IV = V_fasorial / V_Zeq
    elif V_fasorial == 0 or V_Zeq == 0:
        IV = 0
    #guardamos en una lista
    lista_IV[hoja2[0][n]] = np.round(IV, 4)
    for a in range(n+1):
        if hoja2[0][n] == a+1:
            listas_MP[a].append(1/V_Zeq)

#calculo para el caso en que hay varias fuentes conectadas al mismo nodo, es decir, en paralelo
for i in range(1, len(hoja2[0])):
    for k in range(i + 1, len(hoja2[0])):
        if hoja2[0][i] == hoja2[0][k]:
            hoja2[1][k] = "Warning!"
            sys.exit("Error de datos en V_fuente")
            
#Operaciones de los elementos para I_fuente (Hoja3)
#Revisamos primero si los valores en las listas no presentan errores

'''for n in range(1, len(I_fuente["Bus i"])+1):
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
        sys.exit("Error de datos CfI")'''

for n in range(1, len(I_fuente["Bus i"])+1):
    #calculo del angulo de desfase
    I_angulo = w * hoja3[3][n]
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
    I_Zc = np.complex_(I_Xc * -1j) #impedancia del capacitor
    I_Zr = (I_Xr) #impedancia del resistor
    #calculo de la impedancia equivalente en serie
    I_Zeq = np.round(I_Zr + I_Zl + I_Zc, 4)
    #guardamos los valores en listas
    lista_IZeq[hoja3[0][n]] = I_Zeq
    #calculo del corriente rms
    Irms = hoja3[2][n] / np.sqrt(2)
    #calculo del corriente en forma fasorial
    I_fasorial = Irms * np.cos(I_angulo) + np.complex_(Irms * np.sin(I_angulo) * 1j)
    I_fasorial = np.round(I_fasorial, 4)
    #guardamos en una lista
    lista_I_fasorial[hoja3[0][n]] = I_fasorial
    lista_IV[hoja3[0][n]] = I_fasorial

#calculo para el caso en que hay varias fuentes conectadas al mismo nodo, es decir, en paralelo
for i in range(1, len(hoja3[0])):
    for k in range(i + 1, len(hoja3[0])):
        if hoja3[0][i] == hoja3[0][k]:
            hoja3[1][k] = "Warning!"
            sys.exit("Error de datos en I_fuente")
            
#Operaciones de los elementos para Z (Hoja4)
#Revisamos primero si los valores en las listas no presentan errores

'''for n in range(1, len(Z["Bus i"])+1):
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
        sys.exit("Error de datos C")'''

for n in range(1, len(Z["Bus i"])+1):
    #Conversiones de uH a H y uF a F
    converion_C = hoja4[5][n] * 10 ** -3
    conversion_L = hoja4[4][n] * 10 ** -3
    #calculo de la reactancia
    ZL = w * conversion_L #reactancia del inductor 
    ZC = w * converion_C #reactancia del capacitor
    if ZC != 0:
        ZC = 1 / ZC #reactancia del capacitor
    ZR = hoja4[3][n] #reactancia del resistor
    #calculo de impedancia
    Z_Zl = np.complex_(ZL * 1j) #impedancia del inductor
    Z_Zc = np.complex_(ZC * -1j) #impedancia del capacitor
    Z_Zr = (ZR) #impedancia del resistor
    #calculo de las impedancias en serie
    Z_Zeq = np.round(Z_Zr + Z_Zl + Z_Zc, 4)
    #guardamos los valores en listas
    lista_Z_Zieq[hoja4[0][n]] = Z_Zeq
    lista_Z_Zkeq[hoja4[1][n]] = Z_Zeq
    for a in range(n+1):
        if hoja4[0][n] == a+1:
            lista_nodo1.append(np.round(1/Z_Zeq,))
            listas_MP[a].append(np.round(1/Z_Zeq, 4))
        if hoja4[1][n] == a+1:
            lista_nodo1.append(np.round(1/Z_Zeq,4))
            listas_MP[a].append(np.round(1/Z_Zeq, 4))

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
                Z_Zeq = np.round(1 / Z_Zeq, 4)
            #sacamos de las listas las impedancias en paralelo
            lista_Z_Zieq.pop(i)
            lista_Z_Zieq[k-1] = np.complex_(0)
            #guardamos los datos en una lista
            lista_Z_Zieq.insert(i, Z_Zeq)

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
                Z_Zeq = np.round(1 / Z_Zeq, 4)
            #sacamos de las listas las impedancias en paralelo
            lista_Z_Zkeq.pop(i)
            lista_Z_Zkeq[k-1] = np.complex_(0)
            #guardamos los datos en una lista
            lista_Z_Zkeq.insert(i, Z_Zeq)

#crear la matriz ybus
ybus_array = np.zeros((dim, dim), dtype = "complex_")

#Diagonal principal
for i in range(dim+1):
    for j in range(dim+1):
        if i == j:
            ybus_array[i-1][j-1] = sum(listas_MP[i-1])
        elif i != j:
            if hoja4[0][i+1] == i and hoja4[1][j+1] == j:
                ybus_array[i-j][j-i] = 0

#Demas elementos de la matriz                
for a in range(1, len(Z["Bus i"])+1):
    if hoja4[0][a] != 0 and hoja4[1][a] != 0:
        ybus_array[hoja4[1][a-1]][hoja4[0][a-1]] = -1*lista_nodo1[a]
        ybus_array[hoja4[0][a-1]][hoja4[1][a-1]] = -1*lista_nodo1[a]
        if lista_nodo1[a] == 0j:
            ybus_array[hoja4[0][a-2]][hoja4[1][a-2]] = -1*lista_nodo1[a+1]
            ybus_array[hoja4[1][a-2]][hoja4[0][a-2]] = -1*lista_nodo1[a+1]
        
#matriz de corrientes
corrientes = np.zeros((dim,1), dtype = "complex_")
for i in range(0, dim):
    corrientes[i][0] = lista_IV[i+1]
    
#Resolver las matrices
tensiones_nodales = np.linalg.solve(ybus_array, corrientes)
tensiones_nodales = np.round(tensiones_nodales, 4)

#Obtenemos el VTH de las tensiones nodales
VTH = tensiones_nodales

#Potencia de las fuentes
for n in range(1, len(V_fuente["Bus i"])+1):
    a = hoja2[0][n]
    S_fuente = np.conjugate(lista_IV[a]) * lista_V_fasorial[a]
    lista_S_fuente[a] = np.round(S_fuente, 4)

print(lista_S_fuente)
#Matris ybuss invertida
zbuss_array = np.linalg.inv(ybus_array)
#obtenemos el zth de la diagonal principal invertida
ZTH = np.diag(zbuss_array)