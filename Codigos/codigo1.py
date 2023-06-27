import pandas as pd
import numpy as np
import math
from fractions import Fraction

#revisamos el excel con pandas
dat = "data_io_copia.xlsx"
datos = pd.ExcelFile(dat)
print(datos.sheet_names, '\n')

datos1 = pd.read_excel(dat, header=None)
#print(datos1)
V_fuente2 = pd.read_excel(dat, "V_fuente", header=None)
#print(V_fuente2)
#print(V_fuente2[3][3])
#print(V_fuente2[3] [len(V_fuente["Bus i"])-1])

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
if datos1[1][0] == 60: #si la frecuncia es 60Hz, el valor de w es directamente 60
    w = 377
else: #si la frecuencia no es 60, entonces se calcula manualmente el valor de w y se redondea a su entero mas cercano
    w = round(2 * math.pi * datos1[1][0], 4)
    #print(w) #usamos round(variable, 4) para redondear el calculo y que se queden 4 decimales

#variable para guardar los datos en un array
num = 0

for i in range(len(V_fuente["Bus i"]) - 1):
    for k in range(i + 1, len(V_fuente["Bus i"])):
        #if(V_fuente["Bus i"][i] == V_fuente["Bus i"][k]):   #USAR DESPUES PARA EL CASO EN QUE LOS ELEMENTOS ESTEN EN EL MISMO NODO, ES DECIR EN PARALELO
            #print("Frencuencia de operacion: %sHz" %(datos1[1][0]))  #LO ANTERIOR
            #print("Frecuencia angular: %s" %w)                         #LO ANTERIOR X2
            for n in range(1, len(V_fuente["Bus i"])+1):
                #print("Tiempo de onda %d: %sseg" %(n, V_fuente2[3][n]))
                #calculos el angulo de desfase
                angulo = round(w * V_fuente2[3][n], 4)
                #print("El angulo para To%i: %s" %(n, angulo))
                #Conversiones
                converion_Cf = round(V_fuente2[25][n] * 10 ** -3, 4)
                conversion_Lf = round(V_fuente2[5][n] * 10 ** -3, 4)
                #calculo de la reactancia
                V_Xl = round(w * conversion_Lf, 4) #reactancia del inductor
                V_Xc = round(w * converion_Cf, 4) #reactancia del capacitor
                V_Xc = round(1 / V_Xc, 4) #reactancia del capacitor
                V_Xr = V_fuente2[4][n] #reactancia del resistor
                #calculo de impedancia
                V_Zl = np.complex_(V_Xl * 1j) #impedancia del inductor
                V_Zc = np.complex_(V_Xc * 1j) #impedancia del capacitor
                V_Zr = np.complex_(V_Xr * 1j) #impedancia del resistor
                #print(V_Zl)
                #print(V_Zc)
                #print(V_Zr)
                #print()
                #calculo de voltaje rms
                Vrms = round(V_fuente2[2][n] / np.sqrt(2), 4)
                #print(Vrms)
                #calculo de voltaje rms en forma fasorial (CREO QUE ES ASI)
                Vf_rms = (Vrms * np.cos(angulo) + Vrms * np.sin(angulo))
                #print(round(Vf_rms, 4))
                #num = num + 1

            



#imprimimos las tablas para comprobar datos
#print(V_fuente["Bus i"], '\n')
#print("prueba \n", f_and_ouput1["Warning"], "\n")
#print("Tabla V_ fuente, Columna Warning \n", V_fuente["Warning"], "\n")
#imprimimos las tablas para comprobar datos1
#print(datos1)
#print(V_fuente["Vpico f (V)"], '\n')
#print(V_fuente["Cf (uF)"])
#print(V_fuente["Lf (mH)"])
#print(V_fuente["Rf (ohms)"])