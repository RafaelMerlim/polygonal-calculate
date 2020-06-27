import math
import pandas as pd

# Cálculo do azimute
def azimute(ang_ex):
    n = 1
    azimutes = [ang_ex[0]]
    
    while n < len(ang_ex):
        azq_t = (ang_ex[n] - 180) + azimutes[n-1]
        if azq_t > 360:
            azimutes.append(azq_t - 360)
        elif azq_t <0:
            azimutes.append(azq_t + 360)
        else:
            azimutes.append(azq_t)
        
        n+=1
            
    return azimutes

# Cálculo dos deltas
def delta(azimutes, distanciasv):
    n = 0
    d_aprox = []
    f_cor = []
    d_cor = []

# Calculando o delta aproximado.
    while n < len(azimutes):
        
        if azimutes[n] >= 0 and azimutes[n] <= 90:
            de = distanciasv[n]*math.sin(math.radians(azimutes[n]))
            dn = distanciasv[n]*math.cos(math.radians(azimutes[n]))
        elif azimutes[n] > 90 and azimutes[n] <= 180:
            de = distanciasv[n]*math.sin(math.radians(180 - azimutes[n]))
            dn = (-1)*(distanciasv[n]*math.cos(math.radians(180 - azimutes[n])))
        elif azimutes[n] > 180 and azimutes[n] <= 270:
            de = (-1)*(distanciasv[n]*math.sin(math.radians(azimutes[n] - 180)))
            dn = (-1)*(distanciasv[n]*math.cos(math.radians(azimutes[n] - 180)))
        else:
            de = (-1)*(distanciasv[n]*math.sin(math.radians(360 - azimutes[n])))
            dn = distanciasv[n]*math.cos(math.radians(360 - azimutes[n]))
            
        d_aprox.append((de, dn))
        n+=1

# Calculando o delta corrigido.
    fe = 0    
    fn = 0
    sum_de_abs = 0
    sum_dn_abs = 0
    for par in d_aprox:
        fe = fe + par[0]
        fn = fn + par[1]
        sum_de_abs = sum_de_abs + math.fabs(par[0])
        sum_dn_abs = sum_dn_abs + math.fabs(par[1])
    
    f_de_cor = 0
    f_dn_cor = 0
    for par in d_aprox:
        f_de_cor = ((fe*math.fabs(par[0]))/(sum_de_abs))*(-1)
        f_dn_cor = ((fn*math.fabs(par[1]))/(sum_dn_abs))*(-1)
        f_cor.append((f_de_cor, f_dn_cor))
        
    k = 0
    
    while k < len(f_cor):
        d_cor.append((d_aprox[k][0] + f_cor[k][0], d_aprox[k][1] + f_cor[k][1]))
        k+=1
        
    # Exportando deltas corrigidos e erro de fechamento.
    return d_cor, math.sqrt(fe**2 + fn**2)
        

# Cálculo das coordenadas.
def coordenada(coordi, deltas):
    coord = coordi
    n = 0
    
    while n < len(deltas):
        coord.append((coord[n][0]+deltas[n][0], coord[n][1]+deltas[n][1]))
        n+=1
    return coord
    

# Importando planilha com dados do levantamento.
#diretorio = input('Insira o diretório juntamente com o nome do arquivo: ')
#arquivo = input('Insira o nome do arquivo')

diretorio = 'C:\\Users\\rafam\\Documents\\1_CURSOS\\python_3\\plan_teste.xlsx'
df = pd.read_excel(diretorio)

azq = azimute(list(df['AE']))

dd, fecha = delta(azq, list(df['DPV']))
print(dd)
print(fecha)

coord_final = coordenada([(df['E'][0], df['N'][0])], dd)
print(coord_final)