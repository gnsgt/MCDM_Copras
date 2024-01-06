import pandas as pd
import numpy as np

    #Excel dosyasını oku ve DataFrame'e dönustur
df_excel = pd.read_excel("C:/Users/gokha/OneDrive/Masaüstü/İstatistik/Tez/4.1-Copras.xlsx")
maliyet = df_excel.iloc[1:2].copy()
agirlik = df_excel.iloc[:1].copy()
df = df_excel.iloc[2:]  

    #Kac satır kaç sutun?
satir_sayisi = df.shape[0] 
sutun_sayisi = df.shape[1]


    #Toplamlar
toplamlar = []
for j in range(0,sutun_sayisi):
    toplam = 0
    for i in range(0,satir_sayisi):
        toplam = toplam + df.iloc[i,j]
    toplamlar.append(toplam)
    
    #Islem II
df2 = df.copy()
for j in range(0,sutun_sayisi):
    for i in range(0,satir_sayisi):
         df2.iloc[i,j] = (df.iloc[i,j] * agirlik.iloc[0,j]) / toplamlar[j]

    
    #Si'ler
si_plus = []
si_minus = []
for i in range(0,satir_sayisi):
    si_plus_toplam = 0
    si_minus_toplam = 0
    for j in range(0,sutun_sayisi):
        if maliyet.iloc[0,j] == "fayda":
            si_plus_toplam = df2.iloc[i,j] + si_plus_toplam
        else:
            si_minus_toplam = df2.iloc[i,j] + si_minus_toplam
    si_plus.append(si_plus_toplam)
    si_minus.append(si_minus_toplam)


si_minus_min = min(si_minus)
simin_split_siminus = []
for i in range(satir_sayisi):
    simin_split_siminus.append(si_minus_min / si_minus[i])



si_plus_toplam2 = sum(si_plus)
si_minus_toplam2 = sum(si_minus)
simin_split_siminus_toplam = sum(simin_split_siminus)

qi = []
for i in range(satir_sayisi):
    qi.append(si_plus[i] + ((si_minus_min * si_minus_toplam2) / (si_minus[i] * simin_split_siminus_toplam)))
    
qi_max = max(qi) 

ni = []
for i in range(satir_sayisi):
    ni.append(qi[i] / qi_max * 100)
    
print("\n")
for i in range(satir_sayisi):
    max_deger = max(ni)
    max_index = ni.index(max_deger)
    print(f"COPRAS Ni: {max_deger}, İndex: {max_index + 1}")
    ni[max_index] = -99999

    
