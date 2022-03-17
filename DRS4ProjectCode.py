import pandas as pd
import numpy as np
from sxl import Workbook
import time
from scipy.optimize import curve_fit
import matplotlib.pyplot as plt
from scipy.signal import chirp, find_peaks, peak_widths

import warnings
warnings.filterwarnings("ignore")




def parabola(x,a,b,c):
	return a*x*x+b*x+c

niz_maximuma = [] #niz sa maximalnim tackama, totalno nebitno samo za debugovanje
metod1 = [] #samo max vrednost
metod2 = [] #srednja vrednost 3 maximuma
metod3 = [] #parabola


file_name = 'SveDRS43.xlsx' #ime fajla 
broj_sheetova=12 #broj sheetova

start = time.time()

MIN_PIK =300 #minimalni kanal za opseg pika
MAX_PIK =550 #maksimalni kanal za opseg pika

ws = Workbook(file_name) #selektovanje fajla
print(" -----------------\nPocetak: " + time.asctime( time.localtime(start) )+"\n -----------------\n" )

baseline=0 #pocetne vrednosti za sumaciju
N=0 #broj merenja

for sheet in range(broj_sheetova):
	start_sheeta = time.time()
	
	ime_sheeta = str(sheet) #konverzija broja u string jer se tako zovu sheetovi
	
	wb = ws.sheets[ime_sheeta] #ucitavanje sheeta
	all_rows = list(wb.rows) #kreira se lista nizova od sheeta
	df = pd.DataFrame(all_rows) #konvertovanje u dataframe da bi prebacili u numpy 
	res = df.to_numpy() #konverzija u numpy matricu
	merenja = np.transpose(res) #transponovanje matrice da bi se dobilo N nizova od 1024 merenja 
	
	N = N + len(merenja) #brojac merenja
	
	for i in range(len(merenja)): #listanje svih merenja u jednom sheetu
		max_ind = np.argmax(merenja[i][MIN_PIK:MAX_PIK])+MIN_PIK #maximum u opsegu 
		
		sirina_pika = peak_widths(merenja[i], [max_ind], rel_height=0.95)[0] #racunanje sirine pika pod pretpostavkom da je 5% max vrednosti od FONA od svih grafika
		mosa = max_ind - int(sirina_pika/2)-100 #na kom kanalu pocinje rast pika i oduzima se od njega 100 jer se dotle gleda FON za baznu liniju
		baseline=baseline+np.average(merenja[i][0:mosa]) #bazna linija je srednja vrednost svih vrednosti do kanala MOSA, i sumiraju se sve da bi se usrednjile na kraju

		indeksi = [max_ind,max_ind-1,max_ind+1] #niz indeksa 3 maksimalne tacke
		
		niz_maximuma.append([ [indeksi[0],merenja[i][indeksi[0]]], #nbt stvar
					[indeksi[1],merenja[i][indeksi[1]]], #samo za debugovanje koristi
					[indeksi[2],merenja[i][indeksi[2]]] ])
					
		metod1.append(merenja[i][indeksi[0]]) #dodavanje maximuma za prvu metodu
		metod2.append((merenja[i][indeksi[0]]+merenja[i][indeksi[1]]+merenja[i][indeksi[2]])/3) #dodavanje maximuma za drugu metodu
		
		#dodavanje maximuma za trecu metodu
		popt,pcov = curve_fit(parabola,[indeksi[0],indeksi[1],indeksi[2] ], 
					[ merenja[i][indeksi[0]],merenja[i][indeksi[1]],merenja[i][indeksi[2]] ]) #fitovanje parabole na tri maksimuma
		metod3.append(parabola(-popt[1]/(2*popt[0]) ,*popt)) #kalkulacija maksimuma fitovane parabole

	end_sheeta = time.time()
	eta = end_sheeta+ int(end_sheeta - start_sheeta)*(broj_sheetova-sheet-1) #racunanje za ocekivano vreme zavrsetka, nbt 
	print("Vreme za " + ime_sheeta + ". sheet: " + str(int(end_sheeta - start_sheeta)) + "s             ETA~ " + time.asctime( time.localtime(eta) ) )
	
x1 = [] #lista za prvi metod
x2 = [] #lista za drugi metod
x3 = [] #lista za treci metod

baseline = baseline/N #usrednjavanje bazne linije iz svih merenja
print("Bazna linija: " + str(baseline))

for i in range(len(metod1)): #kreiranje novih x osa za sve metode i takodje se od svake vrednosti pika oduzima bazna linija
        if 0<(int(512*(metod1[i]-baseline))) and (int(512*(metod1[i]-baseline)))<228:
                x1.append(int(512*(metod1[i]-baseline)))
        if 0<(int(512*(metod2[i]-baseline))) and (int(512*(metod2[i]-baseline)))<220:
                x2.append(int(512*(metod2[i]-baseline)))
        if 0<(int(512*(metod3[i]-baseline))) and (int(512*(metod3[i]-baseline)))<225:
                x3.append(int(512*(metod3[i]-baseline)))
print(len(metod1))
x1=np.array(x1) #konverzija u numpy array 
x2=np.array(x2) #da bi mogle vrednosti lakse da se analiziraju
x3=np.array(x3)
x1_p, y1_p = np.unique(x1, return_counts=True) #brisanje duplikata i stavljanje svih razlicith vrednosti na x osu
x2_p, y2_p = np.unique(x2, return_counts=True) #dok se na y osu stavlja broj ponavljanja duplikata za svaku unikatnu vrednost
x3_p, y3_p = np.unique(x3, return_counts=True)
print("-----------------\n")
end = time.time()
print("Kraj: " + time.asctime( time.localtime(end) ) + "\n")
print("Vreme rada programa: " + str(int(end-start)) + " s\n")

plt.figure(1)
plt.plot(x1_p, y1_p , label = "Metod najvise vrednosti", color='r')
#plt.savefig('MetodNajviseVrednosti.png')
plt.xlabel('Kanal')
plt.ylabel('Broj Amplituda')
plt.legend()
plt.show()

plt.figure(2)
plt.plot(x2_p, y2_p,  label = "Metod srednje vrednosti", color='b')
#plt.savefig('metodSrednjeVrednosti.png')
plt.xlabel('Kanal')
plt.ylabel('Broj Amplituda')
plt.legend()
plt.show()

plt.figure(3)
plt.plot(x3_p, y3_p,  label = "Metod parabole", color='g')
#plt.savefig('MetodParabole.png')
plt.xlabel('Kanal')
plt.ylabel('Broj Amplituda')
plt.legend()
plt.show()

plt.figure(4)
plt.plot(x1_p, y1_p , label = "Metod najvise vrednosti", color='r')
plt.plot(x2_p, y2_p,  label = "Metod srednje vrednosti", color='b')
plt.plot(x3_p, y3_p,  label = "Metod parabole", color='g')
plt.xlabel('Kanal')
plt.ylabel('Broj Amplituda')
plt.legend()
plt.show()
