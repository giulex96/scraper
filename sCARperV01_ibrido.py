import requests, openpyxl
from bs4 import BeautifulSoup
import re
import string
import xlsxwriter
import numpy as np
import time
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from itertools import chain

import xlwt 
from xlwt import Workbook 

now = datetime.now()
A_A1time = now.strftime("%d/%m/%Y %H:%M:%S")

#crea excel
# Workbook is created 
wb = Workbook()   
count=0
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Sheet 1') 
sheet1.write(count,  0, 'DateTime')
sheet1.write(count,  1, 'Model1')
sheet1.write(count,  2, 'Model2')
sheet1.write(count,  3, 'Automaker')
sheet1.write(count,  4, 'Version')
sheet1.write(count,  5, 'Price')
sheet1.write(count,  6, 'YearMon')
sheet1.write(count,  7, 'Year')
sheet1.write(count,  8, 'Km')
sheet1.write(count,  9, 'VehicleType1')
sheet1.write(count,  10, 'VehicleType2')
sheet1.write(count,  11, 'VehicleUsed1')
sheet1.write(count,  12, 'VehicleUsed1')
sheet1.write(count,  13, 'Owners')
sheet1.write(count,  14, 'Horses')
sheet1.write(count,  15, 'Fuel')
sheet1.write(count,  16, 'GearingType')
sheet1.write(count,  17, 'Gear')
sheet1.write(count,  18, 'Displacement')
sheet1.write(count,  19, 'Cylinder')
sheet1.write(count,  20, 'DriveChain')
sheet1.write(count,  21, 'Weight')
sheet1.write(count,  22, 'BeltSub')
sheet1.write(count,  23, 'ColorExternal')
sheet1.write(count,  24, 'ColorOriginal')
sheet1.write(count,  25, 'PaintType')
sheet1.write(count,  26, 'FuelCombo')
sheet1.write(count,  27, 'FuelUrban')
sheet1.write(count,  28, 'FuelExtra')
sheet1.write(count,  29, 'Emissions')
sheet1.write(count,  30, 'Label')
sheet1.write(count,  31, 'Smoker')
sheet1.write(count,  32, 'Certified')
sheet1.write(count,  33, 'Country')
sheet1.write(count,  34, 'NewLicense')
sheet1.write(count,  35, 'Doors')
sheet1.write(count,  36, 'NextInspection')
sheet1.write(count,  37, 'PastInspection')
sheet1.write(count,  38, 'CertifiedInspectio')
sheet1.write(count,  39, 'Upholstery')
sheet1.write(count,  40, 'Seats')
sheet1.write(count,  41, 'PriceSuper')
sheet1.write(count,  42, 'PriceOttimo')
sheet1.write(count,  43, 'PriceBuono')
sheet1.write(count,  44, 'PriceGrey1')
sheet1.write(count,  45, 'PriceGrey2')
sheet1.write(count,  46, 'Street')
sheet1.write(count,  47, 'City')
sheet1.write(count,  48, 'Country')
sheet1.write(count,  49, 'Description')
sheet1.write(count,  50, 'UrlAuto')
sheet1.write(count,  51, 'UrlSearch')


####################################################################################################
#START SCRAPER
#take a single ZIP code
#for c in range(0,len(cap)):
    
    
    

for p in range(1,21):
    print(p)
   
    url1 = 'https://www.autoscout24.it/lst/?sort=standard&desc=0&offer=S%2CU&ustate=N%2CU&size=20&page='
      
    pag = str(p)
      
    url2 = '&lon=8.95147&lat=45.59314&zip=20023%20Cerro%20Maggiore&zipr=4&cy=I&kmfrom=40000&atype=C&fc=120&qry=&recommended_sorting_based_id=2fb51e98-81cd-4c66-bd01-d610a775bb62&'
    
   
    url = url1 + pag + url2
    
    tries = 2
    
    
    
    while True:
        try:
            time.sleep(1)
            page = requests.get(url)  
            soup = BeautifulSoup(page.content, 'html.parser')
            car_n = soup.find_all("div", {"class": "cl-list-item-gap"})
            car_n=str(car_n)
            car_n = float(re.sub("[^0-9]", "", car_n)) 
        except ValueError:
            time.sleep(1) 
            tries -= 1
            if tries:  # if tries != 0
                continue  # not necessary, just for clarity
            else:
                raise  
        else:
            break  
    
    
    #find all vehicle links in the page    
    for link in soup.find_all("a", attrs={"data-item-name":"detail-page-link"}):
        car_url = 'https://www.autoscout24.it' + link.attrs['href']
        bs = BeautifulSoup(requests.get(car_url).text)
        
        check = bs.find('span', {'class':'cldt-detail-makemodel sc-ellipsis'})
        if check is not None:
            A_model = bs.find('span', {'class':'cldt-detail-makemodel sc-ellipsis'}).text
        else:
            A_model="999999"

        check = bs.find('span', {'class':'cldt-detail-version sc-ellipsis'})
        if check is not None:
            A_version = bs.find('span', {'class':'cldt-detail-version sc-ellipsis'}).text
        else:
            A_version="999999"

        check = bs.find('h4', {'class':'cldt-detail-subheadline sc-font-m sc-ellipsis'})
        if check is not None:
            A_shape = bs.find('h4', {'class':'cldt-detail-subheadline sc-font-m sc-ellipsis'}).text
        else:
            A_shape="999999"

        check = bs.find('div', {'class':'cldt-price'})
        if check is not None:
            A_price = bs.find('div', {'class':'cldt-price'}).text.strip()
            A_price = re.sub("[^0-9,]", "", A_price)
        else:
            A_price="999999"

        check = bs.find('span', {'class':'sc-font-l cldt-stage-primary-keyfact'})
        if check is not None:
            A_km = bs.find('span', {'class':'sc-font-l cldt-stage-primary-keyfact'}).text
        else:
            A_km="999999"

        check = bs.find('span', {'class':'sc-font-s cldt-stage-att-description'})
        if check is not None:
            A_usato =  bs.find('span', {'class':'sc-font-s cldt-stage-att-description'}).text
        else:
            A_usato="999999"

        check = bs.find('span', {'class':'sc-font-l cldt-stage-primary-keyfact' , 'id':'basicDataFirstRegistrationValue'})
        if check is not None:
            A_date = bs.find('span', {'class':'sc-font-l cldt-stage-primary-keyfact' , 'id':'basicDataFirstRegistrationValue'}).text
        else:
            A_date="999999"

        check = bs.find('span', {'class': 'sc-font-m cldt-stage-primary-keyfact'})
        if check is not None:
            A_cv = bs.find('span', {'class': 'sc-font-m cldt-stage-primary-keyfact'}).text
        else:
            A_cv="999999"
        
                
        ####################################################################################################################################################################################
        ####################################################################################################################################################################################
        #dati principali 
        a_cara = bs.find('div', {'class':'cldt-item' , 'data-item-name':'car-details'}).text
        a_cara = re.sub("I dati di consumi ed emissioni per le auto usate si intendono riferiti al ciclo NEDC. Per le auto nuove, a partire dal 16.2.2021, iI rivenditore deve indicare i valori relativi al consumo di carburante ed emissione di CO2 misurati con il ciclo WLTP. Il rivenditore deve rendere disponibile nel punto vendita una guida gratuita su risparmio di carburante e emissioni di CO2 dei nuovi modelli di autovetture. Anche stile di guida e altri fattori non tecnici influiscono su consumo di carburante e emissioni di CO2. Il CO2 è il gas a effetto serra principalmente responsabile del riscaldamento terrestre.", "", a_cara)
        a_cara = a_cara.split('\n')
        a_cara = list(filter(None, a_cara))
        
        #tipo di veicolo
        if "Tipo di veicolo" not in a_cara:
            a_type = "999999"
        else:
            a_type = a_cara[a_cara.index("Tipo di veicolo")+1]
            
        #neo pat
        if "Per neopatentati" not in a_cara:
            a_neo = "999999"
        else:
            a_neo = a_cara[a_cara.index("Per neopatentati")+1]
        
        #tagliando
        if "Ultimo tagliando" not in a_cara:
            a_taglian = "999999"
        else:
            a_taglian = a_cara[a_cara.index("Ultimo tagliando")+1]
        
        #fumo
        if "Veicolo per non fumatori" not in a_cara:
            a_fumo = "999999"
        else:
            a_fumo = "Y"
            
        #tagliandi certificati
        if "Tagliandi certificati" not in a_cara:
            a_taglicert = "999999"
        else:
            a_taglicert = "Y"
        
        #cinghia
        if "Ultimo cambio cinghia distribuzione" not in a_cara:
            a_cinghia = "999999"
        else:
            a_cinghia = a_cara[a_cara.index("Ultimo cambio cinghia distribuzione")+1]
        
        #Usato Garantito
        if "Usato Garantito" not in a_cara:
            a_garantito = "999999"
        else:
            a_garantito = a_cara[a_cara.index("Usato Garantito")+1]
        
        #Prossima revisione
        if "Prossima revisione" not in a_cara:
            a_proxrev = "999999"
        else:
            a_proxrev = a_cara[a_cara.index("Prossima revisione")+1]
        
        #Proprietari
        if "Proprietari" not in a_cara:
            a_prop = "999999"
        else:
            a_prop = a_cara[a_cara.index("Proprietari")+1]
        
        #marca
        if "Marca" not in a_cara:
            a_marca = "999999"
        else:
            a_marca = a_cara[a_cara.index("Marca")+1]
        
        #modello
        if "Modello" not in a_cara:
            a_modello = "999999"
        else:
            a_modello = a_cara[a_cara.index("Modello")+1]
        
        #Anno
        if "Anno" not in a_cara:
            a_anno = "999999"
        else:
            a_anno = a_cara[a_cara.index("Anno")+1]
            
        #Colore esterno
        if "Colore esterno" not in a_cara:
            a_colest = "999999"
        else:
            a_colest = a_cara[a_cara.index("Colore esterno")+1]
        
        #Colore originale
        if "Colore originale" not in a_cara:
            a_colorig = "999999"
        else:
            a_colorig = a_cara[a_cara.index("Colore originale")+1]
            
        #Rivestim
        if "Rivestimenti" not in a_cara:
            a_rives = "999999"
        else:
            a_rives = a_cara[a_cara.index("Rivestimenti")+1]
            
        #Carrozza
        if "Carrozzeria" not in a_cara:
            a_carroz = "999999"
        else:
            a_carroz = a_cara[a_cara.index("Carrozzeria")+1]
        
        #tipo di vernice
        if "Tipo di vernice" not in a_cara:
            a_vernice = "999999"
        else:
            a_vernice = a_cara[a_cara.index("Tipo di vernice")+1]
        
        #Porte
        if "Porte" not in a_cara:
            a_porte = "999999"
        else:
            a_porte = a_cara[a_cara.index("Porte")+1]
        
        #Posti a sedere
        if "Posti a sedere" not in a_cara:
            a_sedute = "999999"
        else:
            a_sedute = a_cara[a_cara.index("Posti a sedere")+1]
        
        #Posti a sedere
        if "Versione per nazione" not in a_cara:
            a_nazio = "999999"
        else:
            a_nazio = a_cara[a_cara.index("Versione per nazione")+1]
        
        #Tipo di cambio
        if "Tipo di cambio" not in a_cara:
            a_cambio = "999999"
        else:
            a_cambio = a_cara[a_cara.index("Tipo di cambio")+1]
        
        #Marce
        if "Marce" not in a_cara:
            a_marce = "999999"
        else:
            a_marce = a_cara[a_cara.index("Marce")+1]
        
        #Cilindrata
        if "Cilindrata" not in a_cara:
            a_cili = "999999"
        else:
            a_cili = a_cara[a_cara.index("Cilindrata")+1]
        
        #Cilindri
        if "Cilindri" not in a_cara:
            a_cilindri = "999999"
        else:
            a_cilindri = a_cara[a_cara.index("Cilindri")+1]
        
        #Peso a vuoto
        if "Peso a vuoto" not in a_cara:
            a_peso = "999999"
        else:
            a_peso = a_cara[a_cara.index("Peso a vuoto")+1]
            
        #Tipo di unità
        if "Tipo di unità" not in a_cara:
            a_unita = "999999"
        else:
            a_unita = a_cara[a_cara.index("Tipo di unità")+1]
            
        #Alimentazione
        if "Alimentazione" not in a_cara:
            a_alimenta = "999999"
        else:
            a_alimenta = a_cara[a_cara.index("Alimentazione")+1]
        
        #Consumo carburante
        if "Consumo carburante:" not in a_cara:
            a_consumcomb = "999999"
        else:
            a_consumcomb = a_cara[a_cara.index("Consumo carburante:")+1]
        
        #Consumo carburante
        try:
         if "Consumo carburante:" not in a_cara:
            a_consumurb = "999999"
         else:
            a_consumurb = a_cara[a_cara.index("Consumo carburante:")+2]
        except IndexError:
         pass
        
        #Consumo carburante
        try:
         if "Consumo carburante:" not in a_cara:
            a_consumextra = "999999"
         else:
            a_consumextra = a_cara[a_cara.index("Consumo carburante:")+3]
        except IndexError:
         pass     
    
        
        #Emissioni di CO2
        if "Emissioni di CO2" not in a_cara:
            a_emis = "999999"
        else:
            a_emis = a_cara[a_cara.index("Emissioni di CO2")+1]
            
        #Classe emissioni
        if "Classe emissioni" not in a_cara:
            a_euro = "999999"
        else:
            a_euro = a_cara[a_cara.index("Classe emissioni")+1]
            
        ####################################################################################################################################################################################
        ####################################################################################################################################################################################
        #equipaggiamento
        
        check = bs.find('div', {'class':'cldt-item' , 'data-item-name':'equipments'})
        if check is not None:
            equip = bs.find('div', {'class':'cldt-item' , 'data-item-name':'equipments'}).text
            equip = re.sub("Equipaggiamento", "", equip)
            equip = re.sub("Comfort", "", equip)
            equip = re.sub("Intrattenimento / Media", "", equip)
            equip = re.sub("Extra", "", equip)
            equip = re.sub("Sicurezza", "", equip)
            equip = equip.split('\n')
            equip = list(filter(None, equip))
            
            str1 = "|"  
            e_equip = ""
            for ele in equip:  
                e_equip = e_equip + ele + str1            
        else:
            equip = "999999"
        
        
        ####################################################################################################################################################################################
        ####################################################################################################################################################################################
        #valutazione del prezzo
        check = bs.find('ul', {'class':'pe-visualization__bar'})
        if check is not None:
            valut = bs.find('ul', {'class':'pe-visualization__bar'}).text
            valut = valut.split('\n')
            valut = list(filter(None, valut))
            #super prezzo
            if "Super prezzo" not in valut:
                p_super = "999999"
            else:
                p_super = valut[valut.index("Super prezzo")+1]
            
            #ottimo prezzo
            if "Ottimo prezzo" not in valut:
                p_ottimo = "999999"
            else:
                p_ottimo = valut[valut.index("Ottimo prezzo")+1]
            
            #ottimo prezzo
            if "Buon prezzo" not in valut:
                p_buon = "999999"
            else:
                p_buon = valut[valut.index("Buon prezzo")+1]
            
            #grigio 1
            if "Buon prezzo" not in valut:
                p_grigio1 = "999999"
            else:
                p_grigio1 = valut[valut.index("Buon prezzo")+2]
            
            #grigio 2
            if "Buon prezzo" not in valut:
                p_grigio2 = "999999"
            else:
                p_grigio2 = valut[valut.index("Buon prezzo")+3]
        
        else:
            p_super = "999999"
            p_ottimo = "999999"
            p_buon = "999999"
            p_grigio1 = "999999"
            p_grigio2 = "999999"
            
            
        ####################################################################################################################################################################################
        ####################################################################################################################################################################################

        
        check = bs.find('div', {'class':'cldt-item' , 'data-item-name':'description'})
        if check is not None:
            descr = bs.find('div', {'class':'cldt-item' , 'data-item-name':'description'}).text
        else:
            descr = "999999"
        
        ####################################################################################################################################################################################
        ####################################################################################################################################################################################
        # Info venditore
        
        check = bs.find('h3', {'class':'sc-font-bold sc-font-m'})
        if check is not None:
            v_company = bs.find('h3', {'class':'sc-font-bold sc-font-m'}).text
        else:
            v_company = "999999"
        
        
        check = bs.find('div', {'data-item-name':'vendor-contact-name'})
        if check is not None:
            v_name = bs.find('div', {'data-item-name':'vendor-contact-name'}).text
        else:
            v_name = "999999"
        
        check = bs.find('div', {'data-item-name':'vendor-contact-street'})
        if check is not None:
            v_street = bs.find('div', {'data-item-name':'vendor-contact-street'}).text
        else:
            v_street = "999999"
        
        check = bs.find('div', {'data-item-name':'vendor-contact-city'})
        if check is not None:
            v_city = bs.find('div', {'data-item-name':'vendor-contact-city'}).text
        else:
            v_city = "999999"
            
        check = bs.find('div', {'data-item-name':'vendor-contact-country'})
        if check is not None:
            v_country = bs.find('div', {'data-item-name':'vendor-contact-country'}).text
        else:
            v_country = "999999"
        
        check = bs.find('a', {'data-type':'callLink'})
        if check is not None:
            v_tel = bs.find('a', {'data-type':'callLink'}).text
        else:
            v_tel = "999999"
        
        #riga,colonna
        sheet1.write(count+1, 0, A_A1time)
        sheet1.write(count+1, 1, A_model)
        sheet1.write(count+1, 2, a_modello)
        sheet1.write(count+1, 3, a_marca)
        sheet1.write(count+1, 4, A_version)
        sheet1.write(count+1, 5, A_price)
        sheet1.write(count+1, 6, A_date)
        sheet1.write(count+1, 7, a_anno)
        sheet1.write(count+1, 8, A_km)
        sheet1.write(count+1, 9, A_shape)
        sheet1.write(count+1, 10, a_carroz)
        sheet1.write(count+1, 11, A_usato)
        sheet1.write(count+1, 12, a_type)
        sheet1.write(count+1, 13, a_prop)
        sheet1.write(count+1, 14, A_cv)
        sheet1.write(count+1, 15, a_alimenta)
        sheet1.write(count+1, 16, a_cambio)
        sheet1.write(count+1, 17, a_marce)
        sheet1.write(count+1, 18, a_cili)
        sheet1.write(count+1, 19, a_cilindri)
        sheet1.write(count+1, 20, a_unita)
        sheet1.write(count+1, 21, a_peso)
        sheet1.write(count+1, 22, a_cinghia)
        sheet1.write(count+1, 23, a_colest)
        sheet1.write(count+1, 24, a_colorig)
        sheet1.write(count+1, 25, a_vernice)
        sheet1.write(count+1, 26, a_consumcomb)
        try:
         sheet1.write(count+1, 27, a_consumurb) 
        except NameError:
         pass
        
        sheet1.write(count+1, 28, a_consumextra)
        sheet1.write(count+1, 29, a_emis)
        sheet1.write(count+1, 30, a_euro)
        sheet1.write(count+1, 31, a_fumo)
        sheet1.write(count+1, 32, a_garantito)
        sheet1.write(count+1, 33, a_nazio)
        sheet1.write(count+1, 34, a_neo)
        sheet1.write(count+1, 35, a_porte)
        sheet1.write(count+1, 36, a_proxrev)
        sheet1.write(count+1, 37, a_taglian)
        sheet1.write(count+1, 38, a_taglicert)
        sheet1.write(count+1, 39, a_rives)
        sheet1.write(count+1, 40, a_sedute)
        sheet1.write(count+1, 41, p_super)
        sheet1.write(count+1, 42, p_ottimo)
        sheet1.write(count+1, 43, p_buon)
        sheet1.write(count+1, 44, p_grigio1)
        sheet1.write(count+1, 45, p_grigio2)
        sheet1.write(count+1, 46, v_street)
        sheet1.write(count+1, 47, v_city)
        sheet1.write(count+1, 48, v_country)   
        sheet1.write(count+1, 49, descr)
        sheet1.write(count+1, 50, car_url)
        sheet1.write(count+1, 51, url)
                
                
        wb.save('example1.xls') 
    
        
        
        count = count+1

        
        