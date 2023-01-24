# -*- coding: utf-8 -*-
"""
Created on Tue Jul  9 15:19:09 2019

Letzte Änderungen
27.11.19: Belichtungszeit richtig angeben
28.11.19: Convert Problem, Index von 5 auf 4 gesetzt
3.8.20: Mit Unterordner und PP-Name nach Ordner
10.11.20 Anscheinend neue EXIF Formatierung.
24.1.21 kleiner Bugfix
17.1.23 Datum statt Kameradaten; Ortsinformationen
"""

"""
Hilfreiche Links:
https://www.shibutan-bloomers.com/python-libraly-pptx-2_en/6310/
"""


# fuer pptx und exifread: pip install python-pptx und pip install exifread

from pptx import Presentation
from pptx.util import Inches,Cm
import os
from PIL import Image
from PIL.ExifTags import TAGS
from pptx.util import Pt

from geopy.geocoders import Nominatim
geoLoc = Nominatim(user_agent="GetLoc")


Ordner="Fotoabend2022"
Mit_UnterOrdner=False
Datum_Statt_Kamera_Daten=True
Overview_Slide=False
neu_konvertieren=True # muss bei erster Ausführung true sein
karte_ausgeben=True
fontsize=12

#Korrekte Umrechnung Zeiten
def belichtungszeit(zeit):
    if zeit<1:
        zeit=int(1/zeit)
        zeit="1/"+str(zeit)
    return zeit

#Minuten und Sekunden addieren
def gps_converter(North,East):
    North_Dezi=North[0]+North[1]/60.+North[2]/3600.
    East_Dezi = East[0]+ East[1]/60.+ East[2]/3600.
    return North_Dezi,East_Dezi


#Exifdaten auslesen
def get_exif(fn):
    print("Pfad für EXIF: ",fn)
    ret = {}
    i = Image.open(fn)
    info = i._getexif()  
    keys = list(info.keys())
    keys = [k for k in keys if k in TAGS]
    for tag, value in info.items():
        decoded = TAGS.get(tag, tag)
        ret[decoded] = value
    return ret

#Bildgröße berechnen
def px_to_inches(path):    
    im = Image.open(path)
    width = im.width 
    height = im.height 
    #print("Image Info",im.info,"breite",width,"höhe",height)
    return (width, height)

#Datumsformat von Exif verbessern
def change_date_format(date):
    zwi=date.split(" ")
    datum=zwi[0].split(":")
    zeit=zwi[1]
    return datum[2]+"-"+datum[1]+"-"+datum[0]+" "+zeit

# Lösche alte Powerpointdatei, da nicht überschrieben wird
try:
    os.remove('Bilder_Fototron'+Ordner+'.pptx')
except:
    print("Keine Datei zum Löschen")

prs = Presentation()



if Mit_UnterOrdner:
    JPG_Dateien=[]
    for sub_ordner in os.listdir(Ordner):
        print (sub_ordner)
        zwi=(os.listdir(Ordner+"/"+sub_ordner))
        for Dateien in zwi:
            JPG_Dateien.append(sub_ordner+"/"+Dateien)
    print(JPG_Dateien)

else:
    #Liste alle Dateien im Ordner aus
    JPG_Dateien=os.listdir(Ordner)
    #JPG_Dateien.sort()
    #print("Alle Dateien im Ordner:",JPG_Dateien)
    #auf Macs gibst eine ds_store Datei die vorher gelöscht werden muss
    #JPG_Dateien=[i for i in JPG_Dateien if i[-2:]!="re"] #ds_store Datei rauslöschen

JPG_Dateien=[i for i in JPG_Dateien if i.find("DS_Store")<0] #ds_store Datei rauslöschen
JPG_Dateien=[i for i in JPG_Dateien if i.find("Thumbs.db")<0] #Thumbs.db Datei rauslöschen

# Konvertiere alle Dateien in JPEG, da es manche Dateien als Format MPO haben, was zu Problemen führt

if neu_konvertieren:
    for image in JPG_Dateien:
        #nur umwandeln wenn Datei noch nicht exisitert
        if (Ordner+"/"+image[:]).find("Convert")>-1:
            print(Ordner+"/"+image[:],"Convert_JPEG Dateien existiern schon")
        else:
            im = Image.open(Ordner+"/"+image)
            im.save(Ordner+"/"+image[:-4] + "_Convert.JPEG", "JPEG")



# nur die Dateien Listen die kein Convert enthalten
JPG_Dateien=[i for i in JPG_Dateien if i.find("Convert")<0] 
JPG_Dateien=sorted(JPG_Dateien, key=lambda item: (int(item.partition('.')[0]) if item[0].isdigit() else float('inf'), item))
print("Dateien zu Bearbeiten:",JPG_Dateien)

Bild_Nummer=1

#for Datei in JPG_Dateien[:]:
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)
slide_size = (25.4, 14.29) # in cm. Größe Folien
prs.slide_width, prs.slide_height = Cm(slide_size[0]), Cm(slide_size[1])

for Datei in JPG_Dateien[:]:

    blank_slide_layout = prs.slide_layouts[6]
    print("Datei",Datei)
    slide = prs.slides.add_slide(blank_slide_layout)
    left = top = right = bottom = Cm(0.5) # Platzierung der Bilder
    pfad_convert=Ordner+"/"+Datei[:-4]+"_Convert.JPEG"
    print("JPEG in Powerpoint",pfad_convert)

    img = px_to_inches(pfad_convert)
    Bildhoehe=slide_size[1]-2
    Bild_ratio=img[0]/img[1]*Bildhoehe
    Folienmitte=slide_size[0]/2.  
    left=Folienmitte-Bild_ratio/2.  
    print("left",left,"Folienmitte",Folienmitte,"Bild_Ratio",Bild_ratio)
    pic = slide.shapes.add_picture(pfad_convert, Cm(left), top, height= Cm(Bildhoehe)) #Maße der Bilder
 
    

    
    #Platzierung der EXIF-Daten
    left = top = width = height = Inches(1)
    left = Cm(1.2)
    top  = Cm(slide_size[1]-1.5) #0.1 oben
    width= Cm(1)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    exif_text = ""
    p = tf.paragraphs[0]
    run_kamera = p.add_run()

    top  = Cm(slide_size[1]-1.0) #0.1 oben
    locationBox = slide.shapes.add_textbox(left, top, width, height)
    tf_location = locationBox.text_frame
    q = tf_location.paragraphs[0]
    run_location = q.add_run()

    if Datei[-3:]!="png":
        #if True:
        try:
            ausgabe=get_exif(Ordner+"/"+Datei)
            belichtungszeit_wert=belichtungszeit( ausgabe["ExposureTime"]  )

            print("Ausgabe",ausgabe)
            print("Brennweite",ausgabe["FocalLength"])
            print("ISO",ausgabe["ISOSpeedRatings"])
            #print("FNumber/Blende",ausgabe["FNumber"],round( ausgabe["FNumber"][0]/ausgabe["FNumber"][1],1))
            print("Exposure Time",ausgabe["ExposureTime"],str(round(ausgabe["ExposureTime"],6)))
            print("Model",ausgabe["Model"])
            #print("Lens Model",ausgabe["LensModel"])
            exif_text=exif_text+str(Bild_Nummer)+")   " 
            exif_text=exif_text+"Brennweite: "+str(ausgabe["FocalLength"]) + "  ISO: "+str(ausgabe["ISOSpeedRatings"]) 
            try:
                # muss hier noch verbessert werden
                exif_text=exif_text+"  Blende: "+str(round( ausgabe["ApertureValue"]))
                exif_text=exif_text+"  Blende: "+str(round( ausgabe["MaxApertureValue"]))
            except:
                print("Problem mit Blende")
            exif_text=exif_text+"  Belichtungszeit: "+str(belichtungszeit_wert)+" s"
            
            if Datum_Statt_Kamera_Daten:
                run_kamera.text=change_date_format(ausgabe["DateTime"])
            else:
                run_kamera.text=exif_text

            # Geo Informationen einbauen    
            try: 
                North_Dezi,East_Dezi=gps_converter(ausgabe["GPSInfo"][2],ausgabe["GPSInfo"][4])
                
                # passing the coordinates
                locname = geoLoc.reverse(f"{North_Dezi}, {East_Dezi}")

                # printing the address/location name
                print(locname.address)
                run_location.text=locname.address
            except:
                print ("Keine GeoDaten")
            Bild_Nummer=Bild_Nummer+1
            print("-----------------------------------")
        except:
            print("Fehler bei Exifdaten")
            #exif_text=Datei
            print("-----------------------------------")


    font = run_kamera.font
    font.name = 'Calibri'
    font.size = Pt(fontsize)

    font = run_location.font
    font.name = 'Calibri'
    font.size = Pt(fontsize)




# Overview
if Overview_Slide:
    # Empty Page for Overview
    blank_slide_layout = prs.slide_layouts[6]
    print("Datei",Datei)
    slide = prs.slides.add_slide(blank_slide_layout)

    index=0
    for hoehe in range(4):
        for laenge in range(5):
            try:
                left = top = right = bottom = Inches(0.5) # Platzierung der Bilder
                left = Inches(0.5)+Inches(laenge*2)
                top =  Inches(0.5)+Inches(hoehe*2)
        
                pfad_convert=Ordner+"/"+JPG_Dateien[index][:-4]+"_Convert.JPEG"
                print("JPEG in Powerpoint Mehrfachplot",pfad_convert)
                pic = slide.shapes.add_picture(pfad_convert, left, top, height= Inches(1.5)) #Maße der Bilder
        
                index=index+1
            except:
                print("Keine Bilder mehr für übersicht")

prs.save('Bilder_Fototron_'+Ordner+'.pptx')
