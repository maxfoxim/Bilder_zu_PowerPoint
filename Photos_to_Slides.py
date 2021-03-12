# -*- coding: utf-8 -*-
"""
Created on Tue Jul  9 15:19:09 2019

Letzte Änderungen
27.11.19: Belichtungszeit richtig angeben
28.11.19: Convert Problem, Index von 5 auf 4 gesetzt
3.8.20: Mit Unterordner und PP-Name nach Ordner
10.11.20 Anscheinend neue EXIF Formatierung.
24.1.21 kleiner Bugfix
"""

# fuer pptx und exifread: pip install python-pptx und pip install exifread

from pptx import Presentation
from pptx.util import Inches
import os
from PIL import Image
from PIL.ExifTags import TAGS



Ordner="HighLow"
Mit_UnterOrdner=True


def belichtungszeit(zeit):
    if zeit<1:
        zeit=int(1/zeit)
        zeit="1/"+str(zeit)
    return zeit


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
    print("Alle Dateien im Ordner:",JPG_Dateien)
    #auf Macs gibst eine ds_store Datei die vorher gelöscht werden muss
    #JPG_Dateien=[i for i in JPG_Dateien if i[-2:]!="re"] #ds_store Datei rauslöschen

JPG_Dateien=[i for i in JPG_Dateien if i.find("DS_Store")<0] #ds_store Datei rauslöschen
JPG_Dateien=[i for i in JPG_Dateien if i.find("Thumbs.db")<0] #Thumbs.db Datei rauslöschen

# Konvertiere alle Dateien in JPEG, da es manche Dateien als Format MPO haben, was zu Problemen führt

for image in JPG_Dateien:
    #nur umwandeln wenn Datei noch nicht exisitert
    if (Ordner+"/"+image[:]).find("Convert")>-1:
        print(Ordner+"/"+image[:],"Convert_JPEG Dateien existiern schon")
    else:
        im = Image.open(Ordner+"/"+image)
        im.save(Ordner+"/"+image[:-4] + "_Convert.JPEG", "JPEG")



# nur die Dateien Listen die kein Convert enthalten
JPG_Dateien=[i for i in JPG_Dateien if i.find("Convert")<0] 


print("Dateien zu Bearbeiten:",JPG_Dateien)

Bild_Nummer=1

#for Datei in JPG_Dateien[:]:
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)



for Datei in JPG_Dateien[:]:
    blank_slide_layout = prs.slide_layouts[6]
    print("Datei",Datei)
    slide = prs.slides.add_slide(blank_slide_layout)
    left = top = right = bottom = Inches(0.5) # Platzierung der Bilder
    pfad_convert=Ordner+"/"+Datei[:-4]+"_Convert.JPEG"
    print("JPEG in Powerpoint",pfad_convert)
    pic = slide.shapes.add_picture(pfad_convert, left, top, height= Inches(6.0)) #Maße der Bilder
    
    
    #Platzierung der EXIF-Daten
    left = top = width = height = Inches(1)
    left = Inches(1.2)
    top  = Inches(7.1)
    width= Inches(2)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.text = ""
    
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
            tf.text=tf.text+str(Bild_Nummer)+")   " 
            tf.text=tf.text+"Brennweite: "+str(ausgabe["FocalLength"]) + "  ISO: "+str(ausgabe["ISOSpeedRatings"]) 
            try:
                tf.text=tf.text+"  Blende: "+str(round( ausgabe["ApertureValue"]))
            except:
                print("Problem mit Blende")
                tf.text=tf.text+"  Blende: "+str(round( ausgabe["MaxApertureValue"]))
            tf.text=tf.text+"  Belichtungszeit: "+str(belichtungszeit_wert)+" s"

            Bild_Nummer=Bild_Nummer+1
            print("-----------------------------------")
        except:
            print("Fehler bei Exifdaten")
            tf.text=Datei
            print("-----------------------------------")



# Overview
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
