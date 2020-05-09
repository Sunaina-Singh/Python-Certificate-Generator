#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri May  8 00:42:52 2020

@author: sunainasingh
"""

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

#importing two files, where one is the current set of data, and the second is the master list
#master list contains details of all the certificates that have been issued till date

file = 'Python Test 1.xlsx'
masterfile = 'Python Master.xlsx'
data = pd.read_excel(file)
masterdata = pd.read_excel(masterfile)

#filtering out IDs that are already in the master list
uniquedata = data.loc[data.ID.isin(masterdata.ID)==False]

#filtering out rewards that have not been approved
finaldata = uniquedata.loc[uniquedata['Reward']=='Yes']

#taking the refined data to a new excel file
finaldata.to_excel('Certificate Generator.xlsx')

finalfile = 'Certificate Generator.xlsx'
filetoextract = pd.ExcelFile(finalfile).parse('Sheet1', index_col = 0)

#creating lists of different data sets required to generate certificate. Simple iteration has been used.
names = []
for i in filetoextract['Name']:
    names.append(i)
signatory = []
for j in filetoextract['Leader']:
    signatory.append(j)
purpose = []
for k in filetoextract['For']:
    purpose.append(k)
ids = []
for l in filetoextract['ID']:
    ids.append(l)

#generating certificates for all the rewards to be issued for the month
#using a dummy certificate created with MS word.
for a, b, c, d in zip(names,signatory,purpose, ids):
    cert = Image.open('Dummy Certificate.png')
    draw = ImageDraw.Draw(cert)
    fonttype1 = ImageFont.truetype("Caviardreams_bi.ttf", 34)
    fonttype2 = ImageFont.truetype("Caviardreams.ttf", 28)
    #you can use other fonts as desired, which can be downloaded as ttf
    #w, h = draw.textsize(a, fonttype) 
    #use the above line to center align the text based on its size for 'names'
    #draw.text(xy=((cert.width - w)/2, 235), text = a, fill = "black", font = fonttype)
    #use the above line if you are center aligning for adding names to the image
    draw.text(xy=(550,265), text = a, fill = "black", font = fonttype1)
    draw.text(xy=(698,500), text = b, fill = "black", font = fonttype2)
    draw.text(xy=(535,348), text = c, fill = "black", font = fonttype2)
    draw.text(xy=(165,496), text = "May 2020", fill = "black", font = fonttype2)
    cert.save("Cert_" + str(d) + ".png")

