#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Author : Rico Fauchard
# Year: 2017


import xlrd as xl
import xlsxwriter

# Specifier ici le nom du fichier à decouper
FILENAME = "../DATA_SOURCE/big_file_excel/articles/Articles DEV9.xlsx"
# Nom de fichier de sortie
OUTPUT_NAME ="../DATA_SOURCE/splited_excel/articles/Articles DEV9"
#Nombre de ligne à mettre dans un fichier
NB_LIGNE = 500

def create_portion(donnee,OUTPUT_NAME,num_fichier,entete):
    nom_de_sortie = OUTPUT_NAME+ "_" + str(num_fichier) + ".xlsx"
    classeur = xlsxwriter.Workbook(nom_de_sortie)
    feuille = classeur.add_worksheet()

    format_entete = classeur.add_format({'bold': True, 'align': 'left'})

    for i,val in enumerate(entete):
        feuille.write(0, i, entete[i] or "",format_entete)

    for i,vals in enumerate(donnee):
        print vals
        for j,valeur in enumerate(vals):
            feuille.write(i+1 ,j,valeur) # i + 1 pour ne pas ecraser l'entete



if __name__ == "__main__":
    classeur = xl.open_workbook(FILENAME)
    classeur.encoding
    feuille = classeur.sheet_by_index(0)

    entete = []
    donnee = []
    num_fichier = 0

    for num_ligne in range(feuille.nrows):
        ligne = feuille.row_values(num_ligne,0)
        if num_ligne == 0:
            entete = ligne
            print entete
        else:
            donnee.append(ligne)
            if (num_ligne % NB_LIGNE)==0:
                num_fichier += 1
                create_portion(donnee,OUTPUT_NAME,num_fichier,entete)
                donnee = []

    if donnee:
        num_fichier += 1
        create_portion(donnee,OUTPUT_NAME,num_fichier,entete)
        donnee = []



