#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Author : Rico Fauchard
# Year: 2017


import xlrd as xl
import xlsxwriter
import os

# Specifier ici le nom du fichier à decouper
NOM_FICHIER = "PATH_TO_FILE/FILE.xlsx"
# Nom de fichier de sortie
NOM_SORTIE =".PATH_TO_FILE/NOM_FICHIER_SORTIE"   # nom sans l'extension
#Nombre de ligne à mettre dans un fichier
NB_LIGNE = 750

def creer_portion(donnee,num_fichier,entete):
    nom_de_sortie = NOM_SORTIE+ "_" + str(num_fichier) + ".xlsx"
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
    classeur = xl.open_workbook(NOM_FICHIER)
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
                creer_portion(donnee,num_fichier,entete)
                donnee = []

    if donnee:
        num_fichier += 1
        creer_portion(donnee,num_fichier,entete)
        donnee = []



