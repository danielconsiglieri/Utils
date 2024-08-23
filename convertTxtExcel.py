#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#Objetivo: Transforma txt em arquivo Excel
#Autor: Daniel Consiglieri
#Data: set-2023
#Revisão: 27-out-2023

import argparse
import csv
import io
import os
import pandas as pd
import sys
import xlsxwriter

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Le arquivos csv e converte para xlsx')
    parser.add_argument('-e', '--entrada', type=str, required=True, help='Arquivo de entrada com txt')
    
    args = parser.parse_args()
    #abre a codificação ANSI
    arquivo_leitura= open(args.entrada,"r", encoding = "ISO-8859-1")
    #Lê cada linha montando montado uma lista
    verificado = arquivo_leitura.readlines()
    arquivo_leitura.close()   
    #Variavel para guardar a leitura
    trabalhado = ""
    #flag de discrepancia
    discrFlag = False
    colVaziaFlag = False    
    
    #Verifica se há discrepancia no número de linhas
    if (len(verificado) > 2):
        #Corrige o caso de coluna final vazia
        if ";\n" in verificado[0]:
            print("Arquivo de origem " + args.entrada + " tem coluna vazia")
            verificado[0] = verificado[0].replace(";\n",";fakecolunaV\n")
            colVaziaFlag = True
        if( verificado[0].count(";") != verificado[1].count(";") ):
            print("Arquivo de origem " + args.entrada + " tem discrepancia de colunas")
            #Seta a flag de discrepancia
            discrFlag = True
            #corrige a discrepancia acrescentando a fakecoluna\n
            verificado[0] = verificado[0].replace('\n','') + ";fakecoluna\n"        
    
    #monta o arquivo de trabalho
    for v in range(len(verificado)):
        trabalhado += verificado[v]
    
    #retira caracter de finalizacao COBOL
    trabalhado.replace('\x1a','')   
    
    read_file = pd.read_csv(io.StringIO(trabalhado), sep=';', dtype=str)
    
    #deleta a coluna da discrepancia
    if discrFlag:
        read_file.drop(['fakecoluna'], axis=1, inplace=True)
    #deleta coluna vazia
    if colVaziaFlag:    
        read_file.drop(['fakecolunaV'], axis=1, inplace=True)
    
    #engine='xlsxwriter' ajuda remover caracteres inválidos, header = True faz com que seja impresso unamed column       
    read_file.to_excel((args.entrada).replace("TXT","xlsx").replace("txt","xlsx"), index=None, header=True, engine='xlsxwriter')
