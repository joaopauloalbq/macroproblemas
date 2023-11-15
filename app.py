import pandas as pd
import streamlit as st
from openpyxl import load_workbook

def getKey(dic, valor):
    for c, v in dic.items():
        if v == valor:
            return c

def getCities(sheet, linha, coluna):
    cidades = {}

    cell = sheet.cell(row=linha, column=coluna).value
    while cell is not None:
        cidades[coluna] = cell

        coluna += 1
        cell = sheet.cell(row=linha, column=coluna).value

    return cidades

def main():        
    planilha = st.file_uploader('Selecione a planilha', type='xlsx')

    if planilha is not None:
        workbook = load_workbook(planilha)

        dataDictionary = {"MACROPROBLEMA":[], "CIDADE":[], "RESULTADO":[], "TOTAL":[], "PRIORIDADE":[], "MACROREGIAO":[]}
        for sheetname in workbook.sheetnames:
            sheet = workbook[sheetname]

            cidades = getCities(sheet, linha = 2, coluna = 2)

            # Constrói data dictionary
            linha = 3
            macroProblema = sheet.cell(row=linha, column=1).value
            while macroProblema is not None:
                coluna = 2
                total = sheet.cell(row=linha, column=getKey(cidades,"TOTAL")).value
                priodade = sheet.cell(row=linha, column=getKey(cidades,"PRIORIDADES SANITÁRIAS")).value
                resultado = sheet.cell(row=linha, column=coluna).value
                while resultado is not None and cidades[coluna] != "TOTAL":
                    
                    dataDictionary["MACROPROBLEMA"].append(macroProblema)
                    dataDictionary["CIDADE"].append(cidades[coluna])
                    dataDictionary["RESULTADO"].append(resultado)
                    dataDictionary["TOTAL"].append(total)
                    dataDictionary["PRIORIDADE"].append(priodade)
                    dataDictionary["MACROREGIAO"].append(sheet.title)

                    coluna += 1
                    resultado = sheet.cell(row=linha, column=coluna).value

                linha += 1
                macroProblema = sheet.cell(row=linha, column=1).value

        df = pd.DataFrame(dataDictionary)
        st.write(df)

main()