# Las dependencias:
from logging import NullHandler
import pandas as pd
import spacy as sp
import os
import PySimpleGUI as sg
import webview as wv
from spacy.matcher import DependencyMatcher as match
from spacy.lang.es.stop_words import STOP_WORDS
from spacy import displacy
import stanza
import spacy_stanza
from string import punctuation
#los siguientes paquetes son para cuando se implemente polyglot:
#import polyglot 
#from polyglot.text import Text, Word





sg.theme('Reddit')  
# El diseño de la ventana.
#primero definimos el menu superior:
menu_def = ['Opciones', ['Descargar Modelo Spacy para Entidades','Descargar Modelo Spacy para sustantativos, verbos y lemas', 'Descargar Modelo Stanza','Cerrar']],['&Ayuda',['&Ayuda', 'About...']],

#luego creamos el layout
layout = [  [sg.Menu(menu_def)],
            [sg.Text('Bienvenido al GN Analizer')],
            [sg.Text('Nombre del archivo de salida'), sg.InputText(key="-OUTFILE-"),sg.Radio('.xlsx',"FILETYPE", key="-XLSX-", default=True), sg.Radio('.csv',"FILETYPE", key="-CSV-", default=True), sg.Radio('.json',"FILETYPE", key="-JSON-", default=True)],
            [sg.Text('Orden de palabras en el archivo generado'),sg.Radio('Mantener palabras en filas originales, una columna por opción',"WORDORDER", key="-KEEPROW-", default=True), sg.Radio('Una palabra por fila y una columna para opción',"WORDORDER", key="-ONEWORDROW-", default=False), sg.Radio('Todas las palabras en la misma columna, una fila por palabra',"WORDORDER", key="-KEEPROWONECOL-", default=False), sg.Radio('Mantener filas, todo en una columna',"WORDORDER", key="-ONECOL-", default=False)],
            [sg.Text('¿Dónde querés guardar el archivo?'), sg.InputText(key="-OUTFOLDER-"), sg.FolderBrowse(target="-OUTFOLDER-")],
            [sg.Text('Elegí archivo a analizar'),sg.Input(), sg.FileBrowse(key="-IN-")],
            [sg.Text('¿Cuál es el nombre de la pestaña?'), sg.InputText(key="-SHEET-")],
            [sg.Text('¿Cuál es el nombre de la columna?'), sg.InputText(key="-COL-")],
            [sg.Text('Elegir un motor de NLP'),sg.Radio('Spacy',"ENGINE", key="-SPAC-", default=True), sg.Radio('Stanford Stanza',"ENGINE", key="-STANZ-", default=False), sg.Radio('Freeling',"ENGINE", key="-FREEL-", default=False), sg.Radio('Gensim',"ENGINE", key="-GENSIM-", default=False), sg.Radio('Polyglot',"ENGINE", key="-POLY-", default=False),sg.Radio('TextBLOB',"ENGINE", key="-BLOB-", default=False)],
            [sg.Text('PoS Tags'), sg.Checkbox('Nombres Propios', key="-PROPN-", default=True), sg.Checkbox('Sustantivos y Verbos', key="-NOUNV-", default=True), sg.Checkbox('Entidades Naturales', key="-NER-", default=False), sg.Checkbox('Separar el label en las entidades naturales', key="-NERL-", default=False), sg.Checkbox('Visualizar entidades de cada frase.', key="-DISPLAY-", default=False), sg.Checkbox('Exportar el sentiment.', key="-SENTIMENT-", default=False)],
            
            [sg.Button('OK'), sg.Button('Cerrar')]]

#Y la ventana:
window = sg.Window('GN Analizer', layout)

# Hacemos un loop para procesar eventos y tomar los inputs de la ventana como values
while True:
    event, values = window.read()
    #primero esperamos eventos del dropdown menu:
    if event == 'Descargar Modelo Spacy para Entidades': 
        os.system('python -m spacy download es_core_news_lg')
    if event == 'Descargar Modelo Spacy para sustantativos, verbos y lemas':
        os.system('python -m spacy download es_dep_news_trf')
    if event == 'Descargar Modelo Stanza':
        stanza.download('es', processors={'ner': 'conll02'})       
    if event == 'About...':     
        sg.popup('Este programa fue creado para probar spacy y otros motores de NLP.', 'Version 0.7', 'PySimpleGUI rocks...')   
    #despues, si el usuario clickea en Ok procesamos todo lo que sigue
    if event == 'OK':
        #Lo primero que hacemos despues del OK es chequear que estén todas las opciones completadas y sino prompteo un error
        if values["-OUTFILE-"] == "":
            sg.popup(f"No determinaste el nombre del archivo que vamos a generar.")
        elif values["-OUTFOLDER-"] == "":
            sg.popup(f"No determinaste la carpeta en la que se va a guardar el archivo generado.")    
        elif values["-IN-"] == "" :
            sg.popup(f"No seleccionaste un archivo para analizar.")
        elif values["-SHEET-"] == "" :
            sg.popup(f"No determinaste en qué hoja del archivo está la columna de texto.")
        elif values["-COL-"] == "":
            sg.popup(f"No determinaste cuál es la columna de texto.")
        elif values["-NOUNV-"] == False and values["-PROPN-"] == False and values["-NER-"] == False and values["-DISPLAY-"] == False and values["-SENTIMENT-"] == False: 
            sg.popup(f"No clickeaste ninguna opción para analizar las Parts of Speech")
        #Si se llenaron todas las opciones (ninguna esta vacia) procedemos a ejecutar el programa
        else:
            original_file = values["-IN-"]
            sheet_name = values["-SHEET-"]
            column = values["-COL-"]
            #A parte de tomar los valores inputeados, usamos os para normalizar la ruta de la carpeta (porque va a cambiar segun el OS) y para joinearla con el nombre del archivo
            new_file = os.path.join(os.path.normcase(values["-OUTFOLDER-"]), values["-OUTFILE-"])
            #Ahora, segun que opcion clickeo el usuario, vamos a definir el motor y modelo con las variables nlp y nlpner
            # Defino las variables:
            if  values["-SPAC-"] == True:
                try:
                    nlp = sp.load('es_dep_news_trf')
                    nlpner = sp.load('es_core_news_lg')
                #Con spacy usamos dos modelos diferentes, uno sirve para NER y el otro deptrees y lematizacion
                except Exception as s:
                    sg.popup(f"Ups! Ocurrió el siguiente error: {s}, seguramente tengas que descargar el motor de idioma de Spacy en las opciones.")
                pos_tag=['NOUN', 'VERB']
                pos_propn=['PROPN']
            elif values["-STANZ-"] == True:
                try:
                   nlp = spacy_stanza.load_pipeline("es", processors={'ner': 'conll02'})
                   nlpner = nlp
                #con stanza usamos un solo modelo para las dos variables
                except Exception as n:
                    sg.popup(f"Ups! Ocurrió el siguiente error: {n}, seguramente tengas que descargar la el motor de Stanza en las opciones.")
                pos_tag=['NOUN', 'VERB']
                pos_propn=['PROPN']
            elif values["-FREEL-"] == True:
                nlp = sp.load('es_dep_news_trf')
                nlpner = sp.load('es_core_news_lg')
                pos_tag=['NOUN', 'VERB']
                pos_propn=['PROPN']
                sg.popup(f"Freeling aún no está disponible. Vamos a usar Spacy.")
            elif values["-GENSIM-"] == True:
                nlp = sp.load('es_dep_news_trf')
                nlpner = sp.load('es_core_news_lg')
                pos_tag=['NOUN', 'VERB']
                pos_propn=['PROPN']
                sg.popup(f"GENSIM aún no está disponible. Vamos a usar Spacy.")
            elif values["-POLY-"] == True:
                nlp = sp.load('es_dep_news_trf')
                nlpner = sp.load('es_core_news_lg')
                pos_tag=['NOUN', 'VERB']
                pos_propn=['PROPN']
                sg.popup(f"Polyglot aún no está disponible. Vamos a usar Spacy.")
            elif values["-BLOB-"] == True:
                nlp = sp.load('es_dep_news_trf')
                nlpner = sp.load('es_core_news_lg')
                pos_tag=['NOUN', 'VERB']
                pos_propn=['PROPN']
                sg.popup(f"TextBlob aún no está disponible. Vamos a usar Spacy.")
            try:
                my_file = pd.read_excel(original_file, sheet_name= sheet_name)
                my_file_index = my_file.index
                number_of_rows = len(my_file_index)
                sg.popup(f"La columna que vamos a analizar tiene {number_of_rows} filas.")
# Ejecuto acá los fors:
                i = 1
                stopwords= list(STOP_WORDS)
                palabras = []
                nombres = []
                entidades = []
                labels = []
                columna = []
#tengo tres rutinas, voy a ejecutar la primera si el usuario pidió una palabra por fila y la segunda si quiere todas las palabras de una fila del archivo original por fila
                if values["-ONEWORDROW-"] == True: 
                    for row_info in my_file[column]:
                        sentence = nlp(str(row_info))
                        sg.OneLineProgressMeter('Avance de la operación.',i, number_of_rows, 'OK')
                        i=i+1
                        for token in sentence:
                            if(token.text in stopwords or token.text in punctuation):
                                continue
                            if values["-PROPN-"] == True and (token.pos_ in pos_propn):
                                nombres.append(str(token.text))
                            if values["-NOUNV-"] == True and token.pos_ in pos_tag:
                                palabras.append(token.lemma_.lower())
                        if values["-NER-"] == True:
                            sentencener = nlpner(str(row_info))
                            if values["-NERL-"] == True:
                                for entidad in sentencener.ents:
                                    entidades.append(str(entidad.text))
                                    labels.append(str(entidad.label_))
                            else:
                                for entidad in sentencener.ents:
                                    entidades.append(str(entidad.text +":"+ entidad.label_))
                        if values["-DISPLAY-"] == True :
                            options = {"fine_grained": True, "compact": False, "add_lema": True, "color": "blue"}                        
                            wv.create_window('Current Spacy Row', html= displacy.render(sentencener, style="ent", options=options))
                            wv.start()
#Si pidió mantener las mismas filas del archivo original, ejecuto la segunda rutina                               
                elif values["-KEEPROW-"] == True:
                    for row_info in my_file[column]:
                        sentence = nlp(str(row_info))
                        sg.OneLineProgressMeter('Avance de la operación.',i, number_of_rows, 'OK')
                        i=i+1
                        nombres_temp = ""
                        palabras_temp = ""
                        entidades_temp = ""
                        labels_temp = ""
                        for token in sentence:
                            if(token.text in stopwords or token.text in punctuation):
                                continue
                            if values["-PROPN-"] == True and (token.pos_ in pos_propn):
                                nombres_temp=str(nombres_temp)+","+str(token.text)
                            if values["-NOUNV-"] == True and token.pos_ in pos_tag:
                                palabras_temp=str(palabras_temp)+","+str(token.lemma_.lower())
                        nombres.append(nombres_temp)
                        palabras.append(palabras_temp)
                        if values["-NER-"] == True:
                            sentencener = nlpner(str(row_info))
                            if values["-NERL-"] == True:
                                for entidad in sentencener.ents:
                                    entidades_temp = entidades_temp + "," +str(entidad.text)
                                    labels_temp = labels_temp + "," + str(entidad.label_)
                            else:
                                for entidad in sentencener.ents:
                                    entidades_temp = entidades_temp + "," +str(entidad.text +":"+ entidad.label_)
                        if values["-NERL-"] == True:
                            entidades.append(entidades_temp)
                            labels.append(labels_temp)
                        else:
                            entidades.append(entidades_temp)
                        if values["-DISPLAY-"] == True :
                            options = {"fine_grained": True, "compact": False, "add_lema": True, "color": "blue"}                        
                            wv.create_window('Current Spacy Row', html= displacy.render(sentencener, style="ent", options=options))
                            wv.start()
#Si pidió mantener las mismas filas del archivo original pero todo en una sóla columna, ejecuto esta
                elif values["-ONECOL-"] == True:
                    for row_info in my_file[column]:
                        sentence = nlp(str(row_info))
                        sg.OneLineProgressMeter('Avance de la operación.',i, number_of_rows, 'OK')
                        i=i+1                
                        for token in sentence:
                            if(token.text in stopwords or token.text in punctuation):
                                continue
                            if values["-PROPN-"] == True and (token.pos_ in pos_propn):
                                columna.append(str(token.text))
                            if values["-NOUNV-"] == True and token.pos_ in pos_tag:
                                columna.append(token.lemma_.lower())
                        if values["-NER-"] == True:
                            sentencener = nlpner(str(row_info))
                            if values["-NERL-"] == True:
                                for entidad in sentencener.ents:
                                    columna.append(str(entidad.text +":"+ entidad.label_))
                            else:
                                for entidad in sentencener.ents:
                                    columna.append(str(entidad.text))
                        if values["-DISPLAY-"] == True :
                            options = {"fine_grained": True, "compact": False, "add_lema": True, "color": "blue"}                        
                            wv.create_window('Current Spacy Row', html= displacy.render(sentencener, style="ent", options=options))
                            wv.start()               
#Si pidió tener una fila por palabra y una sóla columna
                elif values["-KEEPROWONECOL-"] == True:
                    for row_info in my_file[column]:
                        sentence = nlp(str(row_info))
                        sg.OneLineProgressMeter('Avance de la operación.',i, number_of_rows, 'OK')
                        i=i+1
                        columna_temp = ""
                        for token in sentence:
                            if(token.text in stopwords or token.text in punctuation):
                                continue
                            if values["-PROPN-"] == True and (token.pos_ in pos_propn):
                                columna_temp=str(columna_temp+","+str(token.text))
                            if values["-NOUNV-"] == True and token.pos_ in pos_tag:
                                columna_temp=str(columna_temp+","+str(token.lemma_.lower()))
                        if values["-NER-"] == True:
                            sentencener = nlpner(str(row_info))
                            if values["-NERL-"] == True:
                                for entidad in sentencener.ents:
                                    columna_temp = str(columna_temp + "," +str(entidad.text))
                            else:
                                for entidad in sentencener.ents:
                                    columna_temp = str(columna_temp + "," +str(entidad.text +":"+ entidad.label_))
                        columna.append(columna_temp)
                        if values["-DISPLAY-"] == True :
                            options = {"fine_grained": True, "compact": False, "add_lema": True, "color": "blue"}                        
                            wv.create_window('Current Spacy Row', html= displacy.render(sentencener, style="ent", options=options))
                            wv.start()
#Ahora que tenemos las listas armadas (o columna, o palabras o nombres o entidades), vamos a generar el dataframe final, según las opciones que pidió el usuario

                if values["-KEEPROWONECOL-"] == True or values["-ONECOL-"] == True:
                    new_df = pd.DataFrame({"Todas las palabras" : columna})
                elif values["-NOUNV-"] == True and values["-PROPN-"] == True and values["-NER-"] == True:
                    if values["-NERL-"] == True:
                        new_df1 = pd.DataFrame({"Lemas" : palabras})
                        new_df2 = pd.DataFrame({"Nombres propios" : nombres})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df4 = pd.DataFrame({"Labels" : labels})
                        new_df = new_df1.join(new_df2.join(new_df3.join(new_df4)))
                    else:
                        new_df1 = pd.DataFrame({"Lemas" : palabras})
                        new_df2 = pd.DataFrame({"Nombres propios" : nombres})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df = new_df1.join(new_df2.join(new_df3))
                elif values["-NOUNV-"] == True and values["-PROPN-"] == True and values["-NER-"] == False:
                    new_df1 = pd.DataFrame({"Lemas" : palabras})
                    new_df2 = pd.DataFrame({"Nombres propios" : nombres})
                    new_df = new_df1.join(new_df2)
                elif values["-NOUNV-"] == True and values["-PROPN-"] == False and values["-NER-"] == True : 
                    if values["-NERL-"] == True:
                        new_df1 = pd.DataFrame({"Lemas" : palabras})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df4 = pd.DataFrame({"Labels" : labels})                   
                        new_df = new_df1.join(new_df3.join(new_df4))
                    else:
                        new_df1 = pd.DataFrame({"Lemas" : palabras})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df = new_df1.join(new_df3)
                elif values["-NOUNV-"] == False and values["-PROPN-"] == True and values["-NER-"] == True :
                    if values["-NERL-"] == True:
                        new_df2 = pd.DataFrame({"Nombres propios" : nombres})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df4 = pd.DataFrame({"Labels" : labels}) 
                        new_df = new_df2.join(new_df3.join(new_df4))                      
                    else:
                        new_df2 = pd.DataFrame({"Nombres propios" : nombres})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df = new_df2.join(new_df3)
                elif values["-NOUNV-"] == False and values["-PROPN-"] == True and values["-NER-"] == False : 
                    new_df = pd.DataFrame({"Nombres propios" : nombres})
                elif values["-NOUNV-"] == True and values["-PROPN-"] == False and values["-NER-"] == False:
                    new_df = pd.DataFrame({"Lemas" : palabras})
                elif values["-NOUNV-"] == False and values["-PROPN-"] == False and values["-NER-"] == True:
                    if values["-NERL-"] == True:
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df4 = pd.DataFrame({"Labels" : labels}) 
                        new_df = new_df3.join(new_df4)                     
                    else:
                        new_df = pd.DataFrame({"Entidades Naturales" : entidades})
#Ahora según la opción de archivo a generar, vamos a generar el archivo final                
                if values["-XLSX-"] == True:
                    new_file = new_file+".xlsx"
                    new_df.to_excel(new_file,sheet_name='NLP', index = False)
                elif values["-CSV-"] == True:
                    new_file = new_file+".csv"
                    new_df.to_csv(new_file, index = False)
                elif values["-JSON-"] == True:
                    new_file = new_file+".json"
                    new_df.to_json(path_or_buf=new_file)
                sg.popup(f'Archivo {new_file} creado correctamente.')
            except Exception as e:
                sg.popup(f"Ups! Ocurrió el siguiente error: {e}")
    if event == "Cerrar" or event == sg.WIN_CLOSED:
        force_exit = True
        break
window.close()



