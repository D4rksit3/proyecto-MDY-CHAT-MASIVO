import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import time, pprint, string, random
import pymongo
from pymongo import MongoClient

try: 
#GENERAR ID ALEATORIO   
    characters = list(string.ascii_letters + string.digits)
    
    length = int(17)
    random.shuffle(characters)
    password = []
    for i in range(length):
        password.append(random.choice(characters))
    random.shuffle(password)
        
    valor = "".join(password)

    """valor = generate_random_password()"""
    print ("ID generado:",valor)
    
    
#Ruta de excel
    directorio = '/home/rj/Descargas'
    archivo = directorio+'/prueba.xlsx'
#Lectura de Excel
    readUser = pd.read_excel(archivo, sheet_name='Hoja1')
    lista = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60]
    for tabla in lista:
        
        user = readUser.loc[tabla,['USUARIO']]
        camp = readUser.loc[tabla,['CAMPAÑA']]
        date = readUser.loc[tabla,['NOMBRES COMPLETOS']]
        puesto = readUser.loc[tabla,['PUESTO']]
        try:
            user+"@mdy"
            
        except: 
            pass

#Conexion a la DB de Mongo en prueba
    ip = "10.200.52.227"
    uri = 'mongodb://10.200.52.227:27017/?readPreference=primary&directConnection=true&ssl=false'
    client = MongoClient(uri)
    db = client.rocketchat
    coleccion = db.users
#Json de ejemplo

    usuario = {
    '_id': 'sqkwlk1klk213l324',
     'active': True,
 
     'emails': [{'address': 'rcarlos2@mdy', 'verified': False}],
     'name': 'Roberto Carlos2',
     'requirePasswordChange': True,
     'roles': ['bot'],
     'services': {'password': {'bcrypt': '$2b$10$yKB./.bj6fIWJd7yOIU44e/cuxszCRczWUr2gHp5xXry6MUKsU9BG'}},
     'settings': {},
     'status': 'offline',
     'type': 'user',
     'username': 'rcarlos2'
    }   

#Actualizar usuario JSON

    for asesor in user,camp,date,puesto:
        ids = usuario['_id']= f'{str(valor)}'
        emails =  usuario['emails'] = f"{'address': 'rcarlos2@mdy', 'verified': False}"


    try:
        resultado = coleccion.insert_one(usuario)
        print ('objeto creado o actualizado: \n'+ str(resultado.inserted_id))
    
    except pymongo.errors.DuplicateKeyError as err:
        print("Usuario duplicado: \n",usuario['username'] )
        pass
    
    except pymongo.errors.ServerSelectionTimeoutError as err:
        print("Servidor no conectado: \n",err)
        pass
                
except KeyboardInterrupt:
    print("[!]Detenido por el usuario")
    pass
