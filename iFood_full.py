import requests
from bs4 import BeautifulSoup
import json
import pandas as pd

def printProgressBar (iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print('\r%s |%s| %s%% %s' % (prefix, bar, percent, suffix), end = printEnd)
    # Print New Line on Complete
    if iteration == total:
        print()

#Until now, the code was to show the progress o the program

def getNamePhone(site):

    page = requests.get(site)

    soup = BeautifulSoup(page.text, 'html.parser') #library that get the info from the website, and transform it in HTML

    script = soup.find_all('script',{"type": "application/ld+json"}) # find the "Place" that we want to dive in
    script_BM = soup.find_all('script', {"id": "__NEXT_DATA__"}) # find the "Place" that we want to dive in

    for idx in script:
        jsondata = idx.contents[0]

    newDictionary=json.loads(str(jsondata))
    try:
        if newDictionary["name"] != "iFood":
            try:
                telefone = newDictionary["telephone"] #getting the telephone and so on..
            except KeyError as error:
                telefone = "-"

            try:
                nome = newDictionary['name']
            except KeyError as error:
                nome = "-"


            try:
                tipo = newDictionary["servesCuisine"]
            except KeyError as error:
                tipo = "-"


            try:
                nomerua = newDictionary['address']['streetAddress']
            except KeyError as error:
                nomerua = "-"


            try:
                bairro = newDictionary['address']['addressLocality']
            except KeyError as error:
                bairro = "-"


            try:
                CEP = newDictionary['address']['postalCode']
            except KeyError as error:
                CEP = "-"

            try:
                Latitude = newDictionary['geo']['latitude']
            except KeyError as error:
                Latitude = "-"

            try:
                Longitude = newDictionary['geo']['longitude']
            except KeyError as error:
                Longitude = "-"

            try:
                hora = newDictionary['openingHoursSpecification']
                segunda = []
                terca = []
                quarta = []
                quinta = []
                sexta = []
                sabado = []
                domingo = []
            except KeyError as error:
                hora = "-"

            for i in hora:
                if i['dayOfWeek'] == 'http://schema.org/Monday':
                    segunda = i['opens'] +'|'+ i['closes'] #here we need to make a for function because we have a matrix in the web site
            for i in hora:
                if i['dayOfWeek'] == 'http://schema.org/Tuesday':
                    terca = i['opens'] +'|'+ i['closes']
            for i in hora:
                if i['dayOfWeek'] == 'http://schema.org/Wednesday':
                    quarta = i['opens'] +'|'+ i['closes']
            for i in hora:
                if i['dayOfWeek'] == 'http://schema.org/Thursday':
                    quinta = i['opens'] +'|'+ i['closes']
            for i in hora:
                if i['dayOfWeek'] == 'http://schema.org/Friday':
                    sexta = i['opens'] +'|'+ i['closes']
            for i in hora:
                if i['dayOfWeek'] == 'http://schema.org/Saturday':
                    sabado = i['opens'] +'|'+ i['closes']
            for i in hora:
                if i['dayOfWeek'] == 'http://schema.org/Sunday':
                    domingo = i['opens'] +'|'+ i['closes']



            for idx in script_BM:
                jsondata = idx.contents[0]


            newDictionary_BM=json.loads(str(jsondata)) #here is the same thing as before, but the info is allocated in a different part os the web site
            try:
                KA = newDictionary_BM['props']['initialState']['restaurant']['details']['tags']
            except KeyError as error:
                KA = "-"
            categoria=[]
            if "KEY_ACCOUNT" in KA:
                categoria = "Key Account"
            elif "CONTA_ESTRATEGICA" in KA:
                categoria = "City Key Account"
            else:
                categoria = "Normal"


            if "SO_TEM_NO_IFOOD" in KA:
                contrato = "Exclusivo"
            else:
                contrato = "Não Exclusivo"

            try:
                data = newDictionary_BM['props']['initialState']['restaurant']['details']['groups']
            except KeyError as error:
                data = "-"
            bm = []
            for i in data:
                if i['type'] == 'BUSINESS_MODEL':
                    bm = i['name']
            try:
                sr = newDictionary_BM['props']['initialState']['restaurant']['details']['superRestaurant']
            except KeyError as error:
                sr = "-"

            try:
                numrating = newDictionary_BM['props']['initialState']['restaurant']['details']['userRatingCount']
            except KeyError as error:
                numrating = "-"

            try:
                rating = newDictionary_BM['props']['initialState']['restaurant']['details']['evaluationAverage']
            except KeyError as error:
                rating = "-"


        else:
            nome = "Link inválido"
            telefone = ""
            tipo = ""
            nomerua = ""
            bairro = ""
            CEP = ""
            Latitude = ""
            Longitude = ""
            hora= ""
            bm = ""
            categoria=""
            contrato=""
            segunda = ""
            terca = ""
            quarta = ""
            quinta = ""
            sexta = ""
            sabado = ""
            domingo = ""
            sr = ""
            rating = ""
            numrating = ""
    except KeyError as error:
        nome = "Link inválido"
        telefone = ""
        tipo = ""
        nomerua = ""
        bairro = ""
        CEP = ""
        Latitude = ""
        Longitude = ""
        hora = ""
        bm = ""
        categoria = ""
        contrato = ""
        segunda = ""
        terca = ""
        quarta = ""
        quinta = ""
        sexta = ""
        sabado = ""
        domingo = ""
        sr = ""
        rating = ""
        numrating = ""
        
    
        
    return nome, telefone, tipo, nomerua, bairro, CEP, Latitude, Longitude, bm, categoria,contrato, sr,rating,numrating, segunda, terca, quarta, quinta, sexta, sabado, domingo

df = pd.read_excel('links_sp.xlsx', index_col=None, header=None) #reading the excel file with the links

lista = []

total = len(df.index)

printProgressBar(0, total, prefix = 'Progress:', suffix = 'Complete', length = 50)
for index, row in df.iterrows():
    lista.append(list(getNamePhone(row[0])))
    lista[index].append(str(row[0]))
    printProgressBar(index + 1, total, prefix='Progress:', suffix='Complete', length=50)

df=pd.DataFrame(lista,columns=['Nome', 'Tel','tipo','Endereço','Bairro','CEP','Lat','Long','Business Model','Categoria','Contrato','SuperRs','rating','numrating','Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo','Link'])
#creating the excel file with all the things that we returned
df.sort_values('Nome', inplace=True)

df.to_excel("Telefones_spr_ifood.xlsx", index=False) #creating the excel file with all the things that we returned
