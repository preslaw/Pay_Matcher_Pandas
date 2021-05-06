import re
import os
import csv
import pandas as pd
import datetime


def nameConsist(partOfName):
    pattern = re.compile(partOfName)
    for a in (os.listdir()):
        if pattern.search(a):
            return a
    return None


def loadFile(partOfName, encoder, delimiterSign):
    obligatoryFiles = ['allegro', 'faktury', 'paragony'] # brak tych plików powoduje wyjście z programu
    if nameConsist(partOfName):
        tempFile = open(nameConsist(partOfName), encoding = encoder)
        tempDict = csv.DictReader(tempFile, delimiter = delimiterSign)
        tempList = list(tempDict)
    else:
        print (f'\n! Brak pliku {partOfName}\n')
        if partOfName in obligatoryFiles:
            print(f'proszę wrzuc plik {partOfName} do katalogu i uruchom program ponownie')
            input('nacisnij przycisk')
            os._exit(0)
        else:
            tempList = []
            return tempList
    return tempList

def przystosujListyDoPolaczenia(allegroLista, p24Lista, payuLista, paragonyLista,fakturyLista):

    # paragonyLista przygotowuję pod floata i dopisuje rząd "Zamówienie" z numerem z ['Uwagi']
    for row in paragonyLista:
        row['Wartość'] = float(row['Wartość'].replace(',', '.'))
        row['Zamówienie'] = row['Uwagi'][0:4]

    # fakturyLista wartość przygotowuję pod floata
    for row in fakturyLista:
        row['Wartość'] = float(row['Wartość'].replace(',', '.'))

    # payuLista przerabiam kwoty pod float (usuwam ' zl')
    for row in payuLista:
        row['kwota'] = float(row['kwota'][:-3])

    # p24Lista przerabiam kwoty pod float (usuwam ' zl')
    for row in p24Lista:
        row['kwota'] = float(row['kwota'][:-3])

    #allegroLista
    # przerabiam kwoty pod float (usuwam ' zl' i zamieniam ',' na '.')
    for row in allegroLista:
        row['Kwota']=row['Kwota'][:-3].replace(',','.')
    #dodaję klucz 'SumaKwot' wartość to zsumowana wartość PLN o tych samych numerach w 'Zamówienie'
    allegroListKwotaSum = {}
    for row in allegroLista:
        if row['Numer wpłaty'][:-2] not in allegroListKwotaSum:
            try:
                allegroListKwotaSum[row['Numer wpłaty'][:-2]]=float(row['Kwota'])
            except ValueError:
                #continue
                allegroListKwotaSum['Numer wpłaty'] = 'Kwota','ID zew. płatności'
        else:
            allegroListKwotaSum[row['Numer wpłaty'][:-2]]+= float(row['Kwota'])

def polacz(allegroLista, p24Lista, payuLista, paragonyLista, fakturyLista):
    # tworze polaczona - to tabela wynikowa
    polaczona = []

    # dodaję do tabeli połączona rekordy z p24
    for a in p24Lista:
        tempDict = {}
        tempDict['operator'] = a['operator']
        tempDict['data'] = a['data']
        tempDict['identyfikator'] = a['identyfikator']
        tempDict['kwota'] = a['kwota']
        # polaczona += [tempDict.items()]
        polaczona += [tempDict]

    # dodaję do tabeli polaczona rekordy z payuList
    for a in payuLista:
        tempDict = {}
        tempDict['operator'] = a['operator']
        tempDict['data'] = a['data']
        tempDict['identyfikator'] = a['identyfikator']
        tempDict['kwota'] = a['kwota']
        # polaczona += [tempDict.items()]
        polaczona += [tempDict]

    # dodaję do polaczonej numer zamówienia ['Zamówienie'] z allegroList
    # klucz polaczona['identyfikator'] = allegroList['ID zew. płatności']

    for a in polaczona:
        for b in allegroLista:
            if a['identyfikator'] == b['ID zew. płatności']:
                a['nr.zam'] = b['Zamówienie']
                break
            else:
                a['nr.zam'] = ['n/a']

    # dodaję do połączonej nr faktury z fakturylist
    # klucz polaczona['Zamówienie'] = fakturyList['uwagi'[0:5]]
    for a in polaczona:
        for b in fakturyLista:
            if a['nr.zam'] == b['Uwagi'][0:5]:
                a['nr.dok'] = b['Numer']
                a['Symbol kontrahenta'] = b['Symbol kontrahenta']
                break
            else:
                a['nr.dok'] = 'n/a'
                a['ID'] = 'n/a'
    # dodaję do połączonej nr paragonu z paragonyList (jeżeli nie ma wystawionej faktury)
    for a in polaczona:
        for b in paragonyLista:
            if a['nr.dok'] == 'n/a':
                if a['nr.zam'] == b['Uwagi'][0:5]:
                    a['nr.dok'] = b['Numer']
                    a['ID'] = 'brak'

    return polaczona

def zapiszXLS(polaczonaFinal):
    skonwertowanaPolaczonaFinal = pd.DataFrame.from_dict(polaczonaFinal)
    print('WYNIK:')
    print(skonwertowanaPolaczonaFinal.to_string(max_rows=6, max_cols=4))
    outputFileName = '\Output ' + minDateFaktury + ' ' + maxDateFaktury + '.xlsx'
    skonwertowanaPolaczonaFinal.to_excel(os.path.abspath(os.getcwd())+outputFileName)
    print((f'\nPlik {outputFileName} został zapisany\n'))
    input('Naciśnij enter...')

def zwrocZakresDat(lista,kolumnaZDatami,formatDaty):
    if len(lista) ==0:
        # return None, None
        return 'brak pliku', 'brak pliku'


    daty = []
    for row in lista:
        try:
            data = datetime.datetime.strptime(row[kolumnaZDatami], formatDaty)
            data = data.strftime('%d-%m-%Y')
            daty.append(data)
        except ValueError:
            continue
    maxDate = max(daty)
    minDate = min(daty)
    return minDate, maxDate

def pokazZakresyDat():
    print('\nZakresy dat w plikach:\n')
    print('allegro:  ', minDateAllegro,' - ', maxDateAllegro)
    print('P24:      ', minDateP24,' - ', maxDateP24)
    print('payu:     ', minDatePayu,' - ', maxDatePayu)
    print('faktury:  ', minDateFaktury,' - ', maxDateFaktury)
    print('paragony: ', minDateParagony,' - ', maxDateParagony,'\n')


allegroLista = loadFile('allegro', 'ANSI', ';')
minDateAllegro, maxDateAllegro = zwrocZakresDat(allegroLista,'Data ostatniej operacji/transakcji','%Y-%m-%d %H:%M')
p24Lista = loadFile('p24', 'utf8', ',')
minDateP24, maxDateP24 = zwrocZakresDat(p24Lista,'data','%d.%m.%Y %H:%M')
payuLista = loadFile('payu', 'utf8', ',')
minDatePayu, maxDatePayu = zwrocZakresDat(payuLista,'data','%d.%m.%Y %H:%M')
fakturyLista = loadFile('faktury', 'ANSI', ';')
minDateFaktury,maxDateFaktury = zwrocZakresDat(fakturyLista,'Data','%Y-%m-%d')
paragonyLista = loadFile('paragony', 'ANSI', ';')
minDateParagony,maxDateParagony = zwrocZakresDat(paragonyLista,'Data','%Y-%m-%d' )

pokazZakresyDat()
przystosujListyDoPolaczenia(allegroLista, p24Lista, payuLista, paragonyLista, fakturyLista)
polaczonaFinal = polacz(allegroLista, p24Lista, payuLista, paragonyLista, fakturyLista)
zapiszXLS(polaczonaFinal)