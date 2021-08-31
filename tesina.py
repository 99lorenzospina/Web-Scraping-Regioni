import re # to handle regular expressions
from datetime import datetime # data formats conversion
import pandas as pd # to handle output
from selenium import webdriver
from selenium.webdriver.common.keys import Keys # to simulate keyboard and mouse interactions
from selenium.webdriver.chrome.options import Options # options to browse, e.g. opening an Incognito page
from selenium.webdriver.common.by import By # required to search items
from selenium.webdriver.support import expected_conditions as ec # required into explicit Waits
from selenium.webdriver.support.ui import Select # interactions with certain web items
from selenium.webdriver.support.ui import WebDriverWait # Waits
import time # to pause the code for a certain time
from dataclasses import dataclass # for data classes
import locale

locale.setlocale(locale.LC_TIME, "it_IT") # to use italian names for months

pd.set_option("display.max_rows", None, "display.max_columns", None)  # display any number of rows and columns
pd.set_option('display.max_colwidth', None) # no maximum width
opt = Options()  # driver execution features
opt.headless = True  # opening Chrome page without GUI
CHROMEDRIVERPATH = 'C:\Program Files\ChromeDriver\chromedriver.exe' # path of the WebDriver tool

@dataclass
class Result:
    
    def __init__(self, bollettino=None, titolo=None, descrizione=None, tipologia=None, data_pubblicazione=None, data_apertura=None, data_scadenza=None, beneficiari=None, proponente=None, importo=None, stato=None, più_informazioni=None):
        self.titolo = titolo
        self.descrizione = descrizione
        self.data_pubblicazione = data_pubblicazione
        self.data_scadenza = data_scadenza
        self.bollettino = bollettino
        self.data_apertura = data_apertura
        self.beneficiari = beneficiari
        self.più_informazioni = più_informazioni
        self.tipologia = tipologia
        self.proponente = proponente
        self.importo = importo
        self.stato = stato

def elaborazione_risultato(risultato, f, pdf=None):
    v = vars(risultato)
    for d in v:
        if d == 'più_informazioni'and pdf != None:
            s = ''
            for p in pdf:
                 s +='Link al PDF:' + '\n' + p + '\n'
        else:
            s = d.replace('_',' ').upper() + ':' + '\n' + v[d] + '\n'
        f.write(s)
    f.write('\n')


def contributi_finanziamenti_abruzzo(n):
    chrome = webdriver.Chrome(CHROMEDRIVERPATH, options=opt)
    chrome.get('https://www.regione.abruzzo.it/contributi-finanziamenti')
    titoli = []
    categorie = []
    scadenze = []
    descrizione = []
    pubblicazione = []
    link_allegati = []
    f = open('contributi_finanziamenti_abruzzo.txt', 'w', encoding="utf-8")
    i = 1
    while(i<n):
        risultati = WebDriverWait(chrome, 20).until(
            ec.visibility_of_all_elements_located(
                (By.CLASS_NAME, 'views-field.views-field-body')))
        for ris in risultati:
            t = ris.find_element_by_tag_name('a').text
            titoli.append(ris.find_element_by_tag_name('a').text)
            categoria = ris.find_element_by_class_name('post-category').text
            categorie.append(ris.find_element_by_class_name('post-category').text)
            scadenza = ris.find_element_by_class_name('date-display-single').text
            scadenze.append(ris.find_element_by_class_name('date-display-single').text)
            link_da_aprire = ris.find_element_by_tag_name('a')
            link_da_aprire.send_keys(Keys.CONTROL + Keys.RETURN)
            chrome.switch_to.window(chrome.window_handles[1])
            data = WebDriverWait(chrome, 20).until(
            ec.visibility_of_element_located(
                (By.CLASS_NAME, 'post-date')))
            data = data.text
            numRegEx = re.compile(r'\d+ \D+ \d+')
            match = numRegEx.search(data)
            date = datetime.strptime(match.group(0), '%d %B %Y')
            data_p = date.strftime('%d/%m/%Y')
            pubblicazione.append(data_p)
            desc = chrome.find_element_by_class_name('field-item.even').text
            descrizione.append(chrome.find_element_by_class_name('field-item.even').text)
            documenti = chrome.find_element_by_class_name(
                'field.field-name-field-documenti-contributi.field-type-file.field-label-above')
            documenti = documenti.find_elements_by_tag_name('a')
            pdf = []
            for documento in documenti:
                href = documento.get_attribute('href')
                pdf.append(href)
            link_allegati.append(pdf)
            chrome.close()
            chrome.switch_to.window(chrome.window_handles[0])# torno a pagina precedente
            risultato = Result('non disponibile', t, desc, categoria, data_p, 'non disponibile',
            scadenza, 'non disponibile', 'non disponibile', 'non disponibile', 'non disponibile', pdf)
            elaborazione_risultato(risultato, f, pdf)
        try:
            chrome.find_element_by_link_text('seguente ›').click() #pag successiva
        except:
            break
    f.close()
    return pd.DataFrame({'Titolo': titoli, 'Descrizione': descrizione, 'Pubblicazione': pubblicazione,
    'Categoria': categorie, 'Scadenza': scadenze, "Più informazioni": link_allegati})

#df = contributi_finanziamenti_abruzzo(3)
#df.to_excel(r'C:\Users\super\Desktop\abruzzo.xlsx', index=False, header=True)

def bollettino_abruzzo(n):
    chrome = webdriver.Chrome(CHROMEDRIVERPATH, options=opt)
    chrome.get('http://bura.regione.abruzzo.it/DEFAULT.ASPX')
    i = 1
    f = open('bollettino_abruzzo.txt', 'w', encoding="utf-8")
    numeri = []
    pubblicazione = []
    categoria = []
    descrizione = []
    link_allegati = []
    selezione = Select(chrome.find_element_by_id('ctl00_anno')) #devo impostare un anno per i bollettini
    selezione.select_by_visible_text('2021')
    chrome.find_element_by_name('ctl00$ctl01').click()
    while(i<n):
        risultati = WebDriverWait(chrome, 20).until(
                ec.visibility_of_all_elements_located(
                    (By.PARTIAL_LINK_TEXT, 'N.')))
        for ris in risultati:
            titolo = ris.text
            numero = titolo.split(' ')[0]
            data = titolo.split(' ')[2]
            tipo = titolo.split(' ')[3]
            numeri.append(numero)
            pubblicazione.append(data)
            categoria.append(tipo[1:-1])
            ris.send_keys(Keys.CONTROL + Keys.RETURN)
            chrome.switch_to.window(chrome.window_handles[1])
            documenti = WebDriverWait(chrome,10).until(
                ec.visibility_of_all_elements_located(
                    (By.XPATH,'//a[contains(@target,"new")]')
                ))
            pdf = []
            for documento in documenti:
                pdf.append(documento.get_attribute('href'))
            link_allegati.append(pdf)
            chrome.close()
            chrome.switch_to.window(chrome.window_handles[0])# torno a pagina precedente
        risultati = chrome.find_elements_by_xpath('//font[@class="clsricpi"]')
        for ris in risultati:
            descrizione.append(ris.text)
        i = i+1
        try:
            chrome.find_element_by_xpath('//a[@href="?string=&p=' + str(i) + '"]').click() # next page
        except:
            break
    for i in range(len(numeri)):
        risultato = Result(numeri[i], 'non disponibile', descrizione[i], categoria[i], pubblicazione[i],
        'non disponibile', 'non disponibile', 'non disponibile', 'non disponibile', 'non disponibile',
        'non disponibile', link_allegati[i])
        elaborazione_risultato(risultato, f, pdf)
    f.close()
    return pd.DataFrame({'Bollettino': numeri, 'Descrizione': descrizione, 'Pubblicazione': pubblicazione,
    'Categoria': categoria, "Più informazioni": link_allegati})

#df = bollettino_abruzzo(3)
#df.to_excel(r'C:\Users\super\Desktop\abruzzo(1).xlsx', index=False, header=True)


def bandi_campania(n):
    chrome = webdriver.Chrome(CHROMEDRIVERPATH, options=opt)
    chrome.get('https://pgt.regione.campania.it/portale/index.php/bandi')
    Descrizioni = []
    Proponenti = []
    Appaltanti = []
    Importi = []
    Scadenze = []
    Tipi = []
    Dettagli = []
    i = 1
    f = open('bandi_campania.txt', 'w', encoding="utf-8")
    while(i<n):
        tabella = WebDriverWait(chrome, 20).until(
                ec.presence_of_element_located(
                    (By.ID, 'table-lista-bandi')))
        mashup = tabella.find_elements_by_tag_name('td')
        j = 0
        for ris in mashup:
            if j == len(mashup):
                break
            pos = j%7
            if pos == 0:
                descrizione = ris.text
                Descrizioni.append(descrizione)
            elif pos == 2:
                proponente = ris.text
                Proponenti.append(proponente)
            elif pos == 3:
                appaltante = ris.text
                Appaltanti.append(appaltante)
            elif pos == 4:
                importo = ris.text
                Importi.append(importo)
            elif pos == 5:
                scadenza = ris.text
                Scadenze.append(scadenza)
            elif pos == 6:
                dettaglio = ris.find_element_by_tag_name('a')
                det = dettaglio.get_attribute('href')
                Dettagli.append(det)
                dettaglio.send_keys(Keys.CONTROL + Keys.RETURN)
                chrome.switch_to.window(chrome.window_handles[1])
                tipo = WebDriverWait(chrome, 20).until(
                    ec.visibility_of_element_located(
                        (By.XPATH, '//th[text()="Tipo Appalto:"]/following-sibling::td')))
                tipo = tipo.text
                Tipi.append(tipo)
                chrome.close()
                chrome.switch_to.window(chrome.window_handles[0])
                risultato = Result('non disponibile', 'non disponibile', descrizione, tipo,
                'non disponibile', 'non disponibile', scadenza, appaltante, proponente,
                importo, 'non disponibile', det)
                elaborazione_risultato(risultato, f)
            j+=1
        i += 1
        try:
            chrome.find_element_by_partial_link_text('Next').click() #pag successiva
        except:
            break
    f.close()
    return pd.DataFrame({'Descrizione': Descrizioni, 'Scadenza': Scadenze, 'Tipologia': Tipi,
    'Ente proponente': Proponenti, 'Ente appaltante': Appaltanti, 'Importo': Importi,
    'Più informazioni': Dettagli})


#df = bandi_campania(2)
#df.to_excel(r'C:\Users\super\Desktop\campania.xlsx', index=False, header=True)

def bollettino_campania(inizio, fine):
    chrome = webdriver.Chrome(CHROMEDRIVERPATH, options=opt)
    chrome.get('http://burc.regione.campania.it/eBurcWeb/publicContent/search/search.iface')
    burcs = []
    datas = []
    pdfs = []
    atti = []
    i = 1
    advanced = WebDriverWait(chrome, 10).until(
        ec.visibility_of_element_located(
            (By.CLASS_NAME, 'iceCmdLnk'))
        )
    advanced.click()
    dal = WebDriverWait(chrome, 10).until(
        ec.visibility_of_element_located(
            (By.ID, 'frmSearch:j_id126')
        )
    )
    al = chrome.find_element_by_id('frmSearch:j_id131')
    ricerca = chrome.find_element_by_id('frmSearch:j_id138')
    dal.send_keys(inizio)
    al.send_keys(fine)
    ricerca.click()
    hasbroken = False
    oldfirst = ''
    already_not_five = False
    f = open('bollettino_campania.txt', 'w', encoding="utf-8")
    while(i):
        bandi = WebDriverWait(chrome, 100).until(
            ec.visibility_of_all_elements_located(
                (By.CLASS_NAME, 'iceInpTxtArea')
            )
        )
        pdf = chrome.find_elements_by_xpath('//img[@alt="Scarica Allegato PDF"]/ancestor::a[@class="iceOutLnk"]')
        if len(bandi) != 5:
            if already_not_five:
                break
            else:
                already_not_five = True
        for ris in bandi:
            i+=1
            if (i-1)%5 == 1:
                if ris.text == oldfirst:
                    hasbroken = True
                    break
                else:
                    oldfirst = ris.text
            atto = ris.text
            atti.append(atto)
            wrapper_of_BURC = ris.find_element_by_xpath('//ancestor::tr[@class="iceDatTblRow1"]/descendant::table[@class="search"]')
            titolo = wrapper_of_BURC.find_elements_by_class_name('iceOutTxt')
            burc = titolo[0].text[0:-4]
            burcs.append(burc)
            data = titolo[1].text
            datas.append(data)
        if hasbroken:
            break
        for ris in pdf:
            p = []
            p.append(ris.get_attribute('href'))
            pdfs.append(p)
        for j in range(len(atti)):
            risultato = Result(burcs[j], atti[j], 'non disponibile', 'non disponibile',
            datas[j], 'non disponibile', 'non disponibile', 'non disponibile',
            'non disponibile', 'non disponibile', 'non disponibile', pdfs[j])
            elaborazione_risultato(risultato, f, pdfs[j])
        chrome.find_element_by_id('frmSearch:listaBurc:nextpage_1').click()
        time.sleep(2)
    f.close()
    return pd.DataFrame({'Bollettino': burcs, 'Titolo' : atti, 'Pubblicazione': datas, 'Più informazioni': pdfs})

#df = bollettino_campania('5-gen-2021','20-gen-2021')
#df.to_excel(r'C:\Users\super\Desktop\campania(1).xlsx', index=False, header=True)

def bandi_emilia(n):
    chrome = webdriver.Chrome(CHROMEDRIVERPATH, options=opt)
    chrome.get('https://bandi.regione.emilia-romagna.it/search_bandi_form')
    titoli = []
    pubblicazioni = []
    scandenze = []
    stati = []
    tipologie = []
    candidabili = []
    piùInfo = []
    f = open('bandi_emilia.txt', 'w', encoding="utf-8")
    i = 1
    while(i<n):
        risultati = WebDriverWait(chrome, 10).until(
            ec.visibility_of_all_elements_located(
                        (By.CLASS_NAME, 'bando-result')
                        )
        )
        for ris in risultati:
            stato = ris.find_element_by_tag_name('span').text # first element with "span" type
            if stato.lower() == 'chiuso':
                continue
            stati.append(stato)
            titolo = ris.find_element_by_class_name('state-published')
            link = titolo.get_attribute('href')
            piùInfo.append(link)
            titolo.send_keys(Keys.CONTROL + Keys.RETURN)
            chrome.switch_to.window(chrome.window_handles[1])
            table = WebDriverWait(chrome,10).until(
                ec.visibility_of_element_located(
                    (By.CLASS_NAME, 'vertical.listing.tableRight')
                )
            )
            td = table.find_elements_by_tag_name('td')
            tipologia = td[1].text
            tipologie.append(tipologia)
            concorrente = td[2].text
            candidabili.append(concorrente)
            chrome.close()
            chrome.switch_to.window(chrome.window_handles[0])
            titolo = titolo.find_element_by_tag_name('span').text
            titoli.append(titolo)
            date = ris.find_element_by_class_name('bandoDates').find_elements_by_tag_name('p')
            pubblicazione = date[0].find_elements_by_tag_name('span')[1].text
            pubblicazioni.append(pubblicazione)
            scadenza = date[1].find_elements_by_tag_name('span')[1].text
            scandenze.append(scadenza)
            risultato = Result('non disponibile', titolo, 'non disponibile', tipologia, pubblicazione,
            'non disponibile', scadenza, concorrente, 'non disponibile', 'non disponibile', stato, link)
            elaborazione_risultato(risultato, f)
        i += 1
        try:
            next = chrome.find_element_by_class_name('next').find_element_by_tag_name('a')
            next.send_keys(Keys.RETURN)
        except:
            break
    f.close()
    return pd.DataFrame({'Titolo' : titoli, 'Pubblicazione': pubblicazioni, 'Scadenza': scandenze,
    'Tipo': tipologia, 'Enti beneficiari': candidabili, 'Stato': stati, 'Più informazioni': piùInfo})

#df = bandi_emilia(3)
#df.to_excel(r'C:\Users\super\Desktop\emilia.xlsx', index=False, header=True)


def bollettino_emilia(inizio, fine, n):
    chrome = webdriver.Chrome(CHROMEDRIVERPATH, options=opt)
    chrome.get('https://bur.regione.emilia-romagna.it/ricerca')
    titoli=[]
    numeri = []
    link = []
    pubblicazioni = []
    info = []
    dal = WebDriverWait(chrome, 10).until(
        ec.visibility_of_element_located(
            (By.ID, 'dal'))
    )
    dal.send_keys(inizio)
    al = chrome.find_element_by_id('al')
    al.send_keys(fine)
    chrome.find_element_by_id('titleSubmit').click()
    f = open('bollettino_emilia.txt', 'w', encoding="utf-8")
    i = 0
    while(i<n):
        risultati = WebDriverWait(chrome, 10).until(
        ec.visibility_of_all_elements_located(
            (By.CLASS_NAME, 'risultati'))
        )
        for ris in risultati:
            i+=1
            titolo = ris.find_element_by_tag_name('a')
            temp = titolo.get_attribute('href')
            link.append(temp)
            numero = titolo.text
            bollettino = numero.split(' ')[0]
            numeri.append(bollettino)
            nome = numero.split('(')[1]
            titolo = nome[0:-1]
            titoli.append(titolo)
            numRegEx = re.compile(r'\d+.\d+.\d+')
            match = numRegEx.search(numero)
            date = datetime.strptime(match.group(0), '%d.%m.%Y')
            data = date.strftime('%d/%m/%Y')
            pubblicazioni.append(data)
            infos = ris.find_element_by_tag_name('p').text
            info.append(infos)
            risultato = Result(bollettino, titolo, infos, 'non disponibile', data,
            'non disponibile', 'non disponibile', 'non disponibile', 'non disponibile',
            'non disponibile', 'non disponibile', temp)
            elaborazione_risultato(risultato, f)
        try:
            chrome.find_element_by_class_name('next').click()
        except:
            break
    f.close()
    return pd.DataFrame({'Bollettino': numeri, 'Titolo': titoli, 'Descrizione': info,
    'Pubblicazione': pubblicazioni, 'Più informazioni': link})

#df = bollettino_emilia('06/08/2021', '16/08/2021', 50)
#df.to_excel(r'C:\Users\super\Desktop\emilia(1).xlsx', index=False, header=True)

        
def bandi_liguria():
    chrome = webdriver.Chrome(CHROMEDRIVERPATH, options=opt)
    chrome.get('https://www.regione.liguria.it/bandi-e-avvisi/contributi.html')
    link = []
    titoli = []
    pubblicazioni = []
    inizi = []
    scadenze = []
    beneficiari = []
    chrome.find_element_by_xpath('//a[contains(@aria-label, "allow cookies")]').click()
    time.sleep(2)
    f = open('bandi_liguria.txt', 'w', encoding="utf-8")
    i = 1
    while(i):
        risultati = WebDriverWait(chrome, 10).until(
            ec.visibility_of_all_elements_located(
                (By.CLASS_NAME, 'pc_latest_item.hoverbg')
            )
        )
        for ris in risultati:
            wrap = ris.find_elements_by_class_name('pc_latest_item_bando_titolo')
            a = wrap[0].find_element_by_tag_name('a')
            l = a.get_attribute('href')
            link.append(l)
            testo = ''
            for w in wrap:
                testo += w.text
                testo += ' '
            titoli.append(testo)
            wrap = ris.find_element_by_class_name('pc_latest_item_subbox')
            b = wrap.find_element_by_class_name('pc_latest_item_beneficiari.minisize').text
            beneficiari.append(b)
            p = wrap.find_element_by_class_name('pc_latest_item_apertura.minisize').text
            pubblicazioni.append(p)
            try:
                inizio = wrap.find_element_by_class_name('pc_latest_item_apertura_bando.minisize').text
                inizi.append(i)
            except:
                inizi.append('non disponibile')
            try:
                scadenza = wrap.find_element_by_class_name('pc_latest_item_chiusura.minisize').text
                scadenze.append(scadenza)
            except:
                scadenze.append('non disponibile')
            risultato = Result('non disponibile', testo, 'non disponibile', 'non disponibile', p,
            inizio, scadenza, 'non disponibile', 'non disponibile', 'non disponibile',
            'non disponibile', l)
            elaborazione_risultato(risultato, f)
        i+=1
        try:
            next = chrome.find_element_by_xpath('//li[@class="pageNav"][@title="Vai alla pagina ' + str(i) +'"]')
            next.click()
        except:
            break
    f.close()
    return pd.DataFrame({'Titolo': titoli, 'Pubblicazione': pubblicazioni, 'Apertura': inizi,
    'Scadenza': scadenze, 'Enti beneficiari': beneficiari, 'Più informazioni': link})

#df = bandi_liguria()
#df.to_excel(r'C:\Users\super\Desktop\liguria.xlsx', index=False, header=True)


def bollettino_liguria(numPag):
    chrome = webdriver.Chrome(CHROMEDRIVERPATH, options=opt)
    chrome.get('http://www.burl.it/ricerca-bollettino-ufficiale-liguria.html')
    numeri = []
    pubblicazioni = []
    pdf = []
    supplementi = []
    chrome.find_element_by_class_name('btn.btn-warning.popup-modal-dismiss').click()
    chrome.find_element_by_xpath('//input[contains(@type, "submit")]').click()
    paginazione = chrome.find_element_by_id('paginazioneDocumenti')
    num = int(paginazione.find_element_by_class_name('elencoNumeroPagine').text.split()[-1])
    num = min(num, numPag)
    f = open('bollettino_liguria.txt', 'w', encoding="utf-8")
    i = 0
    while(i<num):
        i+=1
        elenco = WebDriverWait(chrome, 10).until(
            ec.visibility_of_element_located(
                (By.CLASS_NAME, 'elencoRisultati')
            )
        )
        risultati = elenco.find_elements_by_tag_name('li')
        for ris in risultati:
            numero = ris.find_element_by_class_name('titoloBollettino').text
            numeri.append(numero)
            pubblicazione = ris.find_element_by_class_name('dataBollettino').text[4:]
            pubblicazioni.append(pubblicazione)
            supplemento = ris.find_element_by_class_name('supplementoBollettino').text
            supplementi.append(supplemento)
            link = ris.find_element_by_tag_name('a').get_attribute('href')
            p = []
            p.append(link)
            pdf.append(p)
            risultato = Result(numero, supplemento, 'non disponibile', 'non disponibile',
            pubblicazione, 'non disponibile', 'non disponibile', 'non disponibile',
            'non disponibile', 'non disponibile', 'non disponibile', link)
            elaborazione_risultato(risultato, f, p)
        paginazione = chrome.find_element_by_id('paginazioneDocumenti')
        next = paginazione.find_elements_by_class_name('arrowSearch')[2]
        next.click()
    f.close()
    return pd.DataFrame({'Bollettino': numeri, 'Titolo': supplementi, 'Pubblicazione': pubblicazioni,
    'Più informazioni': pdf})

#df = bollettino_liguria(5)
#df.to_excel(r'C:\Users\super\Desktop\liguria(1).xlsx', index=False, header=True)