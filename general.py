from selenium import webdriver
import requests
from bs4 import BeautifulSoup
import time
from openpyxl import Workbook, load_workbook


def otworz_strone_selenium(adres_www):
    strona = webdriver.Chrome('C:\Python27\Scripts\chromedriver.exe')
    # strona.implicitly_wait(7)
    strona.get(adres_www)
    return strona


def otworz_strone_soap(adres_www):
    strona = requests.get(adres_www)
    stronaa = BeautifulSoup(strona.content, 'html.parser')
    return stronaa


def ilosc_obiektow_selenium(strona, css):
    return len(strona.find_elements_by_css_selector(css))


def ilosc_obiektow_soap(strona, css):
    return len(strona.select(css))


def tytul_selenium(strona, numer_obiektu, css, css_tytul):
    css_nth = css + ":nth-of-type(" + str(numer_obiektu + 1) + ")"
    if len(css_tytul) > 0:
        return strona.find_element_by_css_selector(css_nth + ' > ' + css_tytul).text.lower()
    return strona.find_element_by_css_selector(css_nth).text.lower()


def tytul_soap(strona, numer_obiektu, css, css_tytul):
    css_nth = css + ":nth-of-type(" + str(numer_obiektu + 1) + ")"
    if len(css_tytul) > 0:
        return strona.select(css_nth)[numer_obiektu].select(css_tytul)[0].text.lower()
    return strona.select(css_nth)[numer_obiektu].text.lower()


def adres_selenium(podst_adress, strona, numer_obiektu, css, css_adres):
    css_nth = css + ":nth-of-type(" + str(numer_obiektu + 1) + ")"
    if len(css_adres) > 0:
        return podst_adress + strona.find_element_by_css_selector(css_nth + ' > ' + css_adres).get_attribute('href')
    else:
        return podst_adress + strona.find_element_by_css_selector(css_nth).get_attribute('href')


def adres_soap(podst_adres, strona, numer_obiektu, css, css_adres):
    css_nth = css + ":nth-of-type(" + str(numer_obiektu + 1) + ")"
    if len(css_adres) > 0:
        return podst_adres + strona.select(css_nth)[0].select(css_adres, href=True)[0]['href']
    else:
        return podst_adres


def odczyt_selenium():
    dane_strony = otworz_strone_selenium(www_strona)
    ilosc_obiektow = ilosc_obiektow_selenium(dane_strony, css_strona)
    for nr in range(ilosc_obiektow):
        lista_tytulow.append(tytul_selenium(dane_strony, nr, css_strona, css_tytul))
        lista_adresow.append(adres_selenium(base_url, dane_strony, nr, css_strona, css_adres))


def odczyt_soap():
    dane_strony = otworz_strone_soap(www_strona)
    ilosc_obiektow = ilosc_obiektow_soap(dane_strony, css_strona)
    for nr in range(ilosc_obiektow):
        lista_tytulow.append(tytul_soap(dane_strony, nr, css_strona, css_tytul))
        lista_adresow.append(adres_soap(base_url, dane_strony, nr, css_strona, css_adres))


def odczyt_historii():
    historia_tytulow = open("historia.txt", "r")
    # print(historia_tytulow.readlines())
    # return historia_tytulow.readlines()
    return historia_tytulow.read()


def zapis_historii(nr):
    historia_tytulow = open("historia.txt", "a+")
    historia_tytulow.write(lista_tytulow[nr] + '\n')
    historia_tytulow.close()


def zapis_raportu_txt(nr):
    historia_tytulow = open("raport_" + aktualna_data + ".txt", "a+")
    historia_tytulow.write(lista_tytulow[nr] + '\n')
    historia_tytulow.close()


def zapis_raportu(nr):
    plik_raportu = load_workbook("Raport_" + aktualna_data + ".xlsx")
    wyniki_raportu = plik_raportu.active
    wyniki_raportu.cell(row=raport_kolumna, column=1).value = '=HYPERLINK("{}", "{}")'.format(lista_adresow[nr], ">Link<")
    wyniki_raportu.cell(row=raport_kolumna, column=2).value = lista_tytulow[nr]
    plik_raportu.save("Raport_" + aktualna_data + ".xlsx")


def tworzenie_raportu():
    try:
        load_workbook("Raport_" + aktualna_data + ".xlsx")
        raport_kolumna += 1
    except:
        plik_raportu = Workbook()
        plik_raportu.save("Raport_" + aktualna_data + ".xlsx")
        raport_kolumna = 1


def sprawdzenie_historii():
    for nr, tytul in enumerate(lista_tytulow):
        if tytul not in odczyt_historii():
            tworzenie_raportu()
            zapis_historii(nr)
            zapis_raportu(nr)


def start_programu(tryb):
    if tryb == "selenium":
        odczyt_selenium()
        sprawdzenie_historii()
    if tryb == "soap":
        odczyt_soap()
        sprawdzenie_historii()


aktualna_data = time.strftime("%Y_%m_%d")
lista_tytulow = []
lista_adresow = []
www_strona = "https://bip.sobotka.pl/zamowienia/lista/13.dhtml"
css_strona = "#lista_zamowien > tbody > tr"
css_tytul = "td:nth-child(1) > a"
# css_tytul = ""
css_adres = "td:nth-child(1) > a"
# css_adres = ""
# base_url = "www.dupa.com"
base_url = ""


start_programu("selenium")

# print(len(lista_adresow))
