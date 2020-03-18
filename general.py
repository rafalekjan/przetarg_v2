from selenium import webdriver
import requests
from bs4 import BeautifulSoup
import time
from openpyxl import Workbook, load_workbook


def otworz_strone_selenium():
    strona = webdriver.Chrome('C:\Python27\Scripts\chromedriver.exe')
    strona.implicitly_wait(10)
    strona.get(aktualna_strona_www)
    return strona


def otworz_strone_soup():
    strona = requests.get(aktualna_strona_www)
    dane = BeautifulSoup(strona.content, 'html.parser')
    return dane


def ilosc_obiektow_selenium(strona):
    return len(strona.find_elements_by_css_selector(aktualna_css_strona))


def ilosc_obiektow_soup(strona):
    return len(strona.select(aktualna_css_strona))


def tytul_selenium(strona, numer_obiektu):
    css_nth = aktualna_css_strona + ":nth-of-type(" + str(numer_obiektu + 1) + ")"
    if len(aktualna_css_tytul) > 0:
        return strona.find_element_by_css_selector(css_nth + ' > ' + aktualna_css_tytul).text.lower()
    return strona.find_element_by_css_selector(css_nth).text.lower()


def tytul_soup(strona, numer_obiektu):
    css_nth = aktualna_css_strona + ":nth-of-type(" + str(numer_obiektu + 1) + ")"
    if len(aktualna_css_tytul) > 0:
        return strona.select(css_nth)[0].select(aktualna_css_tytul)[0].text.lower()
    return strona.select(css_nth)[0].text.lower()


def adres_selenium(strona, numer_obiektu):
    css_nth = aktualna_css_strona + ":nth-of-type(" + str(numer_obiektu + 1) + ")"
    if len(aktualna_css_adres) > 0:
        return aktualna_podst_adres + strona.find_element_by_css_selector(css_nth + ' > ' + css_adresy).get_attribute('href')
    else:
        return aktualna_podst_adres + strona.find_element_by_css_selector(css_nth).get_attribute('href')


def adres_soup(strona, numer_obiektu):
    css_nth = aktualna_css_strona + ":nth-of-type(" + str(numer_obiektu + 1) + ")"
    if len(aktualna_css_adres) > 0:
        return aktualna_podst_adres + strona.select(css_nth)[0].select(aktualna_css_adres, href=True)[0]['href']
    else:
        return aktualna_podst_adres


def odczyt_selenium():
    dane_strony = otworz_strone_selenium()
    ilosc_obiektow = ilosc_obiektow_selenium(dane_strony)
    for nr in range(ilosc_obiektow):
        lista_tytulow.append(tytul_selenium(dane_strony, nr))
        lista_adresow.append(adres_selenium(dane_strony, nr))


def odczyt_soup():
    dane_strony = otworz_strone_soup()
    ilosc_obiektow = ilosc_obiektow_soup(dane_strony)
    print("Znaleziono ogloszen:", ilosc_obiektow)
    for nr in range(ilosc_obiektow):
        lista_tytulow.append(tytul_soup(dane_strony, nr))
        lista_adresow.append(adres_soup(dane_strony, nr))


def odczyt_historii():
    historia_tytulow = open("historia.txt", "r+")
    return historia_tytulow.read()


def zapis_historii(nr):
    historia_tytulow = open("historia.txt", "a+")
    historia_tytulow.write((lista_tytulow[nr]) + '\n')
    historia_tytulow.close()


def zapis_raportu_txt(nr):
    historia_tytulow = open("raport_" + aktualna_data + ".txt", "a+")
    historia_tytulow.write(lista_tytulow[nr] + '\n')
    historia_tytulow.close()


def zapis_raportu(nr):
    plik_raportu = load_workbook("Raport_" + aktualna_data + ".xlsx")
    wyniki_raportu = plik_raportu.active
    if len(lista_adresow[nr]) > 255:
        wyniki_raportu.cell(row=len(wyniki_raportu["A"]) + 1, column=1).value = lista_adresow[nr]
    else:
        wyniki_raportu.cell(row=len(wyniki_raportu["A"]) + 1, column=1).value = '=HYPERLINK("{}", "{}")'.format(
            lista_adresow[nr], ">Link<")
    wyniki_raportu.cell(row=len(wyniki_raportu["A"]), column=2).value = lista_tytulow[nr]
    plik_raportu.save("Raport_" + aktualna_data + ".xlsx")


def tworzenie_raportu():
    try:
        load_workbook("Raport_" + aktualna_data + ".xlsx")
    except:
        plik_raportu = Workbook()
        plik_raportu.save("Raport_" + aktualna_data + ".xlsx")


def sprawdzenie_historii():
    for nr, tytul in enumerate(lista_tytulow):
        if tytul not in odczyt_historii():
            tworzenie_raportu()
            zapis_historii(nr)
            zapis_raportu(nr)


def start_programu():
    if aktualna_metoda_www == "selenium":
        odczyt_selenium()
        sprawdzenie_historii()
    if aktualna_metoda_www == "soup":
        odczyt_soup()
        sprawdzenie_historii()


def odczytaj_z_excela(excel, kol, rza):
    return excel.cell(row=kol, column=rza).value


def odczyt_danych_kolumna(plik_excel, kolumna):
    dane = []
    otwarty_excel = load_workbook(plik_excel).active
    for i in range(1, len(otwarty_excel["A"])):
        if odczytaj_z_excela(otwarty_excel, i, kolumna) is not None:
            dane.append(odczytaj_z_excela(otwarty_excel, i, kolumna))
        else:
            dane.append("")
    return dane


def sprawdz_dane_excela():
    if not len(strony_www) == len(css_strony) == len(css_tytuly) == len(css_adresy) == len(podst_adresy) == len(metody_www):
        breakpoint()


####################
####################

strony_www = odczyt_danych_kolumna("Baza_stron.xlsx", 4)
css_strony = odczyt_danych_kolumna("Baza_stron.xlsx", 5)
css_tytuly = odczyt_danych_kolumna("Baza_stron.xlsx", 8)
css_adresy = odczyt_danych_kolumna("Baza_stron.xlsx", 7)
podst_adresy = odczyt_danych_kolumna("Baza_stron.xlsx", 9)
metody_www = odczyt_danych_kolumna("Baza_stron.xlsx", 6)
gotowosc_raportowania = odczyt_danych_kolumna("Baza_stron.xlsx", 12)
sprawdz_dane_excela()

aktualna_data = time.strftime("%Y_%m_%d")
lista_tytulow = []
lista_adresow = []

print(time.strftime("%T"))
for i in range(1, 27):
    print("Strona nr:", i)
    aktualna_gotowosc_raportowania = gotowosc_raportowania[i]
    if aktualna_gotowosc_raportowania == "TAK":
        aktualna_strona_www = strony_www[i]
        aktualna_css_strona = css_strony[i]
        aktualna_css_tytul = css_tytuly[i]
        aktualna_css_adres = css_adresy[i]
        aktualna_podst_adres = podst_adresy[i]
        aktualna_metoda_www = metody_www[i]
        # print(aktualna_strona_www)
        # print(aktualna_css_strona)
        # print(aktualna_css_tytul)
        # print(aktualna_css_adres)
        # print(aktualna_podst_adres)
        # print(aktualna_metoda_www)
        # print(aktualna_gotowosc_raportowania)
        start_programu()
print(time.strftime("%T"))
