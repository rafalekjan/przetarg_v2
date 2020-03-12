from selenium import webdriver
import requests
from bs4 import BeautifulSoup


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



www_strona = "https://bip.sobotka.pl/zamowienia/lista/13.dhtml"
css_strona = "#lista_zamowien > tbody > tr"
css_tytul = "td:nth-child(1) > a"
# css_tytul = ""
css_adres = "td:nth-child(1) > a"
# css_adres = ""
base_url = "www.dupa.com"
# base_url = ""

dane1 = otworz_strone_selenium(www_strona)
dane2 = otworz_strone_soap(www_strona)
ilosc_obiektow_1 = ilosc_obiektow_selenium(dane1, css_strona)
ilosc_obiektow_2 = ilosc_obiektow_soap(dane2, css_strona)
tytul1 = tytul_selenium(dane1, 0, css_strona, css_tytul)
tytul2 = tytul_soap(dane2, 0, css_strona, css_tytul)
adres1 = adres_selenium(base_url, dane1, 0, css_strona, css_adres)
adres2 = adres_soap(base_url, dane2, 0, css_strona, css_adres)
print(tytul1)
print(adres1)
print(tytul2)
print(adres2)

