from requests import get
import pandas as pd

url_query = 'https://regisonline.de/r3/rest/regis/id?query=%7B%22unternehmenQuery%22%3A%7B%22localeId%22%3A%22de%22%2C%22aenderung-von%22%3A%7B%22%40xsi.nil%22%3A%22true%22%7D%2C%22aenderung-bis%22%3A%7B%22%40xsi.nil%22%3A%22true%22%7D%2C%22betriebsgroessenklasse-liste%22%3A%7B%22betriebsgroessenklasse%22%3A%5B%2250%3A99%22%2C%22100%3A199%22%2C%22200%3A299%22%2C%22300%3A399%22%2C%22400%3A499%22%2C%22500%3A599%22%2C%22600%3A699%22%2C%22700%3A799%22%2C%22800%3A899%22%2C%22900%3A999%22%2C%221000%3A%22%5D%7D%2C%22regiobranche-liste%22%3A%7B%22regiobranche%22%3A%5B58%2C61%2C60%2C57%2C59%2C62%2C64%5D%7D%2C%22gebiet-liste%22%3A%22%22%2C%22text%22%3A%22%22%2C%22has-jobangebote-linklist%22%3Afalse%2C%22has-ausbildungsplaetze-linklist%22%3Afalse%2C%22has-praktika-linklist%22%3Afalse%2C%22has-sucht-mitarbeiter%22%3Afalse%2C%22has-bildet-aus%22%3Afalse%2C%22has-bietet-pa-an%22%3Afalse%2C%22geschaeftsfelder-fulltext%22%3A%7B%22%40xsi.nil%22%3A%22true%22%7D%2C%22ausbildungsberufe-fulltext%22%3A%7B%22%40xsi.nil%22%3A%22true%22%7D%2C%22praktika-abschlussarbeiten-fulltext%22%3A%7B%22%40xsi.nil%22%3A%22true%22%7D%2C%22zuordnungen-liste%22%3A%22%22%2C%22umkreis%22%3A%22%22%7D%7D'
r = get(url_query)
id_list = r.json()['id-list']['id']

df = pd.DataFrame(columns=['Unternehmen', 'Mitarbeiter', 'Homepage', 'Info', 'Kontaktperson'])

for unt in id_list[30:50]:
    company = get(f'https://regisonline.de/r3/rest/regis/{unt}?level=3').json()['unternehmen']

    if type(company['name']['value']) == list:
        unternehmen = company['name']['value'][0]['$']
    else:
        unternehmen = company['name']['value']['$']

    if type(company['link-list']['link']) == list:
        homepage = company['link-list']['link'][0]['url']['value']['$']
    else:
        if type(company['link-list']['link']['url']['value']) == list:
            homepage = company['link-list']['link']['url']['value'][0]['$']
        else:
            homepage = company['link-list']['link']['url']['value']['$']
        if type(company['link-list']['link']) == list:
            homepage = company['link-list']['link'][0]['url']['value'][0]['$']

    if type(company['infoFertigungDienstleistung']['value']) == list:
        info = company['infoFertigungDienstleistung']['value'][0]['$']
    else:
        info = company['infoFertigungDienstleistung']['value']['$']

    telefon = ''
    kontaktpersonen = ''
    anrede = ''
    nachname = ''
    vorname = ''
    funktion = ''

    if type(company['kontaktperson-mit-kategorie-list']['kontaktperson-mit-kategorie']) == list:
        for kontaktperson in company['kontaktperson-mit-kategorie-list']['kontaktperson-mit-kategorie']:
            if kontaktperson['kontaktperson'].get('telefon') is not None:
                telefon = kontaktperson['kontaktperson'].get('telefon').get('@converted')

            if kontaktperson['kontaktperson'].get('funktion'):
                if type(kontaktperson['kontaktperson'].get('funktion').get('value')) == list:
                    funktion = kontaktperson['kontaktperson'].get('funktion').get('value')[0].get('$')
                else:
                    funktion = kontaktperson['kontaktperson'].get('funktion').get('value').get('$')

            text = f"{kontaktperson['kontaktperson'].get('anrede')} " \
                   f"{kontaktperson['kontaktperson'].get('titel')} " \
                   f"{kontaktperson['kontaktperson'].get('vorname')} " \
                   f"{kontaktperson['kontaktperson'].get('nachname')}, " \
                   f"{funktion}, " \
                   f"Tel: {telefon}"

            kontaktpersonen += text + "\n"

    df = df.append({
         'Unternehmen': unternehmen,
         'Mitarbeiter': company.get('beschaeftigtenzahl'),
         'Homepage': homepage,
         'Info': info,
         'Kontaktperson': kontaktpersonen
    }, ignore_index=True)

    print(df)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('data_fetch.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

