import requests  # to get HTML through Request from the WikiPage
from bs4 import BeautifulSoup  # used to extract data from HTML
import pandas as pd  # to manipulating our tables


page_request = requests.get('https://ar.wikipedia.org/wiki/قائمة_أفضل_مئة_رواية_عربية').text
page_html = BeautifulSoup(page_request, 'html.parser')


def clean_row(full_row):
    import re
    '''
        this function takes one row of the table ('tr') and convert it to a cleaned List ready to add to our data frame
        Output: list of cleaned_data(ready to use)
    '''
    row = full_row.findChildren('th') + full_row.findChildren('td')
    clean_row = []

    for column in row:

        link = column.find('a')
        # to remove special chars from text
        text = re.sub(r'[\W_]+', ' ', column.get_text())

        if(link):
            # to Make word hyperlinked in spreadsheet
            clean_row.append(
                '=HYPERLINK("https://ar.wikipedia.org' + link['href'] + '","' + text + '")')
        else:
            clean_row.append(text)

    return clean_row


# scrabe_table
table = page_html.find('table', {'class': 'wikitable'})
table_body = table.find('tbody', recursive=False)
rows = table_body.findChildren('tr')

head = clean_row(rows[0])
cleaned_row = []
for row in rows[1:]:
    cleaned_row.append(clean_row(row))

df = pd.read_html(str(table))
df[0]


df = pd.DataFrame(cleaned_row, columns=head)
df.to_excel("final.xlsx", index=False)
