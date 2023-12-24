import requests
import pandas as pd
import re
import yake
from os import makedirs


def filtering(string: str) -> list:
    string = ' ' + string.upper() + ' '
    string = string.replace("'s ", ' ').replace(' L.L.C ', ' ').replace(' L.L.C. ', ' ').replace(' LLC. ', ' ').replace(
        ' LLC ', ' ').replace(' INC. ', ' ').replace(' INC ', ' ').replace(' LIMITED ', ' ').replace(' LTD. ',
                                                                                                     ' ').replace(
        ' LTD ', ' ').replace('CORPORATION', ' ').replace(' CORP. ', ' ').replace(' CORP ', ' ').replace('COMPANY',
                                                                                                         ' ').replace(
        ' COMP. ', ' ').replace(' COMP ', ' ').replace(' CO. ', ' ').replace(' CO ', ' ').replace('ENTERPRISES',
                                                                                                  ' ').replace(
        'ENTERPRISE', ' ').replace(' ENTER. ', ' ').replace(' ENTER ', ' ').replace(' ENT. ', ' ').replace(' ENT ',
                                                                                                           ' ').replace(
        'GROUP', ' ').replace(' GRP ', ' ')
    string = ''.join([i for i in string if (not i.isdigit()) and (i.isalnum() or i == ' ')])
    string = string.lower()
    string = string.strip()
    string = re.sub(' +', ' ', string)
    keywords = yake.KeywordExtractor(lan="en", n=5, dedupLim=0.9, dedupFunc='seqm', windowsSize=1, top=20,
                                     features=None)
    return keywords.extract_keywords(string)


CLEARBIT_KEY = "Your key"


def get_domain(company_name: str) -> str:
    url = f'https://company.clearbit.com/v1/domains/find?name={company_name}'
    try:
        response = requests.get(url, auth=(CLEARBIT_KEY, ''), timeout=5)
        if response.ok:
            data = response.json()
            return data['domain']
        else:
            # print(f'{company_name} Error:', response.status_code)
            return None
    except Exception as e:
        print(f'{company_name} Error:', e)
        return None


def get_domain_using_keyword(keywords: list) -> str:
    for keyword in keywords:
        domain = get_domain(keyword[0])
        if domain:
            print(f"Name: {keyword[0]}, Domain: {domain}")
            return domain
    #print(f"Name: {keywords[0]}, Domain: Not Found")
    return None


# reading the company names from xlsx file
df = pd.read_excel('for-clearbitAPI.xlsx')

# storing all company names in a list
customers = df['Customer Name'].tolist()
# FOR bATCH PROCESSING
INCREMENT = 1000  # number of records to be processed at a time
START = 12700  # starting index of the list (just increment this by `INCREMENT` to get the next `INCREMENT` records)
END = START + INCREMENT if START + INCREMENT < len(customers) else len(
    customers)  # ending index of the list (no need to change this value)
END = 12701
customers = customers[START:END]

# creating a list of dictionary objects
domains = []
for customer in customers:
    try:
        keywords = filtering(customer)
    except Exception as e:
        print(f'Error: {e}')
        continue
    domain = get_domain_using_keyword(keywords)
    if domain:
        domains.append({
            'Customer Name': customer,
            'Domain': domain
        })
    else:
        domains.append({
            'Customer Name': customer,
            'Domain': 'Not Found'
        })

# make a directory to store the output file
makedirs('results', exist_ok=True)

# creating a dataframe from the list of dictionary objects
df = pd.DataFrame(domains)
df.to_excel(f"results/output_{START}-{END}.xlsx", index=False)