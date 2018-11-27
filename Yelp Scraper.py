# Libraries needed -imports
import requests
from bs4 import BeautifulSoup
import json
import xlsxwriter
import re
import time


start = time.time()  # Timer - shows how fast the script runs.

# Variables needed
# Needed for excel
row = 0
col = 0

# Needed for pagination in YELP
search_term = 0
location = 0
page = 0
counter = 0
global finding_data
finding_data = True

# List of copanies gathered
companies_list=[]

# Func creating output file
def creating_output_file():
    global output_file
    global worksheet
    global col
    global row
    output_file = xlsxwriter.Workbook ('Results - '+str(search_term)+ ' in '+ str(location)+'.xlsx')
    worksheet = output_file.add_worksheet ()
    bold = output_file.add_format({'bold': True})

    worksheet.write (row, col, 'Company Name', bold)
    col = 1
    worksheet.write (row, col, 'Address', bold)
    col = 2
    worksheet.write (row, col, 'Phone number', bold)
    col = 3
    worksheet.write (row, col, 'Type of business', bold)
    col = 4
    worksheet.write (row, col, 'Number of reviews', bold)
    col = 5
    worksheet.write (row, col, 'Rating', bold)
    col = 6
    worksheet.write (row, col, 'Description', bold)
    col = 7
    worksheet.write (row, col, 'Price Range', bold)
    row += 1

# FUNC Getting page and making it Soup object
def making_soup(url):
    page = requests.get(url)
    global soup
    soup = BeautifulSoup(page.content, 'html.parser')

# This func is asking for keyword and location
def search_terms_input():
    global search_term
    global location
    search_term = input('\nType here your search term(ex. tacos, dinner, sushi restaurants etc.): ')
    location = input('\nType here the location you are interested in(ex. San Francisco): ')

# Finding companies
def companies_search():
    global finding_data
    global page

    while finding_data == True:
        finding_data = False
        making_soup('https://www.yelp.com/search?find_desc=' + str(search_term.replace(" ", "+")) + '&find_loc=' + str(location.replace(" ", "+"))+'&start='+str(page))
        page += 10
        for company_pages in soup.find_all('h3', class_="search-result-title"):
            for titles in company_pages(href=re.compile('/biz/')):
                companies_list.append('https://www.yelp.com'+str(titles['href']))
                finding_data = True
    else:
        print('Found all the results!')

def getting_details():
    for page in companies_list:
        making_soup(page)
        global col
        global row
        global counter
        info = soup.find(type="application/ld+json")

        # Getting all attributes of the page
        for items in info.children:
            attr = json.loads(items)

            # Title
            try:
                col = 0
                title = attr['name']
                print(title)
                worksheet.write (row, col, title)

            except:
                worksheet.write(row, col, 'Not available')

            # Address
            try:
                col = 1
                address = attr['address']['streetAddress']+", "+attr['address']['addressLocality']+", "+attr['address']['addressRegion']+", "+attr['address']['postalCode']+", "+attr['address']['addressCountry']
                worksheet.write(row, col, str(address))

            except:
                worksheet.write(row, col, 'Not available')

            # Phone
            try:
                col = 2
                phone = attr['telephone']
                worksheet.write(row, col, phone)

            except:
                worksheet.write(row, col, 'Not available')

            # Type
            try:
                col = 3
                type = attr['@type']
                worksheet.write(row, col, type)

            except:
                worksheet.write(row, col, 'Not available')

            # Review count
            try:
                col = 4
                review_count = attr['aggregateRating']['reviewCount']
                worksheet.write (row, col, review_count)

            except:
                worksheet.write(row, col, 'Not available')

            # Rating value
            try:
                col = 5
                rating_value = attr['aggregateRating']['ratingValue']
                worksheet.write(row, col, rating_value)

            except:
                worksheet.write(row, col, 'Not available')

            # Description
            try:
                col = 6
                description = soup.find(property="og:description")['content']
                worksheet.write(row, col, description)

            except:
                worksheet.write(row, col, 'Not available')

            # Price range
            try:
                col = 7
                price_range = soup.find(attr['priceRange'])
                worksheet.write (row, col, price_range)

            except:
                worksheet.write (row, col, 'Not available')

        row += 1
        counter += 1

        # Showing progress of scraping
        progress = ((counter/len(companies_list))*100.00)
        print('%.2f'% progress + '% ready\n')

# Script start
print('\nHello! This is a Yelp.com scraper. \nIt gets all details for a company such as Company name, \nAddress, Phone, Type, Number of reviews, Rating and description.')

# Getting keyword and location input
search_terms_input()
creating_output_file()
print('\nSearching for '+str(search_term)+' in '+ str(location))

# Searching for companies
companies_search()
print('\nCompanies found: ' + str(len(companies_list)))

# Getting all data for all companies found
print('\nGetting companies info...')
getting_details()


# End of script
output_file.close()
print('\nAll done!')
end = time.time()
print("\nTotal time for running:")
print(end - start)
