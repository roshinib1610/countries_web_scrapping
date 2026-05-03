from bs4 import BeautifulSoup
import requests, openpyxl

try:
    headers = {
        "User-Agent": "Mozilla/5.0"
    }
    
    response = requests.get("https://www.scrapethissite.com/pages/simple/", headers=headers)
    print(response.status_code)

    soup = BeautifulSoup(response.text, 'html.parser')
    countries = soup.find_all("div", class_="country")

    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Countries"

    
    ws.append(["S.No", "Country", "Capital", "Population", "Area (km2)"])

    
    for i, country in enumerate(countries, start=1):
        name = country.find("h3", class_="country-name").text.strip()
        capital = country.find("span", class_="country-capital").text.strip()
        population = country.find("span", class_="country-population").text.strip()
        area = country.find("span", class_="country-area").text.strip()

        
        print(i,"-", name,",",capital,",", population,",", area,"km2")

        
        ws.append([i, name, capital, population, area])

    
    wb.save("countries.xlsx")
    print("Excel file created successfully!")

except Exception as e:
    print(e)
