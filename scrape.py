from bs4 import BeautifulSoup

import requests
import openpyxl


excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "IMDB Top 250"
print(excel.sheetnames)

sheet.append(["Movie Rank", "Movie Name", "Year of Release", "IMDB Rating"])


try:
    url = "https://www.imdb.com/chart/top/?ref_=nv_mv_250"
    page = requests.get(url)
    
    soup = BeautifulSoup(page.text, "html.parser")
   
    movies = soup.find("tbody", class_="lister-list").find_all("tr")

    for film in movies:

        rank = film.find("td", class_="titleColumn").get_text(strip=True).split(".")[0]

        name = film.find("td", class_="titleColumn").a.text
       
        year = film.find("td", class_="titleColumn").span.text.strip("()")

        imdb_rating = film.find("td", class_="ratingColumn imdbRating").strong.text

        print(rank, name, year, imdb_rating)
        sheet.append([rank, name, year, imdb_rating])
        

except Exception as x:
    print(x)


excel.save('IMDB Movie Ratings.xlsx')