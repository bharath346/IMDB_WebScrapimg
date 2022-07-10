import requests,openpyxl
from bs4 import BeautifulSoup

#Coping in Excel Sheets
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title='top Rated movies'
sheet.append(['Movie Rank','Movie Name','Year of release','IMDB rating'])

#Extracting the data from website
source = requests.get("https://www.imdb.com/chart/top/")
soup = BeautifulSoup(source.text,'html.parser')
movies = soup.find('tbody',class_="lister-list")
movie = movies.find_all('tr')

for namemovie in movie:
     name=namemovie.find('td',class_="titleColumn").a.text
     rank =namemovie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
     year=namemovie.find('td',class_="titleColumn").span.text.strip('()')
     rating=namemovie.find('td',class_="ratingColumn imdbRating").strong.text
     print(rank,name,year,rating)

     #Appending into the excel sheet
     sheet.append([rank,name,year,rating])

#Saving the sheets
excel.save('IMDb_movie_Ratings.xlsx')
