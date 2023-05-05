from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title =  'top rated movies'

sheet.append(['Movie Rank','Movie Name','Year of Release','IMDB Rating'])

try:
    url = requests.get('https://www.imdb.com/chart/top/')
    url.raise_for_status()
    soup = BeautifulSoup(url.text,'lxml')
    movies = soup.find('tbody',class_="lister-list")
    mov = movies.find_all('tr')

    for movies in mov:

        movie_name = movies.find('td', class_ = "titleColumn").a.text
        rank = movies.find('td', class_ = "titleColumn").get_text(strip = True).split('.')[0]
        year = movies.find('td', class_ = "titleColumn").span.text.strip('()')
        rating = movies.find('td', class_ ="ratingColumn imdbRating").text
        print(movie_name,rank,year,rating)
        sheet.append([movie_name,rank,year,rating])
        

except Exception as e:
    print(e)


excel.save('movie_rating.xlsx')