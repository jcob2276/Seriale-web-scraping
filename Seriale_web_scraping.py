# Importowanie niezbêdnych bibliotek
from bs4 import BeautifulSoup
import requests
import pandas as pd

try:
    # Pobieranie danych z strony IMDb z rankingiem top 250 filmów
    source = requests.get('https://www.imdb.com/chart/top/?ref_=nv_mv_250')
    source.raise_for_status()

    # Parsowanie strony przy u¿yciu biblioteki BeautifulSoup
    soup = BeautifulSoup(source.text, 'html.parser')

    # Znajdowanie wszystkich filmów na liœcie
    movies = soup.find('tbody', class_='lister-list').find_all('tr')

    # Inicjalizacja s³owników dla statystyk
    year_stats = {}
    rating_stats = {}
    director_stats = {}
    actor_stats = {}

    # Iteracja po filmach i zbieranie informacji
    for movie in movies:
        # Pobieranie nazwy filmu
        name = movie.find('td', class_='titleColumn').a.text

        # Pobieranie rankingu filmu
        rank = movie.find('td', class_='titleColumn').get_text(strip=True).split('.')[0]

        # Pobieranie roku produkcji filmu
        year = int(movie.find('td', class_='titleColumn').span.text.strip('()'))

        # Pobieranie oceny filmu
        rating = float(movie.find('td', class_='ratingColumn imdbRating').strong.text)

        # Pobieranie aktorów wystêpuj¹cych w filmie
        actors = movie.find('td', class_='titleColumn').find('a')['title'].split(', ')[1:]

        # Obliczanie dekady produkcji filmu
        decade = year // 10 * 10
        decade_str = str(decade)

        # Aktualizacja statystyk dla roku produkcji
        year_stats[decade_str] = year_stats.get(decade_str, 0) + 1

        # Aktualizacja statystyk dla oceny filmu
        rating_stats[rating] = rating_stats.get(rating, 0) + 1

        # Pobieranie re¿ysera filmu
        director = movie.find('td', class_='titleColumn').a['title'].split(',')[0]

        # Aktualizacja statystyk dla re¿ysera
        director_stats[director] = director_stats.get(director, 0) + 1

        # Aktualizacja statystyk dla aktorów
        for actor in actors:
            actor_stats[actor] = actor_stats.get(actor, 0) + 1

    # Sortowanie statystyk rok-producji
    sorted_year_stats = sorted(year_stats.items())

    # Sortowanie statystyk oceny filmu
    sorted_rating_stats = sorted(rating_stats.items())

    # Sortowanie statystyk re¿yserów
    sorted_director_stats = sorted(director_stats.items(), key=lambda x: x[1], reverse=True)

    # Sortowanie statystyk aktorów
    sorted_actor_stats = sorted(actor_stats.items(), key=lambda x: x[1], reverse=True)

    # Tworzenie DataFrame dla statystyk roku produkcji
    excel_data = {
        'Decade': [decade for decade, _ in sorted_year_stats],
        'Count': [count for _, count in sorted_year_stats]
    }
    df_year = pd.DataFrame(excel_data)

    # Tworzenie DataFrame dla statystyk oceny filmu
    excel_data = {
        'Rating': [rating for rating, _ in sorted_rating_stats],
        'Count': [count for _, count in sorted_rating_stats]
    }
    df_rating = pd.DataFrame(excel_data)

    # Tworzenie DataFrame dla statystyk re¿yserów
    excel_data = {
        'Director': [director for director, _ in sorted_director_stats],
        'Count': [count for _, count in sorted_director_stats]
    }
    df_director = pd.DataFrame(excel_data)

    # Tworzenie DataFrame dla statystyk aktorów
    excel_data = {
        'Actor': [actor for actor, _ in sorted_actor_stats],
        'Count': [count for _, count in sorted_actor_stats]
    }
    df_actor = pd.DataFrame(excel_data)

    # Pobieranie statystyk gatunków filmowych

    # Ustawianie nag³ówków dla zapytania
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}

    # Adres URL strony z rankingiem filmów
    url = 'https://www.imdb.com/chart/top/?ref_=nv_mv_250'

    # Pobieranie strony z rankingiem filmów
    response = requests.get(url, headers=headers)

    # Parsowanie strony z rankingiem filmów
    soup = BeautifulSoup(response.content, 'html.parser')

    # Znajdowanie tabeli z rankingiem filmów
    table = soup.find('table', {'class': 'chart full-width'})

    # Inicjalizacja s³ownika dla statystyk gatunków filmowych
    genre_counts = {}

    # Iteracja po wierszach tabeli
    for row in table.find_all('tr')[1:]:
        # Pobieranie kolumny z tytu³em filmu
        title_column = row.select_one('.titleColumn')

        # Pobieranie linku do filmu
        title_link = title_column.select_one('a')['href']

        # Pobieranie tytu³u filmu
        title = title_column.a.text

        # Pobieranie roku produkcji filmu
        year = row.select_one('.secondaryInfo').text.strip('()')

        # Tworzenie adresu URL filmu
        movie_url = f'https://www.imdb.com{title_link}'

        # Pobieranie strony filmu
        movie_response = requests.get(movie_url, headers=headers)

        # Parsowanie strony filmu
        movie_soup = BeautifulSoup(movie_response.content, 'html.parser')

        # Pobieranie gatunku filmu
        genre = movie_soup.find('span', {'class': 'ipc-chip__text'}).text

        # Aktualizacja statystyk dla gatunków filmowych
        if genre in genre_counts:
            genre_counts[genre] += 1
        else:
            genre_counts[genre] = 1

    # Tworzenie DataFrame dla statystyk gatunków filmowych
    df_genres = pd.DataFrame({'Genre': list(genre_counts.keys()), 'Count': list(genre_counts.values())})

    # Zapisywanie statystyk do pliku Excel
    with pd.ExcelWriter('movie_stats.xlsx') as writer:
        df_year.to_excel(writer, sheet_name='Year Stats', index=False)
        df_rating.to_excel(writer, sheet_name='Rating Stats', index=False)
        df_director.to_excel(writer, sheet_name='Director Stats', index=False)
        df_actor.to_excel(writer, sheet_name='Actor Stats', index=False)
        df_genres.to_excel(writer, sheet_name='Genres', index=False)

    # Wyœwietlanie statystyk dla gatunków filmowych
    print(df_genres)

except Exception as e:
    # Obs³uga b³êdów
    print(e)

