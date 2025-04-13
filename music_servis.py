# Задача с музыкальным сервисом:
# - Выгрузить все данные в xlsx в папку data
#
# - Найдите среднюю длительность треков (Milliseconds) в каждом жанре (genres).
# - Объедините таблицы tracks, albums и artists. Выведите список треков с названием альбома и именем исполнителя в xlsx
# - Определите топ-5 самых прибыльных жанров (genres) на основе суммы продаж (invoice_items.UnitPrice * invoice_items.Quantity).
# - Найдите клиентов (customers), которые купили больше всего треков в жанре "Rock".

import sqlite3
import pandas as pd
import os

# Создание папки для выгрузки данных
os.makedirs("data", exist_ok=True)

# Подключение к базе данных
conn = sqlite3.connect('./chinook.db')
tables = ['albums', 'artists', 'customers', 'employees', \
          'genres', 'invoice_items', 'invoices', 'media_types',
          'playlist_track', 'playlists', 'tracks']

# Выгрузка таблиц в xlsx
for table in tables:
    df = pd.read_sql(f'SELECT * FROM {table}', conn)
    file_path = os.path.join('data', f'{table}.xlsx')
    df.to_excel(file_path, index=False)
    print(f'Сохранено: {file_path}')
conn.close()

# Загрузка таблиц из Excel
genres = pd.read_excel('./data/genres.xlsx')
tracks = pd.read_excel('./data/tracks.xlsx')
albums = pd.read_excel('./data/albums.xlsx')
artists = pd.read_excel('./data/artists.xlsx')
invoice_items = pd.read_excel('./data/invoice_items.xlsx')
invoices = pd.read_excel('./data/invoices.xlsx')
customers = pd.read_excel('./data/customers.xlsx')

# Расчет средней длительности треков по жанрам
tracks_genres = tracks.merge(genres, on = 'GenreId').rename(columns={'Name_x': 'TrackName'})
avg_duration = tracks_genres.groupby('TrackName')['Milliseconds'].mean().reset_index()
avg_duration.rename(columns={'Name': 'GenreName', 'Milliseconds': 'AvgDuration'}, inplace=True)
avg_duration.to_excel('./data/avg_track_duration_by_genre.xlsx', index=False)


# Объединение tracks, albums, artists
tracks_albums = tracks.merge(albums, on='AlbumId')
tracks_full = tracks_albums.merge(artists, on='ArtistId')
tracks_info = tracks_full[['Name_x', 'Title', 'Name_y']]
tracks_info.columns = ['TrackName', 'AlbumTitle', 'ArtistName']
tracks_info.to_excel('./data/tracks_albums_artists.xlsx', index=False)


# Определение топ-5 самых прибыльных жанров (genres)
sales = invoice_items.merge(tracks, on='TrackId', )
sales = sales.merge(genres, on='GenreId')
sales['Revenue'] = sales['UnitPrice_x'] * sales['Quantity']
genre_revenue = sales.groupby('Name_y')['Revenue'].sum().reset_index()
top5_genres = genre_revenue.sort_values(by='Revenue', ascending=False).head(5)
top5_genres.rename(columns={'Name': 'Genre'}, inplace=True)
top5_genres.to_excel('./data/top_5_genres_by_revenue.xlsx', index=False)


# Топ покупателей в жанре "Rock"
full_data = (invoice_items.merge(tracks, on=['TrackId','UnitPrice']).rename(columns={'Name': 'TrackName'})\
             .merge(genres, on='GenreId').rename(columns={'Name': 'GenreName'})\
             .merge(invoices, on='InvoiceId').merge(customers, on='CustomerId'))

rock_data = full_data[full_data['TrackName'] == 'Rock']
rock_top = rock_data.groupby(['CustomerId', 'FirstName', 'LastName'])['InvoiceLineId'].count().reset_index()
rock_top = rock_top.sort_values(by='InvoiceLineId', ascending=False)
rock_top.rename(columns={'InvoiceLineId': 'RockTracksPurchased'}, inplace=True)
rock_top.to_excel('./data/top_rock_customers.xlsx', index=False)
