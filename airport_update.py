import csv
import requests

def download_airports_data(url):
    response = requests.get(url)
    response.raise_for_status()  # Ensure we notice bad responses
    return response.text

def parse_airports_data(data):
    airports = {}
    reader = csv.reader(data.splitlines())
    for row in reader:
        airport_id, name, city, country, iata, icao, latitude, longitude, altitude, timezone, dst, tz, type, source = row
        if iata and icao:
            airports[iata] = {'icao': icao, 'latitude': float(latitude), 'longitude': float(longitude)}
    return airports

# URL do arquivo CSV contendo dados dos aeroportos
url = 'https://raw.githubusercontent.com/jpatokal/openflights/master/data/airports.dat'

# Fazer o download e parse dos dados dos aeroportos
data = download_airports_data(url)
iata_to_icao_coords = parse_airports_data(data)

# Salvar o dicionário em um arquivo CSV local para uso posterior
with open('iata_to_icao_coords.csv', 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['IATA', 'ICAO', 'Latitude', 'Longitude'])  # Cabeçalho
    for iata, info in iata_to_icao_coords.items():
        writer.writerow([iata, info['icao'], info['latitude'], info['longitude']])

print("Arquivo iata_to_icao_coords.csv criado com sucesso.")
