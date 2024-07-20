import os
import re
import csv
from datetime import datetime
from suntime import Sun
import sys

# Carregar os dados do CSV para um dicionário
def load_iata_to_icao_coords():
    iata_to_icao_coords = {}
    base_path = os.path.abspath(".")
    if getattr(sys, 'frozen', False):
        # Running in a bundle
        base_path = os.path.dirname(sys.executable)
    csv_file_path = os.path.join(base_path, '_internal/iata_to_icao_coords.csv')

    with open(csv_file_path, mode='r') as infile:
        reader = csv.reader(infile)
        next(reader)  # Skip header
        for rows in reader:
            iata_to_icao_coords[rows[0]] = {'icao': rows[1], 'latitude': float(rows[2]), 'longitude': float(rows[3])}
    return iata_to_icao_coords

# Carregar o dicionário IATA para ICAO e coordenadas
IATA_TO_ICAO_COORDS = load_iata_to_icao_coords()

class Flight:
    def __init__(self, date, departure_airport, arrival_airport, departure_time, arrival_time, aircraft_registration, aircraft_type, flight_time, captain, takeoffs_day, takeoffs_night, landings_day, landings_night, night_flight_time, ifr_time):
        self.date = date
        self.departure_airport = departure_airport
        self.arrival_airport = arrival_airport
        self.departure_time = departure_time
        self.arrival_time = arrival_time
        self.aircraft_registration = aircraft_registration
        self.aircraft_type = aircraft_type
        self.flight_time = flight_time
        self.captain = captain
        self.takeoffs_day = takeoffs_day
        self.takeoffs_night = takeoffs_night
        self.landings_day = landings_day
        self.landings_night = landings_night
        self.night_flight_time = night_flight_time
        self.ifr_time = ifr_time
        self.datetime = datetime.strptime(f"{self.date} {self.departure_time}", '%Y/%m/%d %H:%M')
    
    def to_dict(self):
        return {
            'date': self.date,
            'departure_airport': self.departure_airport,
            'arrival_airport': self.arrival_airport,
            'departure_time': self.departure_time,
            'arrival_time': self.arrival_time,
            'aircraft_registration': self.aircraft_registration,
            'aircraft_type': self.aircraft_type,
            'flight_time': self.flight_time,
            'captain': self.captain,
            'takeoffs_day': self.takeoffs_day,
            'takeoffs_night': self.takeoffs_night,
            'landings_day': self.landings_day,
            'landings_night': self.landings_night,
            'night_flight_time': self.night_flight_time,
            'ifr_time': self.ifr_time,
            'datetime': self.datetime
        }

def convert_iata_to_icao(iata_code):
    return IATA_TO_ICAO_COORDS.get(iata_code, {'icao': iata_code})['icao']

def get_airport_coordinates(iata_code):
    return IATA_TO_ICAO_COORDS.get(iata_code, {'latitude': 0.0, 'longitude': 0.0})

def is_night_time(time, sunrise, sunset):
    return time < sunrise or time > sunset

def extract_captain_name(email_body):
    match = re.search(r'Flight Deck Crew\s+\w+\s+:\s+CP\s+:\s+([A-Z\s]+)', email_body)
    if match:
        captain_name = match.group(1).strip().split('\n')[0]  # Split at new line and take the first part
        return captain_name
    return "N/A"

def load_aircraft_data():
    """Carregar dados das aeronaves do ficheiro CSV."""
    aircraft_data = {}
    base_path = os.path.abspath(".")
    if getattr(sys, 'frozen', False):
        # Running in a bundle
        base_path = os.path.dirname(sys.executable)
    csv_file_path = os.path.join(base_path, '_internal/ryanair_aircrafts.csv')

    with open(csv_file_path, mode='r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            registration = row['Registration'].replace("-", "").upper()
            aircraft_data[registration] = {
                'model': row['Aircraft Type'],
                'variant': 'B737-800' if '737-800' in row['Aircraft Type'] else 'B737-8200' if 'MAX 8' in row['Aircraft Type'] else 'NA'
            }
    return aircraft_data

def create_flight_from_email(email_body, aircraft_data):
    flights = []
    date_match = re.search(r'(\d{4}/\d{2}/\d{2})', email_body)
    date = date_match.group(1) if date_match else None
    if not date:
        print("Data não encontrada no corpo do email.")
        return flights
    
    flight_sections = re.split(r'(FlightNumber\s+:\s+\w+)', email_body)[1:]  # Split and remove first empty element
    flight_sections = [''.join(flight_sections[i:i+2]) for i in range(0, len(flight_sections), 2)]  # Join headers with content

    captain_name = extract_captain_name(email_body)

    for section in flight_sections:
        # Debug only.
        # print(f"Processando seção de voo: {section[:100]}...")
        
        city_pair_match = re.search(r'City Pair\s+:\s+(\w+) - (\w+)', section)
        departure_time_match = re.search(r'Airborne\s+:\s+(\d{2}:\d{2})', section)
        arrival_time_match = re.search(r'Landed\s+:\s+(\d{2}:\d{2})', section)
        aircraft_registration_match = re.search(r'Registration\s+:\s+(\w+)', section)
        flight_time_match = re.search(r'Total flight\s+:\s+(\d{2}:\d{2})', section)
        takeoffs_day_match = re.search(r'Take Offs Day\s+:\s+(\d+)', section)
        takeoffs_night_match = re.search(r'Take Offs Night\s+:\s+(\d+)', section)
        landings_day_match = re.search(r'Landings Day\s+:\s+(\d+)', section)
        landings_night_match = re.search(r'Landings Night\s+:\s+(\d+)', section)
        night_flight_time_match = re.search(r'Night Flight Time\s+:\s+(\d{2}:\d{2})', section)
        ifr_time_match = re.search(r'IFR Time\s+:\s+(\d{2}:\d{2})', section)

        if not city_pair_match or not departure_time_match or not arrival_time_match:
            # Debug only:
            # print(f"Pulando secção devido a dados ausentes: {section[:100]}...")
            continue

        # Handle missing data
        departure_airport_iata = city_pair_match.group(1)
        arrival_airport_iata = city_pair_match.group(2)
        departure_airport = convert_iata_to_icao(departure_airport_iata)
        arrival_airport = convert_iata_to_icao(arrival_airport_iata)
        departure_time = departure_time_match.group(1)
        arrival_time = arrival_time_match.group(1)
        aircraft_registration = aircraft_registration_match.group(1) if aircraft_registration_match else "N/A"
        flight_time = flight_time_match.group(1) if flight_time_match else "00:00"
        takeoffs_day = int(takeoffs_day_match.group(1)) if takeoffs_day_match else 0
        takeoffs_night = int(takeoffs_night_match.group(1)) if takeoffs_night_match else 0
        landings_day = int(landings_day_match.group(1)) if landings_day_match else 0
        landings_night = int(landings_night_match.group(1)) if landings_night_match else 0
        night_flight_time = night_flight_time_match.group(1) if night_flight_time_match else "00:00"

        # Calculate IFR time as the difference between arrival and departure times
        departure_time_obj = datetime.strptime(departure_time, "%H:%M")
        arrival_time_obj = datetime.strptime(arrival_time, "%H:%M")
        ifr_time_delta = arrival_time_obj - departure_time_obj
        ifr_time = str(ifr_time_delta)[:-3]

        # Calculate sunrise and sunset times
        coordinates = get_airport_coordinates(departure_airport_iata)
        sun = Sun(coordinates['latitude'], coordinates['longitude'])
        date_obj = datetime.strptime(date, "%Y/%m/%d")
        sunrise = sun.get_sunrise_time(date_obj).time()
        sunset = sun.get_sunset_time(date_obj).time()

        # Determine if the takeoff and landing were during day or night
        takeoffs_day = 1 if not is_night_time(departure_time_obj.time(), sunrise, sunset) else 0
        takeoffs_night = 1 if is_night_time(departure_time_obj.time(), sunrise, sunset) else 0
        landings_day = 1 if not is_night_time(arrival_time_obj.time(), sunrise, sunset) else 0
        landings_night = 1 if is_night_time(arrival_time_obj.time(), sunrise, sunset) else 0

        # Obter o tipo de aeronave a partir dos dados de aeronaves
        registration_no_dash = aircraft_registration.replace("-", "").upper()
        aircraft_type = aircraft_data.get(registration_no_dash, {'model': 'NA'})['model']

        flight = Flight(
            date=date,
            departure_airport=departure_airport,
            arrival_airport=arrival_airport,
            departure_time=departure_time,
            arrival_time=arrival_time,
            aircraft_registration=aircraft_registration,
            aircraft_type=aircraft_type,
            flight_time=flight_time,
            captain=captain_name,
            takeoffs_day=takeoffs_day,
            takeoffs_night=takeoffs_night,
            landings_day=landings_day,
            landings_night=landings_night,
            night_flight_time=night_flight_time,
            ifr_time=ifr_time
        )
        
        # Debug only
        # print(f"Voo criado: {flight.to_dict()}")
        
        flights.append(flight)

    return flights

def create_flight_dicts(flights):
    return [flight.to_dict() for flight in flights]
