import os
from dotenv import load_dotenv
from database import get_database, db_connect
from email_connect import *
from flights import Flight, create_flight_from_email, create_flight_dicts, load_aircraft_data
from excel_manager import add_flights_to_excel
from googleapiclient.discovery import build
import logging
from datetime import datetime

load_dotenv()

# Configurar o logger
logging.basicConfig(filename='logbook_creator.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main():
    try:
        logging.info("Iniciando o programa")
        
        # Conectar à DB
        db_link = os.getenv("DB_LINK")
        db = db_connect(db_link)
        flights_collection = get_database(db)

        # Autenticar com Gmail
        creds = connect_to_gmail()
        if not creds:
            logging.error("Falha ao autenticar com Gmail")
            return

        # Construir o serviço
        service = build('gmail', 'v1', credentials=creds)

        # Buscar emails
        email_ids = fetch_emails(service, "INBOX")

        # Imprimir todos os assuntos dos emails
        print_email_subjects(service, email_ids)

        # Filtrar emails com assunto "logbook"
        logbook_email_ids = filter_logbook_emails(service, email_ids)

        all_flights = []

        # Carregar os dados das aeronaves
        aircraft_data = load_aircraft_data()

        for id in logbook_email_ids:
            email_id = id['id']
            body = get_email_body(service, email_id)
            logging.info(f"Processando email com ID: {email_id}")

            # Criar objetos Flight a partir do corpo do email
            flights = create_flight_from_email(body, aircraft_data)

            if not flights:
                logging.info(f"Não foram encontrados voos no email com ID: {email_id}")
                continue

            logging.info(f"Voos encontrados no email com ID: {email_id}")

            for flight in flights:
                logging.info(flight.to_dict())

            all_flights.extend(flights)

            # Deletar o email processado
            delete_specific_email(service, email_id)

        # Adicionar os voos na database
        flight_dicts = create_flight_dicts(all_flights)
        if flight_dicts:
            flights_collection.insert_many(flight_dicts)

        # Buscar todos os voos da DB
        all_flights_from_db = list(flights_collection.find())

        # Ordenar voos por data e hora
        all_flights_from_db.sort(key=lambda x: datetime.strptime(f"{x['date']} {x['departure_time']}", '%Y/%m/%d %H:%M'))

        # Adicionar voos ao Excel
        add_flights_to_excel("logbook.xlsx", all_flights_from_db, "_internal/ryanair_aircrafts.csv")

        logging.info("Programa concluído com sucesso")

    except Exception as e:
        logging.error("Ocorreu um erro durante a execução do programa", exc_info=True)

if __name__ == "__main__":
    main()
