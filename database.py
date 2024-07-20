import pymongo
import logging

def db_connect(db_link):

    if not db_link:
        raise ValueError("DB_LINK environment variable not found")

    logging.info(f"DB_LINK: {db_link}")

    try:
        # Connection to the DB
        myclient = pymongo.MongoClient(db_link)
        logging.info("Conexão com MongoDB estabelecida com sucesso.")

        return myclient
    except Exception as e:
        logging.error(f"Erro ao conectar ao MongoDB: {e}")
        raise

def get_database(myclient):

    try:
        # Select the database
        mydb = myclient["logbook"]

        # Select the collection
        mycol = mydb["flights"]

        logging.info("Database e coleção selecionados com sucesso.")
        return mycol
    except Exception as e:
        logging.error(f"Erro ao selecionar a base de dados ou coleção: {e}")
        raise
