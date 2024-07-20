
## Functions Overview

### `main.py`

- Orchestrates the whole process: fetching emails, parsing flight data, storing in DB, and generating the Excel logbook.

### `database.py`

- Connects to MongoDB and manages database interactions.

### `email_connect.py`

- Handles Gmail API authentication and email fetching.

### `flights.py`

- Parses flight information from email bodies and converts them into `Flight` objects.

### `excel_manager.py`

- Manages reading, writing, and updating the Excel logbook.

### `airport_update.py`

- Downloads and processes airport data to keep the IATA to ICAO mappings up to date.

## Contributing

1. Fork the repository.
2. Create your feature branch (`git checkout -b feature/YourFeature`).
3. Commit your changes (`git commit -am 'Add some feature'`).
4. Push to the branch (`git push origin feature/YourFeature`).
5. Create a new Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Google API](https://developers.google.com/gmail/api)
- [MongoDB](https://www.mongodb.com/)
- [cx_Freeze](https://cx-freeze.readthedocs.io/en/latest/)
- [OpenFlights](https://openflights.org/data.html)

---
