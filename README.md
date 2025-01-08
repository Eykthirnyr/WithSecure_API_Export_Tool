# WithSecure API Export Tool

## Overview

The **WithSecure API Export Tool** is a user-friendly application that allows you to retrieve and export device data and current OS details from the WithSecure API into a CSV file. This tool simplifies the process of managing and analyzing OS versions across organizations.

## Features

- **Authentication**: Securely authenticate with WithSecure using your Client ID and Client Secret.
- **Data Retrieval**: Fetch data about organizations and devices linked to your WithSecure account.
- **Export to CSV**: Save the data in a UTF-8 encoded CSV file, ensuring compatibility with tools like Google Sheets and Excel.
- **User-Friendly GUI**: Easy-to-use graphical interface built with Python's Tkinter.
- **Progress Tracking**: Real-time status updates and progress bar during data retrieval.
- **Retrieve current OS type (e.g., Windows, Ubuntu, Android, etc.) and current version.
- **Retrieve the names of all computers (NETBIOS names) for every organization you have access to in https://elements.withsecure.com .

## Requirements

- Python 3.7+
- Internet connection


## Usage

1. Launch the application.
2. Enter your WithSecure **Client ID** and **Client Secret**.
3. Choose an export folder for the CSV file.
4. Acknowledge the security warning about providing **READ ONLY** API access.
5. Click "Export to CSV" to start retrieving data.
6. Open the generated CSV file in the export folder.

## Notes

- **UTF-8 Encoding**: The exported CSV file includes a BOM to ensure proper display of special characters (e.g., accents) in Excel.
- **API Key Generation**: Create your API keys from the [WithSecure Elements API Key Page](https://elements.withsecure.com/apps/ccr/api_keys).

## Screenshots

![WSAPIET](https://github.com/user-attachments/assets/66a4ff0d-c74b-49fa-ac0d-c90f30f2c323)

![Export_CSV](https://github.com/user-attachments/assets/565d495b-499c-4779-9523-167d2af4408f)

## Acknowledgments

Made by [Clément GHANEME](https://clement.business/) 08/01/2025.


