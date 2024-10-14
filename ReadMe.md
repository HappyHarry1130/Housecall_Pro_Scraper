# Automation Script for Housecall Pro

This project automates the process of extracting and processing data from the Housecall Pro platform using Selenium and Google Sheets API. The script logs into the platform, navigates through the calendar, extracts relevant data, and writes it to a Google Sheet.

## Features

- Automates login and navigation on Housecall Pro.
- Extracts names, links, and other relevant data from the calendar.
- Writes extracted data to a Google Sheet.
- Handles errors and retries operations when necessary.

## Prerequisites

- Python 3.x
- Google Chrome
- [ChromeDriver](https://sites.google.com/chromium.org/driver/) (managed automatically by `webdriver_manager`)
- [Google API credentials](https://developers.google.com/sheets/api/quickstart/python) for accessing Google Sheets
- Environment variables for sensitive data

## Installation

1. **Clone the repository:**

   ```bash
   git clone https://github.com/yourusername/your-repo.git
   cd your-repo
   ```

2. **Install the required packages:**

   ```bash
   pip install -r requirements.txt
   ```

3. **Set up environment variables:**

   Create a `.env` file in the root directory and add the following:

   ```plaintext
   EMAIL=your_email@example.com
   PASSWORD=your_password
   OPENAI_API_KEY=your_openai_api_key
   ```

4. **Set up Google Sheets API:**

   - Follow the [Google Sheets API Quickstart](https://developers.google.com/sheets/api/quickstart/python) to enable the API and download your credentials file.
   - Place the credentials file in the project directory and update the script to load these credentials.

## Usage

1. **Run the script:**

   ```bash
   python automation.py
   ```

2. **The script will:**

   - Log into Housecall Pro using the provided credentials.
   - Navigate through the calendar and extract data.
   - Write the extracted data to the specified Google Sheet.

## Troubleshooting

- Ensure that the ChromeDriver version matches your installed version of Google Chrome.
- Check that your `.env` file is correctly configured with valid credentials.
- Verify that your Google Sheets API credentials are correctly set up and accessible by the script.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request for any improvements or bug fixes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.