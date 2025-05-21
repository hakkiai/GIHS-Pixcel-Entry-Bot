# GIHS Pixcel Entry Bot

An automated form-filling bot for the GIHS Pixcel Entry system using Selenium and Excel data.

## Project Structure

```
SELENIUM-AUTOMATE-FORM/
├── data/                # Excel file(s)
├── main.py               
└── README.md            # Project instructions
```

## Prerequisites

- Python 3.8 or higher
- Chrome browser installed
- ChromeDriver matching your Chrome version

## Setup

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd pixcel-entry-bot
   ```

2. Create a virtual environment and activate it:
   ```bash
   python -m venv venv
   .\venv\Scripts\activate  # Windows
   source venv/bin/activate  # Linux/Mac
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Place your Excel file in the `data/` directory

## Usage

1. Ensure your Excel file is in the correct format with the following columns:
   - FormNumber
   - Name
   - Birthday
   - MothersMaiden
   - Address
   - Zipcode
   - Country
   - Occupation
   - Company

2. Run the script:
   ```bash
   python src/main.py
   ```

## Security Notes

- Never commit the `.env` file to version control
- Keep your credentials secure
- Regularly update ChromeDriver to match your Chrome version

## Troubleshooting

If you encounter any issues:
1. Verify your ChromeDriver version matches your Chrome browser
2. Check that all environment variables are set correctly in `.env`
3. Ensure your Excel file is in the correct format and location

