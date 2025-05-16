# HT Consumer Data Scraper

This Python script scrapes consumer data from the Tamil Nadu Electricity Board website for consumer numbers ranging from 0001 to 9999. The data is saved into an Excel file named `consumer_data.xlsx`.

---

## Features

- Selects district and section on the website automatically.
- Inputs consumer number and fetches consumer details.
- Handles pagination and errors gracefully.
- Saves collected data periodically to avoid file permission issues.

---

## Requirements

- Python 3.7 or higher
- [Playwright](https://playwright.dev/python/)
- openpyxl

---

## Installation

1. Clone this repository or download the script files.

2. Open your terminal or command prompt, then install the required Python packages:

```bash
pip install playwright openpyxl
````

3. Install Playwright browser dependencies:

```bash
python -m playwright install
```

---

## Usage

Run the scraper script with:

```bash
python scraper.py
```

* The script will start fetching data from consumer number `0001` to `9999`.
* Progress and errors will be printed to the console.
* Data will be saved in `consumer_data.xlsx` in the same folder.

---

## Notes

* Ensure `consumer_data.xlsx` is **closed** before running the script to avoid file permission errors.
* If you want to change the range or other parameters, edit the `scraper.py` file accordingly.
* The script uses Playwright for browser automation, so it may take some time depending on network speed.

---

## Troubleshooting

* If you get a `PermissionError` for the Excel file, make sure it is not open in any other program.
* If you encounter missing browser errors, run `python -m playwright install` again.
* For any other errors, check your Python and package versions.

---

## License

This project is released under the MIT License.


