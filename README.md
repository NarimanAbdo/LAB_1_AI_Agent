# Lab 1: Working with a Public REST API Using Python

This program calls NASA's **Near-Earth Object Web Service (NeoWs)** REST API, parses the JSON response, stores each asteroid in an `Asteroid` class, prints a formatted summary of the asteroids passing close to Earth over the next 3 days, and exports the data to an Excel file (`asteroids.xlsx`).

## Files

| File | Purpose |
|---|---|
| `main.py` | Source code: API call, JSON parsing, `Asteroid` class, formatted output, Excel export |
| `.env` | Holds the API key (`NASA_API_KEY`), **not to be shared or uploaded** |
| `requirements.txt` | Dependencies (`requests`, `python-dotenv`, `openpyxl`) |
| `.gitignore` | Keeps `.env` and the virtual environment out of git |
| `asteroids.xlsx` | Created when the program runs: an `Asteroids` sheet and a `Summary` sheet |

## Setup and run (Windows)

```powershell
cd F:\ai_agent_iti\Lab_Python_Solution
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

On macOS/Linux, activate the environment with `source .venv/bin/activate`.

## API key

Put your key in `.env` as `NASA_API_KEY=...`. `DEMO_KEY` works for testing but allows only a few requests per hour; a free personal key is available at https://api.nasa.gov.
