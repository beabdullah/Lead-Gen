# Lead Finder (free)

Python (Flask) backend + HTML/CSS/JS frontend. Uses OpenStreetMap (free, no API key).
Optional OpenRouter AI for unusual keywords.

## Run
    pip install -r requirements.txt
    python app.py
Open http://127.0.0.1:5000

## Optional AI
Copy `.env.example` to `.env` and add your OPENROUTER_API_KEY (free models work).

## Google Places (recommended for full data)
1. console.cloud.google.com -> create project -> enable "Places API (New)" -> create an API key.
2. Put it in `.env` as GOOGLE_API_KEY. Google gives a free monthly usage allowance (check current limits and enable a budget alert).
Results from Google and OpenStreetMap are merged and de-duplicated.

## Google Maps scraper (no API key)
    pip install playwright
    playwright install chromium
Then tick "Use Google Maps scraper" in the app. It is slow (about 2 seconds per lead) and
scraping may go against Google's terms, so use it lightly for personal use.
