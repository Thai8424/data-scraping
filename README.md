# Data Scraping (TypeScript)

Two independent TypeScript scrapers that export company directory listings to Excel:

- `hsctvn-scraper/` collects company listing and detail fields from hsctvn.com.
- `trangvang-scraper/` collects company directory fields from trangvangvietnam.com.

## Requirements

- Node.js (use a current LTS release)
- Install dependencies in the scraper you plan to run: `npm install`

## Run

From the selected scraper folder:

```bash
npm run dev
```

The scripts write `.xlsx` files in their working directory. Generated spreadsheets and other collected data are excluded from Git by the root `.gitignore`.

## Notes

These scripts make requests to third-party websites. Review each site's terms and robots guidance, keep request rates reasonable, and collect only information you are permitted to use. The repository intentionally excludes scraped datasets.
