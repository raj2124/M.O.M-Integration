# M.O.M Integration

Minutes of Meeting web app with Zoho Projects integration, PDF export, print support, and Outlook draft handoff.

## Features

- Dashboard with recent Zoho projects and project status badges
- M.O.M editor aligned to your sheet format
- Zoho project selection from dropdown while creating a new M.O.M
- Auto-fill of project fields from Zoho (project name/number and team users where available)
- Agenda and attendees row management
- PDF generation and browser print flow
- Outlook compose link with prefilled M.O.M details (user sends manually)

## Stack

- Node.js + Express
- Vanilla HTML/CSS/JS
- PDFKit

## Project Structure

- `src/server.js` - API and app server
- `src/zohoClient.js` - Zoho Projects integration
- `src/pdfService.js` - PDF generation
- `public/index.html` - App UI
- `public/app.js` - Frontend logic
- `public/styles.css` - Styling
- `scripts/smoke-test.js` - smoke checks

## Local Setup

1. Install dependencies

```bash
npm install
```

2. Create environment file

```bash
cp .env.example .env
```

3. Configure Zoho (live mode)

- Set `ZOHO_USE_MOCK=false`
- Set `ZOHO_PORTAL_ID`
- Set `ZOHO_CLIENT_ID`
- Set `ZOHO_CLIENT_SECRET`
- Set `ZOHO_REFRESH_TOKEN`
- Set `ZOHO_BASE_URL` and `ZOHO_ACCOUNTS_BASE_URL` for your Zoho data center (`.com` or `.in`)
- Optional: set `ZOHO_PROJECTS_ENDPOINT` if your org uses a custom endpoint

4. Run app

```bash
npm run dev
```

5. Open `http://localhost:3000`

## Deployment Notes

- Keep all secrets in host environment variables (Render/GitHub/other host), never in Git.
- Generated PDFs are written to `generated-pdfs/`.
- Browser security does not allow automatic file attachment to Outlook compose; users attach generated PDF manually.

## License

MIT (see `LICENSE`).
