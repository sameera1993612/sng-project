# SNG Drawing Management System

A browser-based workspace for managing SNG projects, invoice drawings, as-built drawings, progress updates, and issue logs.

## Project structure

- `index.html` - authenticated entry page
- `dashboard.html` - main application workspace
- `css/` - shared application styles
- `js/` - page behavior and Firebase data operations
- `config/` - shared client configuration
- `assets/brand/` - application branding and icons
- `data/` - future static data sources and import notes
- `docs/` - architecture and operational notes

## Run locally

Because the app uses ES modules, serve the project through a local web server instead of opening the HTML files directly. For example, use VS Code Live Server or any static HTTP server, then open `index.html`.

Users must be created in Firebase Authentication. Firestore access is controlled by the project's security rules.

## Security notes

The Firebase web API key is an identifier, not a password. Keep authorization in Firebase Authentication and Firestore Rules. Never commit service-account keys or private credentials.
