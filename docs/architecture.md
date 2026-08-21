# Architecture Notes

## Runtime

This is a static, browser-first application. HTML pages load shared CSS and native JavaScript modules. Firebase Authentication provides sign-in, and Cloud Firestore stores projects, profiles, and issue updates.

## Module boundaries

- `js/app.js` handles login only.
- `js/dashboard.js` handles authenticated dashboard behavior and Firestore operations.
- `config/firebase-config.js` is the single source for Firebase client settings.
- `css/style.css` contains shared visual styles for both entry pages.

## Data ownership

The `osp_projects` Firestore collection is the source of truth for project and progress data. The Excel workbook in the project root is an import source, not the live database.

## Release checklist

1. Verify Firebase Authentication providers and Firestore Rules.
2. Serve the site from an HTTPS origin in production.
3. Test login, project import, progress updates, issue resolution, and admin actions.
4. Confirm that no service-account credentials are present in the published files.
