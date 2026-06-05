# Recruiter Dashboard

## Project Overview

`recruiter-dashboard` is a browser-only React dashboard for recruiters to upload candidate exports (CSV/XLSX), explore the dataset, apply multi-dimensional filters, perform keyword search, and export filtered results.

The app is designed for fast client-side processing, with no backend dependency. All parsing, filtering, rendering, and exporting happen in the browser.

## Tech Stack

- React 19
- Vite 8
- JavaScript / JSX
- PapaParse for client-side CSV parsing
- xlsx for Excel import/export
- ESLint for code quality
- Inline CSS-in-JS styling

## Core Features

- Upload `.csv`, `.xlsx`, or `.xls` candidate data
- Custom multi-select filters for Department, Role, Industry, UG/PG degree, and university
- Comma-separated skill matching with partial, case-insensitive search
- Global search across all displayed columns
- Virtualized table rendering for smooth scrolling
- Export filtered candidate results as CSV or Excel

## Architecture

- `src/main.jsx`: bootstraps the React application and renders the root component.
- `src/App.jsx`: root shell that renders `RecruiterDashboard`.
- `src/recruiter-dashboard.jsx`: single-page dashboard implementation containing upload, filter, search, export, and table logic.
- `vite.config.js`: Vite configuration with React plugin.
- `index.html`: application shell with root mount point.

## Data Flow

1. User uploads a CSV or Excel file.
2. The app parses the file using `Papa.parse` for CSV or `xlsx` for Excel.
3. Parsed rows become application state.
4. Distinct filter values are computed from uploaded rows.
5. User-selected filters, skill terms, and global search query are applied via memoized validation.
6. The filtered rows are rendered in a customized virtualized table.
7. User can export the filtered dataset back to CSV or XLSX.

## Workflow

- `npm install`: install dependencies.
- `npm run dev`: start the Vite development server.
- `npm run build`: build the production bundle.
- `npm run lint`: run ESLint across the codebase.
- `npm run preview`: preview the production build locally.

## Notes

- The project is a single-page app with no backend API.
- All operations are client-side, so data never leaves the browser during normal use.
- UI styling is handled via inline style objects directly in React components.

## Author

Rishi Rithesh
