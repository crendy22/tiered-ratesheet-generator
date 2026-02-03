# Tiered Ratesheet Generator

A web-based tool for generating multiple tiered ratesheets from a base ratesheet with configurable pricing deltas.

## Overview

This tool allows mortgage lenders and investors to:
1. Upload a base ratesheet (Excel file with multiple product tabs)
2. Configure pricing adjustments (deltas) for each tier and product
3. Generate and download adjusted ratesheet files

## Features

- **Drag & Drop Upload**: Easy file upload for base ratesheets
- **Auto-Detection**: Automatically detects pricing grids in each sheet
- **Multiple Tiers**: Create unlimited pricing tiers
- **Per-Tab Deltas**: Set different adjustments for each product tab
- **Import Deltas**: Bulk import deltas from an Excel file
- **Instant Generation**: Generate all tier files with one click

## Usage

### Manual Delta Entry

1. Upload your base ratesheet (.xlsx file)
2. Click "Add Tier" to create a new pricing tier
3. Enter delta values for each product tab (positive = better pricing, negative = worse)
4. Click "Generate Ratesheets" to download

### Importing Deltas

You can import deltas from an Excel file with this format:

| Tier Name | Tab Name | Delta |
|-----------|----------|-------|
| Tier 2    | DSCR1    | -0.25 |
| Tier 2    | DSCR2    | 0.75  |
| Tier 3    | DSCR1    | -0.50 |
| Tier 3    | DSCR2    | 0.50  |

## Deployment

### GitHub Pages

1. Push to GitHub
2. Go to Settings > Pages
3. Select "Deploy from a branch" > main > / (root)
4. Your app will be live at `https://yourusername.github.io/repo-name`

### Vercel / Netlify

Simply connect your GitHub repo - no build configuration needed.

## Local Development

Just open `index.html` in a browser. No build step required.

```bash
# Or use a local server
npx serve .
# or
python -m http.server 8000
```

## Technical Notes

- Pure HTML/CSS/JavaScript - no framework dependencies
- Uses [SheetJS](https://sheetjs.com/) for Excel file processing
- Pricing grid detection looks for "Rate" header followed by lock period columns
- Deltas are applied to all numeric cells within detected pricing grids

## File Structure

```
├── index.html      # Main HTML file
├── styles.css      # Styling
├── app.js          # Application logic
└── README.md       # This file
```

## License

MIT
