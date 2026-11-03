# HQ Territory Mapper

A Google Apps Script web app that turns a Google Sheet of territory boundary points into an interactive map. Each area's boundary is drawn on OpenStreetMap using Leaflet, and the **latest entry for an area is treated as its current location**, so the site always shows the most recent update without anyone editing the map by hand.

## Features

- **Interactive map** built on OpenStreetMap + Leaflet, with each area drawn from its boundary points
- **Editable boundaries**: adjust points directly on the map or in the web app and save back to the sheet
- **Latest-wins tracking**: when an area has multiple entries, the newest one is shown as the current location
- **Distance calculations**: point-to-point and diagonal distances in km for every area
- **Dashboard, HQs, Areas, Distances and Import tabs** in a single branded web app
- **Sheet-backed**: no separate database; Google Sheets is the source of truth

## How the data is laid out

Each area is a block in the sheet:

| Field | Description |
|---|---|
| Area name | e.g. Ijebu, Kaduna, Abeokuta |
| `New Map` | Marker that starts a new boundary entry |
| Point 1 to 4 | Latitude / longitude pairs defining the boundary |
| Distances | Point-to-point and diagonal distances in km, beside the points |

Region and AHQ tables hold the same Point 1 to 4 coordinates and distances, grouped by region.

If an area appears more than once, the **most recent entry** is used as its current location.

## Project structure

```
.
├── Code.gs          # Apps Script backend: reads/writes the sheet, serves the app
├── WebApp.html      # Frontend: map, tabs, editing UI
└── README.md
```

> Adjust the file names above to match your repo.

## Setup

1. Create or open the Google Sheet that holds your boundary points.
2. In the sheet, open **Extensions → Apps Script**.
3. Add the backend code (`Code.gs`) and create an HTML file named `WebApp` with the frontend code.
4. Save, then choose **Deploy → New deployment → Web app**.
5. Set **Execute as** to *Me* and choose who has access.
6. Open the deployment URL to use the app.

After changing the code, create a **new deployment version** for the changes to go live.

## Usage

| Tab | What it does |
|---|---|
| **Dashboard** | Overview of areas and their latest status |
| **HQs** | HQ locations and their coordinates |
| **Areas** | Per-area boundaries on the map; edit points here |
| **Distances** | Point-to-point and diagonal distances in km |
| **Import** | Bring in new boundary data to the sheet |

To update an area, edit its points on the map or in the Areas tab. The new entry becomes that area's current location.

## Tech stack

- Google Apps Script
- Google Sheets
- [Leaflet](https://leafletjs.com/)
- [OpenStreetMap](https://www.openstreetmap.org/) tiles
- HTML / CSS / JavaScript

## Notes

- Map tiles come from OpenStreetMap, so follow their [tile usage policy](https://operations.osmfoundation.org/policies/tiles/) for production use.
- Boundary data may be sensitive. Keep the repo private unless you are sure the data can be shared.
- Never commit sheet IDs, deployment URLs or credentials you don't want public.

## Author

Ajagbe Adedamola Daniel
