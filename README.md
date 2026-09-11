# TPN Dedicated Day — Render / GitHub package

This package follows the hosted approach used by the Undelivered Export:
GitHub stores the code and Render runs the browser automation and serves the dashboard.

## What it does

### 08:00 — full refresh
- Logs into TPN from Render using Render environment variables.
- Integration: exports the last 7 working days.
- Browse: exports the last 7 working days.
- Does NOT use TPN grid filters.
- Filters locally:
  - Service starts with `DD`
  - Browse Requesting Depot / Req = `8`
  - Browse Delivery Depot / Del != `8`
- Uses the latest Delivery Date found in the filtered Browse export for the dashboard population.
- Matches Integration Status to Browse rows by Docket.
- Writes `dedicated-day-data.json`.

### 10:00 / 12:00 / 14:00 / 16:00 / 18:00 — status refresh
- Logs into TPN.
- Re-exports Browse only.
- Applies the same filters locally.
- Matches by Docket.
- Updates only Status values in `dedicated-day-data.json`.

### Dashboard
- Served directly by Render.
- Automatically checks `dedicated-day-data.json` every 5 minutes.
- Existing search, headline filtering and acknowledgements remain in the dashboard.

## GitHub

Create a new repository, or a separate branch/folder in your existing TPN repository, and add every file from this package.

Do **not** commit real TPN credentials.

## Render

Easiest route: use the included `render.yaml` as a Render Blueprint.

Set these secret environment variables in Render:
- `TPN_USERNAME`
- `TPN_PASSWORD`

Optional:
- `TPN_HOLIDAYS` as comma-separated ISO dates, e.g. `2026-12-25,2026-12-28`.

The Blueprint also creates a `REFRESH_TOKEN`.

## URLs after deployment

- `/` dashboard
- `/health` automation status
- `/admin` simple status page

Manual refresh API:
- POST `/refresh/full`
- POST `/refresh/status`
with HTTP header `X-Refresh-Token: <REFRESH_TOKEN>`

## Important first-live-run note

The Browse selectors are based on the already proven Undelivered Export automation pattern.
The Integration screen has more variation between TPN environments, so its first live run may
need a small selector adjustment if your Integration page labels differ.

If a TPN step fails the service saves:
- `tpn_failure.png`
- `tpn_failure.html`

Check the Render logs and `/health` to see the failed stage.

## Render filesystem

Render's normal service filesystem is ephemeral. This package mitigates that by doing a full
refresh automatically on service startup when today's data is not present. A persistent Render
disk can still be added later if you want the JSON to survive every restart without recollection.

## Security

Credentials are read only from Render environment variables. They are not written into
`dashboard.html`, `dedicated-day-data.json`, logs, or the GitHub repository.
