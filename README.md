# Room Availability Log — web & mobile app

Staff record one entry per patient who couldn't get the room category they asked for. Every entry is written to a Google Sheet.

- `backend/Code.gs`: Google Apps Script that reads and writes the Google Sheet.
- `web/`: a single responsive app (PWA). It runs in any desktop browser, and on Android or iPhone it installs to the home screen and opens full-screen like a native app. It works offline: entries queue on the phone and sync when the connection returns.

## 1. Set up the Google Sheet backend (about 5 minutes)

1. Create a new Google Sheet, for example "Room availability log".
2. Open **Extensions → Apps Script**, delete the sample code and paste in `backend/Code.gs`. Save.
3. Choose `setup` in the function dropdown and click **Run**, then approve the permissions. This creates three tabs:
   - **Room availability log**: has the same 26 columns as the Excel file and nothing more. Each row's ID is stored as hidden row metadata, so it never shows in the sheet.
   - **Lists**: the dropdown options (sources, doctors, room categories, statuses, staff). Edit this tab to change the options in the app.
   - **Rates**: the per-day rate for each room category. "Revenue lost (₹)" is calculated from it, so **enter the real rates here.**
4. Open **Project Settings → Script properties → Add property**: `APP_KEY` = a secret phrase of your choice.
5. Go to **Deploy → New deployment → Web app**. Set Execute as **Me** and Who has access **Anyone**, then click Deploy. Copy the `/exec` URL.

## 2. Connect the app

Edit `web/config.js`:

```js
window.APP_CONFIG = {
  API_URL: 'https://script.google.com/macros/s/XXXX/exec',
  APP_KEY: 'your-secret-phrase'
};
```

If `API_URL` is empty, the app runs in **demo mode** and saves entries only in that browser.

## 3. Host the web app

Upload the `web/` folder to any static host that serves HTTPS, such as GitHub Pages, Netlify (drag and drop), Firebase Hosting or the hospital's own web server. Phones can only install the app and use offline mode over HTTPS.

**Install on a phone:** open the URL in Chrome on Android and choose ⋮ → *Install app*. On iPhone, open it in Safari and choose Share → *Add to Home Screen*.

To run it locally: `npx http-server web -p 8080`, then open http://localhost:8080.

## Calculated columns (grey in the Excel file)

| Column | Rule |
|---|---|
| Wait (days) | Earliest date we could give − Date patient wanted |
| Patient-days lost | Planned stay, if Status is *Lost*, *Went elsewhere* or *Cancelled*; otherwise 0 |
| Revenue lost (₹) | Patient-days lost × the rate for the category asked for (from the Rates tab) |
| Month | Month of the enquiry date, for example "Oct 2026" |

The lost-status list is `LOST_STATUSES` in `Code.gs`, with a matching `LOST` in `app.js`.

## Using the app

- **New entry**: The enquiry date defaults to today and "Entered by" remembers the last name used. Wait (days) updates as you type.
- **Entries**: Search by name, phone, UHID or city, and filter by status. The summary tiles show waiting, lost, average wait and patient-days lost. Tap an entry to update its status, admission date or follow-up. The update overwrites the same row in the sheet.

## Note on the existing Excel data

In the attached file, the data under "Offered instead" and "Earliest date we could give" is swapped: dates are in M and text is in N. Some cells under Admission date / Category admitted also contain text. Fix these before you import the old rows into the new sheet.
