# Room Availability Log — web & mobile app

Staff record one entry per patient who couldn't get the room category they asked for. Every entry is written to a Google Sheet.

- `backend/Code.gs`: Google Apps Script that reads and writes the Google Sheet.
- `web/`: a single responsive app (PWA). It runs in any desktop browser, and on Android or iPhone it installs to the home screen and opens full-screen like a native app. It works offline: entries queue on the phone and sync when the connection returns.

## 1. Set up the Google Sheet backend (about 5 minutes)

1. Create a new Google Sheet, for example "Room availability log".
2. Open **Extensions → Apps Script**, delete the sample code and paste in `backend/Code.gs`. Save.
3. Choose `setup` in the function dropdown and click **Run**, then approve the permissions. This creates three tabs:
   - **Room availability log**: the app finds each column by its header name in row 1, so you can add, move or remove columns. Each row's ID is stored as hidden row metadata, so it never shows in the sheet.
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
| Patient-days lost | Lost / Went elsewhere / Cancelled: the planned stay. Otherwise: days between Date patient wanted and Admission date, if admitted later |
| Revenue lost (₹) | **Lost / Went elsewhere / Cancelled:** planned stay × rate of the category asked for.<br>**Admitted later than wanted:** delay days × rate of the category asked for.<br>**Booked / Admitted in a cheaper room:** (rate asked − rate given) × planned stay. The room given is "Admission category", or "Offerred Room" if that is blank.<br>The delay and cheaper-room amounts add up when both apply. |
| Month | Month of the enquiry date, for example "Oct 2026" |

The status lists are `LOST_STATUSES` and `BOOKED_STATUSES` in `Code.gs`, with matching `LOST` and `BOOKED` in `app.js`. Phone numbers are saved with the country code, for example `+91 9876543210`. Every field in the first two sections is mandatory. UHID is mandatory only for returning patients.

## Using the app

- **New entry**: The enquiry date defaults to today and "Entered by" remembers the last name used. Wait (days) updates as you type.
- **Status** is fixed to Waiting, Booked, Admitted and Lost (the Lists tab can't change it). The Outcome section shows only the fields for the chosen status: Waiting shows Last follow-up and Remarks. Booked and Admitted add Admission date and Admission category, both mandatory. Lost adds Reason not admitted (mandatory) and Where they went. Fields hidden by the status are cleared when you save.
- **Entries**: Search by name, phone, UHID or city, and filter by status. The summary tiles show waiting, lost, average wait and patient-days lost. Tap an entry to update its status, admission date or follow-up. The update overwrites the same row in the sheet.

## Note on the existing Excel data

In the attached file, the data under "Offered instead" and "Earliest date we could give" is swapped: dates are in M and text is in N. Some cells under Admission date / Category admitted also contain text. Fix these before you import the old rows into the new sheet.
