# Laura Husseiny Tutoring

Personal tutoring website: about me, subjects, a live booking calendar (available, requested, booked) with weekly repeating sessions, resources and revision sites.

The whole site is one file, `index.html`. Everything you would normally change (slots, fixed sessions, counters, resources, links) sits in the **CONFIG** block near the bottom of that file, marked `EDIT THIS SECTION`.

## Put it online with GitHub Pages (about 10 minutes)

1. Sign in at github.com (or create a free account).
2. Click **+** (top right) → **New repository**. Name it `laura-tutoring`, set it to **Public**, and click **Create repository**.
3. On the new page, click **uploading an existing file**. Drag in `index.html`, `README.md`, `booking-sheet.gs` and `.nojekyll`, then click **Commit changes**.
   (`.nojekyll` is hidden on Mac. In Finder press Cmd + Shift + . to show it. If it won't upload, the site still works without it.)
4. Go to **Settings** → **Pages**. Under *Build and deployment*, choose **Deploy from a branch**, branch **main**, folder **/ (root)**, and click **Save**.
5. After a minute or two the site is live at `https://YOUR-USERNAME.github.io/laura-tutoring/`.

To use your own domain later (e.g. `laurahusseiny.com`), add it under **Settings → Pages → Custom domain**.

## Requested slots and booking emails (Google Sheet)

The calendar shows three states: **Available** (green), **Requested** (amber, dashed) and **Booked** (grey).
When a student sends a request, those slots turn amber for every visitor, you get an email, and the request appears in a Google Sheet.
To confirm, change the row's Status to **Approved**: the student gets a confirmation email from you and the slot shows as Booked. Choose **Declined** and the student gets a polite email and the slot becomes available again.
Tip for repeating bookings: select all of that student's Status cells, type Approved, and press Ctrl+Enter (Cmd+Enter on Mac) so they get one email listing every date.

One-time setup, about 10 minutes:

1. Go to sheets.google.com (signed in as husseinylaura@gmail.com) and create a blank sheet. Name it "Tutoring bookings".
2. Click **Extensions → Apps Script**. Delete what's there, paste the whole of `booking-sheet.gs`, and click **Save**.
3. Click **Deploy → New deployment**. Click the gear icon, choose **Web app**, and set:
   - *Execute as:* **Me**
   - *Who has access:* **Anyone**
4. Click **Deploy**. Google asks you to authorise it: choose your account → **Advanced** → **Go to (project)** → **Allow**. This lets the script write to your sheet and email you.
5. Copy the **Web app URL** (it ends in `/exec`).
   Then, back in the Apps Script editor, pick **setup** in the function menu at the top and click **Run** once. This switches on the approval emails.
6. In `index.html`, find `sheetEndpoint: ""` and paste the URL between the quotes. Upload the file to GitHub.

A "Requests" tab appears in the sheet with the first request. Visitors only ever see dates, times and status, never names or emails.

If you edit the script later, use **Deploy → Manage deployments → Edit → New version** so the same URL keeps working.

Without this setup, "Requested" only shows on the student's own device, and the request goes to you through their email app or WhatsApp.

(Formspree is an alternative for email only: put its form URL in `formEndpoint`. It can't mark slots as requested on the calendar.)

## Everyday edits

On GitHub, open `index.html` → click the pencil icon → change the CONFIG values → **Commit changes**. The live site updates within a minute.

- **New fixed weekly session:** add it under `weeklyBooked` (1 = Monday … 6 = Saturday, 0 = Sunday).
- **Changing a weekly session for one week only:** add the date to `weeklyBookedSkip`, then add that day's new time to `booked`.
- **Confirming requests:** change Status to Approved in the Google Sheet. No website edit needed.
- **One-off booking you arranged yourself:** add a row to the sheet with Status Approved, or a line to `booked` in `index.html`.
- **Counters:** change the numbers in `stats`.
- **Resources:** paste each Stripe, Payhip or Gumroad payment link into `link`. Sale emails come from that platform, so sign up there with husseinylaura@gmail.com.
