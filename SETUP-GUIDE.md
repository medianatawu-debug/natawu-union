# NATAWU GitHub Website — Setup Guide

## 1. GitHub Pages
Upload **everything inside this folder** to the root of your GitHub repository. `index.html` must be in the repository root. Then enable GitHub Pages from **Settings → Pages → Deploy from branch → main → /(root)**.

## 2. Membership application
`apply-membership.html` is the six-step online membership application. It already supports validation, electronic agreement ticks (in place of a hand-drawn signature), a review step and submission. The form expects a Google Apps Script Web App URL in its `SCRIPT_URL` setting. The same page also includes a working "Contact NATAWU" enquiry form, sharing the contact-form logic already used on `location.html`.

## 3. Google Sheets + email
1. Create a Google Sheet.
2. Open **Extensions → Apps Script**.
3. Paste `google-apps-script.gs`.
4. Set `SPREADSHEET_ID` to the ID in your Google Sheet URL.
5. Save.
6. Deploy as **Web app**, execute as **Me**, access **Anyone**.
7. Copy the Web App URL.
8. In `apply-membership.html`, replace `PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE` with that URL.
9. In `js/main.js`, replace the same placeholder with the URL (this powers every "Contact NATAWU" form on the site, including the one on the application page).

Submissions will be written to the Google Sheet and membership applications will be emailed to **sgift8083@gmail.com**.

## 4. Social media
Footer social icons are already connected to the supplied NATAWU Facebook, Instagram and TikTok pages. LinkedIn has been removed.

## 5. Important
The membership form contains legal/terms wording marked as draft in the supplied form. Verify the approved NATAWU wording before public launch.
