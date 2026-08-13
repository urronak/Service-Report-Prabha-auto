# PRABHA AUTO PRO

Existing mobile-first PWA upgraded into a premium business-management dashboard without changing the existing local data key or Google Apps Script architecture.

## Existing functionality preserved
- Dashboard, Sale, Purchase, Payment, Stock, Party and Report navigation
- Local storage key: `prabha_auto_pro_v1`
- Google Sheets sync through Apps Script
- Daily/monthly reporting
- WhatsApp/share daily report
- Browser Print / Save PDF
- Stock calculation and negative-stock prevention
- PWA manifest, service worker and iPhone metadata

## Branding
- PRABHA AUTO
- Daily Report
- Service Manager: Chanchal Kumar
- Default supplier: Ganpati

Existing supplier/party records are not deleted or rewritten. The default value is now Ganpati for new purchases/payments.

## Data and business rules
Sale = Quantity × Selling Rate. Credit sales increase customer due. Cash sales increase cash. Bank and UPI are presented as one dashboard/report category named `Bank / UPI`, while the existing stored modes remain compatible.

Purchase = Quantity × Purchase Rate. Purchases increase stock and supplier balance.

Supplier Payment decreases supplier balance. Customer Receipt decreases customer due.

Profit / Difference uses the application's available daily sale minus purchase values; it is labeled as Difference rather than an invented accounting profit.

## Google Sheets
Keep `GOOGLE_SCRIPT_URL` configurable in `config.js`. No private credentials are stored in the frontend.

## PWA
Deploy over HTTPS, then use Safari > Share > Add to Home Screen on iPhone. The service worker caches the application shell for offline reopening. Google Sheet sync requires connectivity.

## Vercel
Static deployment; no build step is required. Deploy the existing project root. Authentication/deployment credentials are not stored in this project.
