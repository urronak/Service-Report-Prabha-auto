# PRABHA AUTO FINAL

This build is based on the working Current Stock version.

## Stock rules
- 284 `stockMaster:true` parts are the actual Current Stock master.
- Opening stock total is 1093; the built-in 10-piece invoice seed makes the displayed initial total 1103.
- PDF/catalogue is lookup-only. Its 18,000+ catalogue parts are not automatically added to stock.
- Add Part can use catalogue/PDF master data by Part Code and auto-fill Name, HSN, MRP, NDP and GST. Saving adds it to actual Current Stock.
- Sale decreases stock; Purchase increases stock.
- Low/Zero/Order List are calculated only from actual Current Stock.

## Dashboard
- `TOTAL SERVICE` is the number of sale/service entries in the selected period.
- Mechanic-wise Revenue shows service count + revenue for each machine/mechanic.
- This Month Purchase uses actual purchase entries, including supplier-specific purchases.
- Ganpati Automobile and Mahaveera Agro are default suppliers.

## Invoice
- Sale invoice has A4 Print / Save PDF.
- Edit and Delete are available from Sale History.
- Deleting an invoice restores its sold stock automatically.
- WhatsApp opens only from the WhatsApp Share button.

## Stock exports
- Export Excel exports the actual Current Stock table.
- Export PDF / Print creates an A4 stock report.

## Google Sheet
- `config.js` is prefilled with the deployed Web App URL supplied in the conversation.
- Updated Apps Script stores an exact AppData JSON backup plus Sales, Purchases, Payments, Stock and Suppliers tabs.
- Use `Sync to Sheet` to push the current device data.
- Use `Load from Sheet` to restore the saved central database on another device.
- After replacing the Apps Script code, deploy a new Web App version using the same deployment URL.
