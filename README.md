# RentSheet — Rental Property Income & Expense Tracker (landlords, 1–20 units)

**Live:** https://makavelimachiavelli.github.io/rentsheet/

## What it is
A one-time $6.99 (or ₱199 GCash) Excel/Google-Sheets workbook for small landlords: PROPERTIES with automatic ⚠️ EXPIRED lease alerts, RENT LOG with ✅ PAID / 🟡 PARTIAL / 🔴 LATE status per month, per-unit EXPENSES with dropdowns, and a DASHBOARD (collected vs expected, collection rate, unpaid balance, net income, per-unit summary). Free limited demo as lead magnet.

## Buyer persona
- **Who:** small landlords (1–20 units) — often accidental landlords with a day job; global, English-speaking.
- **Pain:** PMS software is $20–50/month per unit and overkill; tracking rent in notebooks means late rent and lapsed leases go unnoticed; accountant asks for numbers they can't produce quickly.
- **Why pay $6.99:** one row per payment and the dashboard answers "who owes me, which unit profits, what do I give my accountant". Less than one week of any PMS subscription.
- **Where they hang out:** r/Landlord, BiggerPockets forums, landlord Facebook groups, "rental income tracker spreadsheet" searches (Etsy/Gumroad buyers already exist).

## Demand evidence (per REVENUE GATES)
- Paid competitors/listings: MARCHBORNS landlord tracker on Gumroad, Spreadsheet Serenity rental analysis, My Digital Plan Notion dashboard, plus the Etsy "rental income tracker" marketplace category (multiple listings, typically $9–30). Gate passed (3+ paid + marketplace listings).

## Monetization
Same working code-gate as NegosyoSheet: card-link (LemonSqueezy/Gumroad — Allen 5-min setup, see `PAYMENTS.md`) or QR; unlock code reveals the full-file download. `/full/` blocked in robots.txt; Gumroad auto-delivery is the scale-up path.

## Tech
`build_sheet.py` (openpyxl) regenerates both files. Verified by **real formula evaluation** (17/17 — collection math, lease-expiry via TODAY(), per-unit SUMIFs). The evaluator caught a genuine bug pre-ship: sample dates written as ISO strings made date comparisons silently wrong in both the engine and real Excel — fixed to real date objects.

## Deploy
```bash
../toolkit/deploy-pages.sh . rentsheet
```

## Owner TODO (Allen, ~5 min)
Swap `pay-qr.svg`, set `PRO_CODES` in `app.js` (`RENTSHEET-699`, `RS-DEMO` placeholders), add card-link URL to `#payLink`. Cross-linked from hub + NegosyoSheet.
