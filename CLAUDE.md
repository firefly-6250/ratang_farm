# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a two-file ordering system for 亞輹恩.喇當的果園 (Ratang Farm), selling 椴木香菇 (log-grown shiitake mushrooms) from 白姑大山 in Nantou, Taiwan.

- **`order.html`** — Single-file, mobile-first, 4-step order form (select products → fill info → confirm/payment → success). No framework, no build step.
- **`Code.gs`** — Google Apps Script (GAS) backend deployed as a Web App. Handles `doGet` (returning customer lookup) and `doPost` (writing new orders).

## Architecture

The frontend and backend communicate as follows:

1. **Products** are loaded at page load from a public Google Sheets CSV export (`CSV_URL` in `order.html`). Expected columns: `id`, `name`, `unit`, `price`, `picture`.
2. **Customer lookup** hits `GAS_URL?action=getCustomer&phone=<phone>` via GET to pre-fill the form for returning customers.
3. **Order submission** POSTs JSON to `GAS_URL` with `Content-Type: text/plain` (required for GAS CORS compatibility). The GAS script parses `e.postData.contents`.

## Configuration (required before deployment)

Two placeholders must be replaced before the system works:

In **`order.html`**:
```js
const CSV_URL = 'https://docs.google.com/spreadsheets/d/.../pub?gid=0&single=true&output=csv';
const GAS_URL = 'YOUR_GAS_WEB_APP_URL';
```

In **`Code.gs`**:
```js
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';
```

## Google Sheets Structure

The spreadsheet (identified by `SPREADSHEET_ID`) requires two sheets:

**Sheet「訂單」** — one row per order, columns written in this order:
`訂單編號 | 時間 | 購買人 | 電話 | Email | 禮物 | 收件人 | 收件人電話 | 地址 | 備註 | 品項 | 金額 | 付款狀態`

**Sheet「客戶」** — customer lookup table, columns:
`電話 | 姓名 | Email | 地址`

**Sheet「商品」(or any sheet published as CSV)** — product catalog, columns:
`id | name | unit | price | picture`

## Deploying Code.gs

1. Open the Google Sheet → Extensions → Apps Script → paste `Code.gs`
2. Deploy → New Deployment → Type: Web App
3. Execute as: Me; Who has access: Anyone
4. Copy the deployment URL → paste into `GAS_URL` in `order.html`

## Payment Flow

Orders expect bank transfer (郵局 code 700, account 0401312-0059371). The QR code image is hosted at `https://i.meee.com.tw/5c5NlYZ.png`. Payment confirmation is manual — the seller verifies and updates the 付款狀態 column in the spreadsheet.
