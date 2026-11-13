# Contract Comparator — Ironclad × Zycus

A **fully client-side** contract comparison tool. No server, no backend, no data leaves your browser.

## Live Demo
Deploy to GitHub Pages: `https://<your-username>.github.io/contract-comparator/`

## Features
- 📁 Drag-and-drop Ironclad & Zycus files (`.xlsx`, `.xls`, `.csv`)
- 🔍 **Auto-detects** the reference ID column — works with hex IDs (`8a9e820b...`) or contract numbers (`HNK136121`)
- 🗺️ **26-field column mapping** between Ironclad and Zycus naming conventions
- 📊 **Per-field match scores** (0–100) using normalized string similarity + date normalization
- 🔗 **Side-by-side comparison** of every matched contract
- ⚠️ **Mismatch drill-down** with searchable/filterable table
- 📋 **Missing column detection** — highlights fields not found in either file
- 📥 **Multi-sheet XLSX download**:
  - `Ironclad Data (Org)` — original Ironclad data
  - `Zycus Data (Org)` — original Zycus data
  - `Comparing` — side-by-side field values per contract ID
  - `Summary` — match scores, mismatches, available-only lists
- ⚡ Handles 3000+ contracts with async batch processing

## Column Mapping
| Ironclad | Zycus |
|----------|-------|
| Record Name | contract_title |
| Agreement Date | signedon |
| Contract Category | contract_category_value |
| Contract Owner Email - ZY | owner_email |
| Contract Type | document_type |
| Currency Code | base_currency_value |
| Data Privacy Clause | data_privacy |
| Effective Date | contract_effective_date |
| Expiration Date | amendment_expiration_date |
| Contract Description | contract_description |
| Is the Contract Confidential? | confidential_contract |
| Operating Company | type_display_name |
| Operating Model | operating_model |
| Payment Terms | payment_terms |
| Price Adjustment Formula (PAF): | price_adjustment_formula |
| Pricing in other than OpCo local Currency | pricing_currency |
| Rebate Clause in Contract | rebate_clause_contract |
| Region | region |
| Renewal Type | perpetual_term |
| Supplier Finance | supplier_finance |
| System Identification | requested_from |
| Total Contract Value 2 | contract_value |
| Underlying Commodity | underlying_commodity |
| Value/Volume Commitment | value_commitment |
| Record Type Id | Record Type |
| contract_number | Zycus Contract Number |

## How to Deploy on GitHub Pages

1. Fork or clone this repo
2. Push to GitHub
3. Go to **Settings → Pages → Source: Deploy from branch → main / root**
4. Your app is live at `https://<username>.github.io/<repo-name>/`

## File Structure
```
contract-comparator/
├── index.html      # App shell + UI
├── styles.css      # Dark futuristic theme
├── app.js          # Engine: parsing, comparison, report generation
└── README.md
```

## Tech Stack
- Vanilla JavaScript (ES2020)
- [SheetJS (xlsx)](https://sheetjs.com/) — Excel parsing & generation (CDN)
- CSS custom properties + canvas animation

## Notes
- Preview in browser tables is capped at 500/1000 rows for performance — the downloaded report contains **all data**
- Date values are normalized before scoring
- Ironclad columns with 1000+ columns are supported — only mapped fields are processed
