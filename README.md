# Loan Flat Interest Calculator

Static, no-build web app. Upload an Excel file, calculate monthly **flat** interest on the original principal, preview, and download the updated `.xlsx`. Everything runs in the browser (SheetJS is bundled in `vendor/`), nothing is uploaded.

**Formulas:** `monthlyInterest = principal × rate%` · `totalInterest = monthlyInterest × tenure` · `totalRepayment = principal + totalInterest` · `monthlyRepayment = totalRepayment ÷ tenure`. Principal comes from the **New Principal** column (header matching is case/spacing tolerant).

Existing `Interest`, `Gross Loan`, `Monthly repayment` columns are updated in place; `Monthly Flat Rate`, `Monthly Interest`, `Loan Tenure` are appended if missing. Dates stay as Excel serial dates (no timezone conversion). Note: the free SheetJS build keeps data, number formats, dates, column widths and sheet names, but not cell colours/fonts.

## Run locally
    npx serve .        # or: python3 -m http.server
## Test
    npm test           # runs the two required test cases + workbook tests
## Deploy to Vercel
Push to GitHub → Import project in Vercel (Framework: **Other**, no build command, output dir = root), or run `npx vercel --prod`. `vercel.json` adds security headers.
