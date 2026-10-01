// Flat monthly interest: always on the ORIGINAL principal (no reducing balance / EMI).
const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const clean = (n) => Number(n.toPrecision(12)); // strip float noise

export function calculateLoan(principal, ratePercent, tenure) {
  const monthlyInterest = clean((principal * ratePercent) / 100);
  const totalInterest = clean(monthlyInterest * tenure);
  const totalRepayment = clean(principal + totalInterest);
  const monthlyRepayment = clean(totalRepayment / tenure);
  return {
    principal: round2(principal), rate: ratePercent, tenure,
    monthlyInterest: round2(monthlyInterest), totalInterest: round2(totalInterest),
    totalRepayment: round2(totalRepayment), monthlyRepayment: round2(monthlyRepayment),
  };
}

// Bank payments arrive net of a deduction (4% => gross = bank / 0.96).
export function grossFromBank(bank, deductionPercent) {
  return round2(clean(bank / (1 - deductionPercent / 100)));
}

export function validateSettings(rate, tenure, auto, deduction = 4) {
  if (!Number.isFinite(rate) || rate < 0 || rate > 100) return 'Interest rate must be a number between 0 and 100.';
  if (!Number.isFinite(deduction) || deduction < 0 || deduction >= 100) return 'Bank deduction must be a number from 0 to under 100.';
  if (!auto && (!Number.isInteger(tenure) || tenure < 1 || tenure > 600)) return 'Tenure must be a whole number of months (1–600).';
  return null;
}

export function parseAmount(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : NaN;
  if (typeof v !== 'string') return NaN;
  const s = v.replace(/[₦,\s]|NGN/gi, '');
  return /^-?\d*\.?\d+$/.test(s) ? Number(s) : NaN;
}

// Whole months between two {y,m,d} dates, end day inclusive (1-Oct-2026 → 30-Sep-2027 = 12). No timezones.
export function monthsBetween(a, b) {
  const n = new Date(Date.UTC(b.y, b.m - 1, b.d + 1));
  return Math.round((n.getUTCFullYear() - a.y) * 12 + (n.getUTCMonth() + 1 - a.m) + (n.getUTCDate() - a.d) / 30);
}

export const formatNaira = (n) => '₦' + Number(n).toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
