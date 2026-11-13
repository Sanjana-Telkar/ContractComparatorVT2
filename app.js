/**
 * Contract Comparator — app.js
 * Pure client-side JavaScript. No server, no backend.
 * Reads Ironclad & Zycus Excel/CSV files, detects reference IDs,
 * maps columns, computes match scores, and generates a multi-sheet XLSX report.
 */

'use strict';

/* ═══════════════════════════════════════════════════════════
   0.  VALUE NORMALIZERS  (must be defined before FIELD_MAP)
   ═══════════════════════════════════════════════════════════ */

/**
 * Normalize Renewal Type: Ironclad uses "Evergreen", Zycus uses "Yes" for the same meaning.
 * Any other value falls back to bool-yes-no normalization.
 */
function normalizeRenewalType(v) {
  const s = String(v).trim().toLowerCase().replace(/-/g, ' ');
  if (s === 'evergreen' || s === 'yes' || s === 'true' || s === '1' || s === 'auto renew') return 'yes';
  if (s === 'no' || s === 'false' || s === '0') return 'no';
  return s;
}

/**
 * Score override for Region:
 * If either side is blank → 100. Only mismatch when both have values and they differ.
 */
function scoreRegion(ironVal, zyVal) {
  if (String(ironVal).trim() === '' || String(zyVal).trim() === '') return 100;
  return null;
}

/**
 * Score override for System Identification:
 * Ironclad concatenates words camelCase (e.g. "AuthoringToRepository"),
 * Zycus uses just one segment (e.g. "Repository" or "Authoring").
 * If every word in the shorter value appears in the longer → 100.
 */
function scoreSystemIdentification(ironVal, zyVal) {
  const ironWords = new Set(normalizeSystemIdentification(ironVal).split(' ').filter(Boolean));
  const zyWords   = new Set(normalizeSystemIdentification(zyVal).split(' ').filter(Boolean));
  if (!ironWords.size || !zyWords.size) return null;
  const smaller = ironWords.size <= zyWords.size ? ironWords : zyWords;
  const larger  = ironWords.size <= zyWords.size ? zyWords   : ironWords;
  if ([...smaller].every(w => larger.has(w))) return 100;
  return null;
}

/**
 * Score override for boolean fields:
 * Ironclad false/0/no + Zycus blank → 100 (blank Zycus means "No" implicitly).
 */
function scoreBoolBlank(ironVal, zyVal) {
  const iron = String(ironVal).trim().toLowerCase();
  const zy   = String(zyVal).trim();
  // Ironclad blank/false/no/dash  +  Zycus blank/0  → acceptable
  const ironBlank = iron === '' || iron === 'false' || iron === '0' || iron === 'no' || iron === '—' || iron === '–' || iron === 'none';
  const zyBlank   = zy   === '' || zy   === '0';
  if (ironBlank && zyBlank) return 100;
  return null;
}

/**
 * Score override for Payment Terms:
 * If Zycus payment_terms is blank, any Ironclad value is acceptable → 100.
 */
function scorePaymentTerms(ironVal, zyVal) {
  if (String(zyVal).trim() === '') return 100;
  return null;
}

/**
 * Score override for Renewal Type:
 * If perpetual_term (Zycus) = "No", Ironclad will be blank — that is expected → 100.
 */
function scoreRenewalType(ironVal, zyVal) {
  const zy   = String(zyVal).trim().toLowerCase();
  const iron = String(ironVal).trim().toLowerCase();
  // Zycus perpetual_term = No/0  +  Ironclad blank/None/dash → acceptable
  const zyNo      = zy   === 'no' || zy   === 'false' || zy   === '0';
  const ironBlank = iron === ''   || iron === 'none'  || iron === '—' || iron === '–';
  if (zyNo && ironBlank) return 100;
  return null; // fall through to normal scoring
}

/**
 * Score override for Is the Contract Confidential?:
 * Ironclad false + Zycus blank (—) is acceptable → 100.
 */
function scoreConfidential(ironVal, zyVal) {
  const iron = String(ironVal).trim().toLowerCase();
  const zy   = String(zyVal).trim().toLowerCase();
  if ((iron === 'false' || iron === '0' || iron === 'no') && zy === '') return 100;
  if (iron === '' && (zy === 'no' || zy === 'false' || zy === '0')) return 100;
  return null; // fall through to normal scoring (incl. normalizeBoolYesNo)
}

/**
 * Score override for Expiration Date:
 * If is_amendment=0, Ironclad will not have an Expiration Date — blank is expected → 100.
 * zyRow and resolved field are passed in so we can read is_amendment.
 */
function scoreExpirationDate(ironVal, zyVal, zyRow, f) {
  const isAmend = norm(zyRow?.[f.zyIsAmendmentIdx]);
  if ((isAmend === '0' || isAmend === '') && String(ironVal).trim() === '') return 100;
  return null; // fall through to normal scoring
}

/**
 * Normalize System Identification: Ironclad appends extra words to Zycus values.
 * e.g. "AuthoringToRepository" → "authoring", "Authoring" → "authoring"
 * Strip everything from the second capital letter onward (camelCase boundary).
 */
function normalizeSystemIdentification(v) {
  const s = String(v).trim().toLowerCase();
  // Known equivalences between Ironclad and Zycus system identification values
  // Two equivalence buckets:
  // 1. AuthoringToRepository (Ironclad) ↔ Authoring (Zycus)
  if (s === 'authoring' || s === 'authoringtorepository' || s === 'authoringtorepositary') return 'authoring';
  // 2. Repository (Ironclad) ↔ Bulk Upload (Zycus)  /  Repository (Zycus)
  if (s === 'repository' || s === 'bulk upload') return 'repository';
  // Split on camelCase boundary, drop connector words ("To", "And", etc.), lowercase
  const words = s
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .split(' ')
    .map(w => w.toLowerCase())
    .filter(w => w.length > 2); // drop short connectors like "to", "of"
  return words.join(' ');
}

/**
 * Treat "Local Agreement" and "Framework Agreement" as the same Contract Type.
 * Any other value is returned lowercased for standard comparison.
 */
function normalizeContractType(v) {
  const s = String(v).trim().toLowerCase();
  if (s === 'local agreement' || s === 'framework agreement') return 'framework agreement';
  return s;
}

/**
 * Normalize Operating Model: treat "Center" and "Centre" as identical (US/UK spelling).
 */
function normalizeOperatingModel(v) {
  return String(v).trim().toLowerCase().replace(/\bcenter\b/g, 'centre');
}

/**
 * Normalize Contract Value: strip leading 3-letter currency code (e.g. "GBP500000" → "500000").
 */
function normalizeContractValue(v) {
  const s = String(v).trim();
  // Treat dashes (em dash —, en dash –, hyphen -) as blank
  if (s === '—' || s === '–' || s === '-') return '';
  // Strip leading currency code (1–5 letters, e.g. USD, EUR, CHF)
  // then strip commas/spaces used as thousand separators
  return s.replace(/^[A-Za-z]{1,5}\s*/, '').replace(/,/g, '').trim();
}

/**
 * Score override for Contract Type:
 * Ironclad "Framework Agreement" (or "Local Agreement") + Zycus blank is acceptable → 100.
 */
function scoreContractType(ironVal, zyVal) {
  const iron = normalizeContractType(ironVal);
  const zy   = String(zyVal).trim();
  // Zycus blank or "0" (not filled) + Ironclad Framework/Local Agreement → acceptable
  if (iron === 'framework agreement' && (zy === '' || zy === '0')) return 100;
  return null;
}

/**
 * Normalize boolean-style values: false/0/no → "no", true/1/yes → "yes".
 * Ironclad stores these as false/true; Zycus stores them as No/Yes.
 */
function normalizeBoolYesNo(v) {
  const s = String(v).trim().toLowerCase();
  if (s === 'false' || s === '0' || s === 'no' || s === 'no/not applicable' || s === '—' || s === '–' || s === 'none') return 'no';
  if (s === 'true'  || s === '1' || s === 'yes') return 'yes';
  return s;
}

/* ═══════════════════════════════════════════════════════════
   1.  COLUMN MAPPING  (Ironclad column → Zycus column)
   ═══════════════════════════════════════════════════════════ */
const FIELD_MAP = [
  { label: 'Record Name',                       iron: 'Record Name',                         zy: 'contract_title' },
  { label: 'Agreement Date',                    iron: 'Agreement Date',                      zy: 'signedon',
    valueNorm: v => normalizeDate(String(v).trim()),
    scoreOverride: (a, b) => {
      // Ironclad dash (—) or blank + Zycus 0 or blank → not filled on either side → acceptable
      const ia = String(a).trim(); const ib = String(b).trim();
      const aBlank = ia === '' || ia === '—' || ia === '–';
      const bBlank = ib === '' || ib === '0';
      if (aBlank && bBlank) return 100;
      return null;
    }
  },
  { label: 'Contract Category',                 iron: 'Contract Category',                   zy: 'contract_category_value' },
  // "Local Agreement" and "Framework Agreement" in Ironclad are treated as equivalent (valueNorm applied)
  { label: 'Contract Type',                     iron: 'Contract Type',                       ironAlt: ['Contract Type_v1'], zy: 'document_type',          valueNorm: normalizeContractType, scoreOverride: scoreContractType },
  { label: 'Currency Code',                     iron: 'Currency Code',                       zy: 'base_currency_value',             scoreOverride: (a, b) => (String(a).trim() === '' || String(b).trim() === '') ? 100 : null },
  { label: 'Data Privacy Clause',               iron: 'Data Privacy Clause',                 zy: 'data_privacy',                    valueNorm: normalizeBoolYesNo },
  // Effective Date: use contract_effective_date when is_amendment=0, amendment_effective_date when is_amendment=1
  // zyAliases: fallback column names some Zycus exports use instead of contract_effective_date
  { label: 'Effective Date',  iron: 'Effective Date',
    zy: 'contract_effective_date',  zyAliases: ['effective date (date)', 'effective date'],
    zyAlt: 'amendment_effective_date',
    valueNorm: v => normalizeDate(String(v).trim()),
    zySelector: (zyRow, f) => norm(zyRow?.[f.zyIsAmendmentIdx]) === '1' && f.zyAltIdx !== -1 ? f.zyAltIdx : f.zyIdx },
  // Expiration Date: use contract_expiration_date when is_amendment=0, amendment_expiration_date when is_amendment=1
  // If is_amendment=0 and Ironclad has no Expiration Date, that is expected → score 100
  // ironAlt: Ironclad sometimes exports this column as "Expiration Date (Date)"
  // zyAliases: fallback column names some Zycus exports use instead of contract_expiration_date
  { label: 'Expiration Date', iron: 'Expiration Date', ironAlt: ['Expiration Date (Date)'],
    zy: 'contract_expiration_date', zyAliases: ['expiration date (date)', 'expiration date'],
    zyAlt: 'amendment_expiration_date',
    valueNorm: v => normalizeDate(String(v).trim()),
    zySelector: (zyRow, f) => norm(zyRow?.[f.zyIsAmendmentIdx]) === '1' && f.zyAltIdx !== -1 ? f.zyAltIdx : f.zyIdx,
    scoreOverride: scoreExpirationDate },
  // Contract Description: match against BOTH contract_description and amendment_description —
  // if either one matches the Ironclad value, it is not a mismatch.
  { label: 'Contract Description',              iron: 'Contract Description',                ironAlt: ['Contract Detail'],
    zy: 'contract_description', zyAlt: 'amendment_description',
    scoreOverride: (a, b, zyRow, f) => {
      const ia = a.trim();
      // Helper: score one Zycus description candidate against the Ironclad value
      function scoreDesc(zyCandidate) {
        // norm() handles mojibake fix + whitespace collapse + invisible chars
        const ib = norm(zyCandidate ?? '').trim();
        if ((ia === '' || ia === '—' || ia === '–') && (ib === '' || ib === '0')) return 100;
        // normalizeSymbols so curly quotes / dashes / ellipsis differences are ignored
        const na = normalizeSymbols(ia.toLowerCase());
        const nb = normalizeSymbols(ib.toLowerCase());
        if (na && nb && nb.startsWith(na)) return 100;
        if (ib === '') return null; // blank Zycus → skip (don't penalise)
        return similarity(na, nb);
      }
      // Score primary (contract_description) and alt (amendment_description), take the best
      const s1 = f.zyIdx    !== -1 ? scoreDesc(zyRow?.[f.zyIdx])    : null;
      const s2 = f.zyAltIdx !== -1 ? scoreDesc(zyRow?.[f.zyAltIdx]) : null;
      if (s1 === null && s2 === null) return null;
      return Math.max(s1 ?? 0, s2 ?? 0);
    }
  },
  { label: 'Is the Contract Confidential?',     iron: 'Is the Contract Confidential?',       zy: 'confidential_contract',           valueNorm: normalizeBoolYesNo, scoreOverride: scoreConfidential },
  { label: 'Operating Company',                 iron: 'Operating Company',                   zy: 'type_display_name',               scoreOverride: (a, b) => (String(a).trim() === '' || String(b).trim() === '') ? 100 : null },
  { label: 'Operating Model',                   iron: 'Operating Model',                     zy: 'operating_model',                 valueNorm: normalizeOperatingModel, scoreOverride: (a, b) => (String(a).trim() === '' || String(b).trim() === '') ? 100 : null },
  { label: 'Payment Terms',                     iron: 'Payment Terms',                       zy: 'payment_terms',                   scoreOverride: scorePaymentTerms },
  { label: 'Price Adjustment Formula (PAF):',   iron: 'Price Adjustment Formula (PAF):',     zy: 'price_adjustment_formula',        valueNorm: normalizeBoolYesNo },
  { label: 'Pricing in other than OpCo local Currency', iron: 'Pricing in other than OpCo local Currency', zy: 'pricing_currency', valueNorm: normalizeBoolYesNo, scoreOverride: scoreBoolBlank },
  { label: 'Rebate Clause in Contract',         iron: 'Rebate Clause in Contract',           zy: 'rebate_clause_contract',          valueNorm: normalizeBoolYesNo, scoreOverride: scoreBoolBlank },
  { label: 'Region',                            iron: 'Region',                              zy: 'region',                          scoreOverride: scoreRegion },
  { label: 'Renewal Type',                      iron: 'Renewal Type',                        zy: 'perpetual_term',                  valueNorm: normalizeRenewalType, scoreOverride: scoreRenewalType },
  { label: 'Supplier Finance',                  iron: 'Supplier Finance',                    zy: 'supplier_finance',                valueNorm: normalizeBoolYesNo, scoreOverride: scoreBoolBlank },
  { label: 'System Identification',             iron: 'System Identification',               zy: 'requested_from',                  valueNorm: normalizeSystemIdentification, scoreOverride: scoreSystemIdentification },
  { label: 'Total Contract Value',              iron: 'Total Contract Value (Monetary Amount)', ironAlt: ['Total Contract Value 1', 'Total Contract Value 2'], zy: 'contract_value', valueNorm: normalizeContractValue, rawValue: true, scoreOverride: (a, b) => (normalizeContractValue(a) === '' || normalizeContractValue(b) === '') ? 100 : null },
  { label: 'Underlying Commodity',              iron: 'Underlying Commodity',                zy: 'underlying_commodity',            valueNorm: normalizeBoolYesNo, scoreOverride: scoreBoolBlank },
  { label: 'Value/Volume Commitment',           iron: 'Value/Volume Commitment',             zy: 'value_commitment',                valueNorm: normalizeBoolYesNo, scoreOverride: scoreBoolBlank },
  // Zycus Contract Number (Ironclad) ↔ contract_number (Zycus)
  { label: 'Zycus Contract Number',             iron: 'Zycus Contract Number',               zy: 'contract_number' },
];

/* ═══════════════════════════════════════════════════════════
   2.  STATE
   ═══════════════════════════════════════════════════════════ */
const STATE = {
  ironFile:  null,
  zyFile:    null,
  pbFile:    null,   // Power BI ID-map file (optional)
  ironData:  null,   // { headers: [], rows: [] }
  zyData:    null,
  pbData:    null,   // { headers: [], rows: [] }
  result:    null,   // comparison result
  mode:      'pre', // 'pre' or 'post'
  preResult: null,   // pre-validation result
};

/* ═══════════════════════════════════════════════════════════
   2b.  PRE-VALIDATION ENGINE
   ═══════════════════════════════════════════════════════════ */

/**
 * Check a date field value:
 *   - Must not be blank or literal 0
 *   - Must contain an actual date component (not just a time string like "00:00:00")
 *   - Accepted formats: MM-DD-YYYY, YYYY-MM-DD, M/D/YYYY, etc. (any that include a year)
 * Returns null on pass, or an issue message string on fail.
 * labelPrefix is used in the message for modified_date (optional).
 */
function checkDateField(v, labelPrefix) {
  const s = norm(v);
  const prefix = labelPrefix ? `${labelPrefix}: ` : '';
  if (s === '' || s === '0') return `${prefix}Blank or zero value — must contain a valid date`;
  // Excel error values
  if (/^#(N\/A|NAME\?|REF!|VALUE!|DIV\/0!|NULL!|NUM!)/i.test(s)) return `${prefix}Excel error value "${s}" — must be a valid date`;
  // Reject pure time strings (HH:MM or HH:MM:SS with no date)
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(s)) return `${prefix}Contains only a time "${s}" — date part is missing`;
  // Must contain at least a 4-digit year somewhere
  if (!/\d{4}/.test(s)) return `${prefix}Value "${s}" does not appear to be a valid date — expected format like MM-DD-YYYY`;
  return null;
}

/** Excel formula-error pattern used in the whole-row scan */
const EXCEL_ERROR_RE = /^#(N\/A|NAME\?|REF!|VALUE!|DIV\/0!|NULL!|NUM!)/i;

/**
 * Pre-validation field rules.
 * Each rule defines:
 *   col        — Zycus column name (case-insensitive fuzzy match)
 *   label      — human-friendly name
 *   rule       — short description of the rule
 *   severity   — 'error' | 'warning'
 *   check(val, row, headers) — returns null (pass) or a string message (fail)
 */
const PRE_VALIDATION_RULES = [
  {
    col: 'contract_number',
    label: 'Contract Number',
    rule: 'Must not be blank',
    severity: 'error',
    check: v => norm(v) === '' ? 'Blank value' : null,
  },
  {
    col: 'id',
    label: 'ID',
    rule: 'Must not be blank',
    severity: 'error',
    check: v => norm(v) === '' ? 'Blank value' : null,
  },
  {
    col: 'is_amendment',
    label: 'Is Amendment',
    rule: 'Must be 0 or 1',
    severity: 'error',
    check: v => {
      const s = norm(v);
      if (s === '' ) return 'Blank value';
      if (s !== '0' && s !== '1') return `Invalid value "${s}" — expected 0 or 1`;
      return null;
    },
  },
  {
    col: 'sub_type_display_name',
    label: 'Sub Type Display Name',
    rule: 'Must not be blank',
    severity: 'error',
    check: v => norm(v) === '' ? 'Blank value' : null,
  },
  {
    col: 'contract_title',
    label: 'Contract Title',
    rule: 'Must not be blank',
    severity: 'error',
    check: v => norm(v) === '' ? 'Blank value' : null,
  },
  {
    col: 'contract_creation_date',
    label: 'Contract Creation Date',
    rule: 'Must not be blank; must contain a date (mm-dd-yyyy or similar), not just a time',
    severity: 'error',
    check: v => checkDateField(v),
  },
  {
    col: 'contract_value',
    label: 'Contract Value',
    rule: 'Must not be blank',
    severity: 'error',
    check: v => norm(v) === '' ? 'Blank value' : null,
  },
  {
    col: 'contract_category_value',
    label: 'Contract Category Value',
    rule: 'Must not be blank (ZZ is acceptable)',
    severity: 'error',
    check: v => norm(v) === '' ? 'Blank value' : null,
  },
  {
    col: 'base_currency_value',
    label: 'Base Currency Value',
    rule: 'Must not be blank',
    severity: 'error',
    check: v => norm(v) === '' ? 'Blank value' : null,
  },
  {
    col: 'contract_status',
    label: 'Contract Status',
    rule: 'Allowed: "Active", "Repository Published", "Repository Draft", "Repository Pending Published"',
    severity: 'error',
    check: v => {
      const s = norm(v).toLowerCase();
      if (s === '') return 'Blank value';
      const allowed = ['active', 'repository published', 'repository draft', 'repository pending published'];
      if (!allowed.includes(s)) return `Value is "${norm(v)}" — expected "Active", "Repository Published", "Repository Draft", or "Repository Pending Published"`;
      return null;
    },
  },
  {
    col: 'contract_source',
    label: 'Contract Source',
    rule: 'Allowed: "AuthoringToRepository", "Repository"',
    severity: 'error',
    check: v => {
      const s = norm(v).toLowerCase().replace(/\s+/g, '').trim();
      if (s === '') return 'Blank value';
      // Accept "AuthoringToRepository" (any casing/spacing) and "Repository"
      const allowed = ['authoringtorepositary', 'authoringtorepository', 'repository'];
      if (!allowed.includes(s)) return `Invalid value "${norm(v)}"`;
      return null;
    },
  },
  {
    col: 'requested_from',
    label: 'Requested From',
    rule: 'Allowed: "Authoring", "Repository", "Bulk upload"',
    severity: 'error',
    check: v => {
      const s = norm(v).toLowerCase().replace(/\s+/g, ' ').trim();
      if (s === '') return 'Blank value';
      const allowed = ['authoring', 'repository', 'bulk upload'];
      if (!allowed.includes(s)) return `Invalid value "${norm(v)}"`;
      return null;
    },
  },
  {
    col: 'contract_expiration_date',
    label: 'Contract Expiration Date',
    rule: 'Must not be blank; if blank, perpetual_term must be "Auto renew" or "Yes"',
    severity: 'error',
    check: (v, row, headers) => {
      const expiry = norm(v);
      if (expiry !== '') return null; // has expiry date — OK
      // expiry is blank — check perpetual_term
      const ptIdx = findColFuzzy(headers, 'perpetual_term');
      const pt = ptIdx !== -1 ? norm(row[ptIdx]).toLowerCase() : '';
      const validPt = ['auto renew', 'yes'];
      if (!validPt.includes(pt)) {
        return `Expiration date blank and perpetual_term is "${norm(pt) || '(blank)'}" — expected "Auto renew" or "Yes"`;
      }
      return null;
    },
  },
  {
    col: 'perpetual_term',
    label: 'Perpetual Term',
    rule: 'Allowed: "Auto renew", "No", "Yes". If "No" → expiration date must be present',
    severity: 'error',
    check: (v, row, headers) => {
      const s = norm(v).toLowerCase().replace(/\s+/g, ' ').trim();
      if (s === '') return 'Blank value';
      const allowed = ['auto renew', 'no', 'yes'];
      if (!allowed.includes(s)) return `Invalid value "${norm(v)}"`;
      if (s === 'no') {
        // expiration date must be present
        const expIdx = findColFuzzy(headers, 'contract_expiration_date');
        const exp = expIdx !== -1 ? norm(row[expIdx]) : '';
        if (exp === '') return 'perpetual_term is "No" but contract_expiration_date is blank';
      }
      return null;
    },
  },
  {
    col: 'confidential_contract',
    label: 'Confidential Contract',
    rule: 'Must not be blank',
    severity: 'error',
    check: v => norm(v) === '' ? 'Blank value' : null,
  },
  {
    col: 'payment_terms',
    label: 'Payment Terms',
    rule: 'If blank, OPCO must correct it',
    severity: 'warning',
    check: v => norm(v) === '' ? 'Blank — OPCO to correct' : null,
  },
  {
    col: 'operating_model',
    label: 'Operating Model',
    rule: 'Must not be blank; if blank, contract_status must be "Repository Draft"',
    severity: 'warning',
    check: (v, row, headers) => {
      const s = norm(v);
      if (s !== '') return null;
      const statusIdx = findColFuzzy(headers, 'contract_status');
      const status = statusIdx !== -1 ? norm(row[statusIdx]).toLowerCase() : '';
      if (status === 'repository draft') return null; // acceptable
      return `Blank operating_model (contract_status is "${norm(row[statusIdx] ?? '')||'(blank)'}" — expected "Repository Draft")`;
    },
  },
  {
    col: 'amendment_effective_date',
    label: 'Amendment Effective Date',
    rule: 'If is_amendment = 1, this date must be present',
    severity: 'error',
    check: (v, row, headers) => {
      const amendIdx = findColFuzzy(headers, 'is_amendment');
      const isAmend = amendIdx !== -1 ? norm(row[amendIdx]) : '';
      if (isAmend !== '1') return null; // not an amendment — skip
      if (norm(v) === '') return 'is_amendment = 1 but amendment_effective_date is blank';
      return null;
    },
  },
  {
    col: 'vendor_id',
    label: 'Vendor ID',
    rule: 'Must not be blank',
    severity: 'error',
    check: v => norm(v) === '' ? 'Blank value' : null,
  },
  {
    col: 'modified_date',
    label: 'Modified Date',
    rule: 'Should be present; must contain a date (mm-dd-yyyy or similar), not just a time',
    severity: 'warning',
    check: v => {
      const s = norm(v);
      if (s === '') return 'Modified date is blank';
      return checkDateField(v, 'Modified date');
    },
  },
];

/**
 * Run pre-validation on a parsed Zycus file.
 * Returns { issues, fieldSummary, passCount, failCount, totalRows }
 */
async function runPreValidation(zyData, onProgress) {
  const { headers, rows } = zyData;
  const issues = [];        // [{ rowNum, contractId, col, label, rule, value, message, severity }]
  const fieldSummary = [];  // per-rule aggregate

  // Detect contract ID column (hex or HNK)
  const idDet = detectIdColumn(headers, rows);
  const idCol = idDet.col;
  log(`Pre-validation ID col: ${idCol !== -1 ? `[${idCol}] "${headers[idCol]}"` : 'not found'}`, 'info');

  // Resolve Zycus contract_number column once
  const contractNumColIdx = findColFuzzy(headers, 'contract_number');

  onProgress(20, `Validating ${rows.length} records…`);

  for (let ri = 0; ri < rows.length; ri++) {
    if (ri % 200 === 0) {
      const pct = 20 + Math.round((ri / rows.length) * 60);
      onProgress(pct, `Checking row ${ri + 1} / ${rows.length}…`);
      await tick();
    }
    const row = rows[ri];
    const contractId = idCol !== -1 ? norm(row[idCol]) : `row_${ri + 2}`;
    const contractNumber = contractNumColIdx !== -1 ? norm(row[contractNumColIdx]) : '';

    for (const rule of PRE_VALIDATION_RULES) {
      const colIdx = findColFuzzy(headers, rule.col);
      if (colIdx === -1) continue; // column not in file — will be reported in summary
      const val = row[colIdx] ?? '';
      const msg = rule.check(val, row, headers);
      if (msg) {
        issues.push({
          rowNum: ri + 2,          // 1-based + header row offset
          contractId,
          contractNumber,
          col: rule.col,
          label: rule.label,
          rule: rule.rule,
          value: norm(val),
          message: msg,
          severity: rule.severity,
        });
      }
    }

    // Whole-row scan: flag any cell that contains an Excel error value (#NA, #NAME?, #REF!, etc.)
    for (let ci = 0; ci < headers.length; ci++) {
      const cellVal = norm(row[ci] ?? '');
      if (EXCEL_ERROR_RE.test(cellVal)) {
        issues.push({
          rowNum: ri + 2,
          contractId,
          contractNumber,
          col: headers[ci] || `col_${ci}`,
          label: headers[ci] || `Column ${ci + 1}`,
          rule: 'No Excel error values (#NA, #NAME?, #REF!, #VALUE!, etc.)',
          value: cellVal,
          message: `Excel error value "${cellVal}" found — fix formula or replace with actual data`,
          severity: 'error',
        });
      }
    }
  }

  onProgress(85, 'Building field summary…');
  await tick();

  // Build per-rule summary
  // rowsWithIssue covers ALL issues (including whole-row Excel-error scan) for accurate failCount
  const rowsWithIssue = new Set(issues.map(i => i.rowNum));
  for (const rule of PRE_VALIDATION_RULES) {
    const colIdx  = findColFuzzy(headers, rule.col);
    const present = colIdx !== -1;
    const ruleIssues = issues.filter(i => i.col === rule.col);
    fieldSummary.push({
      col:       rule.col,
      label:     rule.label,
      rule:      rule.rule,
      severity:  rule.severity,
      issueCount: ruleIssues.length,
      colPresent: present,
    });
  }
  // Add a synthetic summary row for Excel-error issues (not tied to a single defined rule)
  const excelErrIssues = issues.filter(i => EXCEL_ERROR_RE.test(i.value));
  if (excelErrIssues.length > 0) {
    fieldSummary.push({
      col:        '(any column)',
      label:      'Excel Error Values',
      rule:       'No #NA, #NAME?, #REF!, #VALUE!, #DIV/0!, #NULL!, #NUM! in any cell',
      severity:   'error',
      issueCount: excelErrIssues.length,
      colPresent: true,
    });
  }

  const failCount = rowsWithIssue.size;
  const passCount = rows.length - failCount;

  onProgress(95, 'Done!');
  return { issues, fieldSummary, passCount, failCount, totalRows: rows.length };
}

/* ═══════════════════════════════════════════════════════════
   3.  UTILITY HELPERS
   ═══════════════════════════════════════════════════════════ */

/**
 * Fix UTF-8 mojibake: Ironclad exports are sometimes saved as Latin-1 but read
 * as UTF-8, turning e.g. "ü" → "Ã¼", "±" → "Â±".
 * U+00C3 (Ã) is the lead byte for U+00C0–U+00FF (ä ö ü etc.)
 * U+00C2 (Â) is the lead byte for U+0080–U+00BF (± © ° etc.)
 * Re-encode each char as its Latin-1 byte value and decode as UTF-8.
 * If the round-trip produces the same or a longer string we leave it untouched.
 */
function fixMojibake(s) {
  // Quick bail-out: neither lead byte present → no mojibake to fix
  if (!s.includes('\u00C3') && !s.includes('\u00C2')) return s;
  try {
    const bytes = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i) & 0xFF;
    const decoded = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    // Only use the decoded result if it is strictly shorter (multi-byte → single char)
    return decoded.length < s.length ? decoded : s;
  } catch (e) {
    return s; // invalid UTF-8 sequence — leave original
  }
}

/** Normalize a raw value without date-serial conversion (for numeric fields like contract value) */
function normRaw(v) {
  if (v === null || v === undefined) return '';
  return fixMojibake(String(v))
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/[\u00AD\u200B\u200C\u200D\u2060\uFEFF]/g, '');
}

/** Normalize a value to a clean string for comparison */
function norm(v) {
  if (v === null || v === undefined) return '';
  // Handle Excel serial dates
  if (typeof v === 'number' && v > 25000 && v < 80000) {
    try {
      const d = XLSX.SSF.parse_date_code(v);
      if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
    } catch(e) {}
  }
  return fixMojibake(String(v))
    .trim()
    // collapse multiple internal whitespace (tabs, double spaces, \r\n) to single space
    .replace(/\s+/g, ' ')
    // strip invisible unicode: zero-width space, zero-width non-joiner, soft hyphen, BOM, etc.
    .replace(/[\u00AD\u200B\u200C\u200D\u2060\uFEFF]/g, '');
}

/** Case-insensitive, trimmed column header lookup */
function findCol(headers, name) {
  if (!name) return -1;
  const n = name.trim().toLowerCase();
  return headers.findIndex(h => norm(h).toLowerCase() === n);
}

/** Fuzzy column lookup — checks contains too. Returns -1 for null/empty name. */
function findColFuzzy(headers, name) {
  if (!name) return -1;
  let idx = findCol(headers, name);
  if (idx !== -1) return idx;
  const n = name.trim().toLowerCase();
  // substring match
  idx = headers.findIndex(h => norm(h).toLowerCase().includes(n) || n.includes(norm(h).toLowerCase()));
  return idx;
}

/**
 * Collapse typographic symbol variants to plain ASCII equivalents so that
 * e.g. a curly apostrophe in one system and a straight one in the other
 * do not cause a false mismatch.
 */
function normalizeSymbols(s) {
  return s
    // Smart / curly apostrophes → straight apostrophe
    .replace(/[\u2018\u2019\u02BC\u02B9\u0060]/g, "'")
    // Smart / curly double quotes → straight double quote
    .replace(/[\u201C\u201D\u201E\u201F\u2033\u02BA]/g, '"')
    // Em-dash / en-dash / figure dash → hyphen-minus
    .replace(/[\u2013\u2014\u2015\u2212]/g, '-')
    // Horizontal ellipsis → three dots
    .replace(/\u2026/g, '...')
    // Bullet / middle dot variants → asterisk
    .replace(/[\u2022\u2023\u00B7\u2043]/g, '*')
    // Non-breaking space / narrow NBSP → regular space
    .replace(/[\u00A0\u202F\u2009\u2008\u2007]/g, ' ')
    // Collapse any resulting multiple spaces
    .replace(/  +/g, ' ');
}

/** Compute similarity score 0–100 between two strings */
function similarity(a, b) {
  a = normalizeSymbols(norm(a).toLowerCase());
  b = normalizeSymbols(norm(b).toLowerCase());
  if (a === b) return 100;
  if (!a && !b) return 100;
  if (!a || !b) return 0;

  // Normalize common date formats before comparing
  a = normalizeDate(a);
  b = normalizeDate(b);
  if (a === b) return 100;

  // Day/month swap tolerance: if both look like YYYY-MM-DD, also accept the
  // swapped variant (Zycus sometimes stores month and day in reversed order).
  const dateRe = /^(\d{4})-(\d{2})-(\d{2})$/;
  const aDate = a.match(dateRe);
  const bDate = b.match(dateRe);
  if (aDate && bDate) {
    // Check if swapping MM and DD on b gives a
    const bSwapped = `${bDate[1]}-${bDate[3]}-${bDate[2]}`;
    if (a === bSwapped) return 100;
    // Also check swapping on a gives b (symmetric)
    const aSwapped = `${aDate[1]}-${aDate[3]}-${aDate[2]}`;
    if (aSwapped === b) return 100;
  }

  // Levenshtein-based
  const la = a.length, lb = b.length;
  const maxLen = Math.max(la, lb);
  if (maxLen === 0) return 100;
  const dist = levenshtein(a, b);
  return Math.round((1 - dist / maxLen) * 100);
}

function normalizeDate(s) {
  // Strip trailing timestamp (e.g. "2010-01-01 00:00:00" → "2010-01-01")
  s = s.replace(/^(\d{4}-\d{2}-\d{2})\s+\d{2}:\d{2}(:\d{2})?$/, '$1');
  // YYYY-M-D or YYYY-MM-D or YYYY-M-DD (ISO with missing leading zeros) → YYYY-MM-DD
  const ymd = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (ymd) return `${ymd[1]}-${ymd[2].padStart(2,'0')}-${ymd[3].padStart(2,'0')}`;
  // M/D/YYYY → YYYY-MM-DD
  const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdy) return `${mdy[3]}-${mdy[1].padStart(2,'0')}-${mdy[2].padStart(2,'0')}`;
  // D-M-YYYY → YYYY-MM-DD
  const dmy = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (dmy) return `${dmy[3]}-${dmy[1].padStart(2,'0')}-${dmy[2].padStart(2,'0')}`;
  return s;
}

function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({length: m+1}, (_, i) => Array.from({length: n+1}, (_,j) => i === 0 ? j : j === 0 ? i : 0));
  for (let i = 1; i <= m; i++)
    for (let j = 1; j <= n; j++)
      dp[i][j] = a[i-1] === b[j-1] ? dp[i-1][j-1] : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
  return dp[m][n];
}

/**
 * Detect the best ID column to use for joining rows across files.
 * Priority 1: column whose values are 32-char hex strings (e.g. 8a9e820b6a39e85e016a4a0dd7b65ac5)
 * Priority 2: column whose values start with HNK + digits
 * Scores each column by the fraction of sampled rows that match, picks the highest.
 */
function detectIdColumn(headers, rows) {
  const sampleSize = Math.min(rows.length, 50);
  let hexBest = { col: -1, score: 0 };
  let hnkBest = { col: -1, score: 0 };

  for (let ci = 0; ci < headers.length; ci++) {
    let hexHits = 0, hnkHits = 0;
    for (let ri = 0; ri < sampleSize; ri++) {
      const s = norm(rows[ri][ci]).toLowerCase();
      if (/^[0-9a-f]{32}$/.test(s))       hexHits++;
      else if (/^hnk\d+$/i.test(s))        hnkHits++;
    }
    const hexScore = hexHits / sampleSize;
    const hnkScore = hnkHits / sampleSize;
    if (hexScore > hexBest.score) hexBest = { col: ci, score: hexScore };
    if (hnkScore > hnkBest.score) hnkBest = { col: ci, score: hnkScore };
  }

  // Prefer hex ID column; fall back to HNK column
  if (hexBest.col !== -1 && hexBest.score > 0) return { col: hexBest.col, confidence: hexBest.score };
  if (hnkBest.col !== -1 && hnkBest.score > 0) return { col: hnkBest.col, confidence: hnkBest.score };
  return { col: -1, confidence: 0 };
}

/** Parse an uploaded file (xlsx/csv) → { headers, rows } */
async function parseFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: false, raw: false });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (raw.length < 2) return reject(new Error('File has no data rows'));
        // Find first non-empty row as header
        let hi = 0;
        while (hi < raw.length && raw[hi].every(c => !norm(c))) hi++;
        const headers = raw[hi].map(h => norm(h));
        const rows = raw.slice(hi + 1).filter(r => r.some(c => norm(c)));
        resolve({ headers, rows });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('FileReader error'));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Slim a parsed dataset down to only the columns needed for comparison.
 * Keeps every FIELD_MAP column (by fuzzy name match) plus the detected ID column.
 * All other columns are dropped — critical for Ironclad files with 1000+ columns.
 * Returns a new { headers, rows } with re-based column indices.
 */
function slimData(data, idCol) {
  const needed = new Set();
  needed.add(idCol);
  for (const field of FIELD_MAP) {
    const i = findColFuzzy(data.headers, field.iron);
    if (i !== -1) needed.add(i);
    // Also try ironAlt fallback column names
    if (field.ironAlt) {
      for (const alt of field.ironAlt) {
        const ia = findColFuzzy(data.headers, alt);
        if (ia !== -1) needed.add(ia);
      }
    }
    if (field.zy) { const j = findColFuzzy(data.headers, field.zy); if (j !== -1) needed.add(j); }
    // Also try zyAliases (e.g. "Expiration Date (Date)" for contract_expiration_date)
    if (field.zyAliases) {
      for (const alias of field.zyAliases) {
        const ja = findColFuzzy(data.headers, alias);
        if (ja !== -1) needed.add(ja);
      }
    }
    if (field.zyAlt) { const k = findColFuzzy(data.headers, field.zyAlt); if (k !== -1) needed.add(k); }
  }
  // Always keep is_amendment from Zycus so zySelector can read it after slimming
  const isAmendIdx = findColFuzzy(data.headers, 'is_amendment');
  if (isAmendIdx !== -1) needed.add(isAmendIdx);
  // Always keep Record Type from Ironclad for the Record Type breakdown tab
  const recTypeIdx = findColFuzzy(data.headers, 'Record Type');
  if (recTypeIdx !== -1) needed.add(recTypeIdx);

  const keepCols = [...needed].sort((a, b) => a - b);
  const headers = keepCols.map(ci => data.headers[ci]);
  const rows = data.rows.map(row => keepCols.map(ci => row[ci] ?? ''));
  return { headers, rows };
}

/* ═══════════════════════════════════════════════════════════
   4.  COMPARISON ENGINE
   ═══════════════════════════════════════════════════════════ */
async function runComparison(ironData, zyData, onProgress) {
  const result = {
    matched: [],          // [{ id, ironRow, zyRow, fieldScores: [{...}] }]
    ironOnly: [],         // IDs only in Ironclad
    zyOnly: [],           // IDs only in Zycus
    fieldSummary: [],     // aggregate per-field scores
    mismatches: [],       // flat list of per-contract mismatches
    missing: { iron: [], zy: [] }, // missing columns
    ironIdCol: -1,
    zyIdCol: -1,
    recordTypeCounts: [], // [{ recordType, count }] sorted desc
  };

  onProgress(5, 'Detecting ID columns…');

  // Auto-detect ID columns (hex ID preferred, HNK fallback)
  const ironDet = detectIdColumn(ironData.headers, ironData.rows);
  const zyDet   = detectIdColumn(zyData.headers,   zyData.rows);

  if (ironDet.col === -1) { log('Ironclad: no ID column found (expected 32-char hex or HNK number) — cannot match rows', 'error'); return result; }
  if (zyDet.col   === -1) { log('Zycus: no ID column found (expected 32-char hex or HNK number) — cannot match rows', 'error');    return result; }

  result.ironIdCol = ironDet.col;
  result.zyIdCol   = zyDet.col;

  log(`Ironclad ID col: [${ironDet.col}] "${ironData.headers[ironDet.col]}" (conf: ${Math.round(ironDet.confidence*100)}%)`, 'info');
  log(`Zycus ID col:    [${zyDet.col}]   "${zyData.headers[zyDet.col]}"   (conf: ${Math.round(zyDet.confidence*100)}%)`, 'info');

  // Drop unwanted columns before any further processing (handles 1000+ col Ironclad files)
  onProgress(12, 'Trimming unused columns…');
  const ironColsBefore = ironData.headers.length;
  const zyColsBefore   = zyData.headers.length;
  const ironIdHeader = ironData.headers[ironDet.col];
  const zyIdHeader   = zyData.headers[zyDet.col];
  ironData = slimData(ironData, ironDet.col);
  zyData   = slimData(zyData,   zyDet.col);
  // Re-locate ID columns in slimmed data by their known header name
  ironDet.col = ironData.headers.indexOf(ironIdHeader);
  zyDet.col   = zyData.headers.indexOf(zyIdHeader);
  result.ironIdCol = ironDet.col;
  result.zyIdCol   = zyDet.col;
  // HNK number lives in Zycus 'contract_number' column (always kept by slimData via FIELD_MAP)
  const zyHnkCol = findColFuzzy(zyData.headers, 'contract_number');
  log(`Ironclad columns: ${ironColsBefore} → ${ironData.headers.length} (kept only mapped fields)`, 'info');
  log(`Zycus columns: ${zyColsBefore} → ${zyData.headers.length} (kept only mapped fields)`, 'info');

  onProgress(15, 'Building ID index…');

  // Build lookup maps: id → row array  (lowercase for case-insensitive matching)
  const ironMap = new Map();
  for (const row of ironData.rows) {
    const id = norm(row[ironDet.col]).toLowerCase();
    if (id) { if (!ironMap.has(id)) ironMap.set(id, []); ironMap.get(id).push(row); }
  }
  const zyMap = new Map();
  for (const row of zyData.rows) {
    const id = norm(row[zyDet.col]).toLowerCase();
    if (id) { if (!zyMap.has(id)) zyMap.set(id, []); zyMap.get(id).push(row); }
  }

  log(`Ironclad unique IDs: ${ironMap.size}`, 'info');
  log(`Zycus unique IDs: ${zyMap.size}`, 'info');
  // Show first 3 sample IDs from each side to help diagnose mismatches
  log(`Ironclad sample IDs: ${[...ironMap.keys()].slice(0,3).join(' | ')}`, 'info');
  log(`Zycus sample IDs:    ${[...zyMap.keys()].slice(0,3).join(' | ')}`, 'info');

  // Find matched / only-in-iron (Zycus-only contracts are not considered)
  const allIronIds = [...ironMap.keys()];
  const allZyIds   = new Set(zyMap.keys());

  for (const id of allIronIds) {
    if (allZyIds.has(id)) {
      result.matched.push({ id, ironRows: ironMap.get(id), zyRows: zyMap.get(id) });
    } else {
      result.ironOnly.push(id);
    }
  }

  log(`Matched: ${result.matched.length} | Ironclad-only: ${result.ironOnly.length}`, 'ok');

  // Detect which mapped columns actually exist in each file
  onProgress(25, 'Resolving column mapping…');
  // Resolve is_amendment column index in Zycus (needed by zySelector fields)
  const zyIsAmendmentIdx = findColFuzzy(zyData.headers, 'is_amendment');

  const resolvedFields = [];
  for (const field of FIELD_MAP) {
    // Try primary iron column, then ironAlt fallbacks
    let ironIdx = findColFuzzy(ironData.headers, field.iron);
    // If ironAlt is defined, scan primary + alts and pick the first column with actual data
    if (field.ironAlt) {
      const candidates = [field.iron, ...field.ironAlt];
      let firstFound = -1;
      for (const candidate of candidates) {
        const ci = findColFuzzy(ironData.headers, candidate);
        if (ci === -1) continue;
        if (firstFound === -1) firstFound = ci; // remember first existing col as last-resort
        const norm_ = field.valueNorm || (x => x);
        const hasData = ironData.rows.some(r => norm_(norm(r[ci] ?? '')).trim() !== '');
        if (hasData) { ironIdx = ci; break; }
      }
      // If no column had data, fall back to the first existing one
      if (ironIdx === -1 || !ironData.rows.some(r => (field.valueNorm || (x => x))(norm(r[ironIdx] ?? '')).trim() !== '')) {
        if (firstFound !== -1) ironIdx = firstFound;
      }
    }
    let zyIdx          = field.zy    ? findColFuzzy(zyData.headers, field.zy)    : -1;
    // If primary zy column not found, try zyAliases (e.g. "Expiration Date (Date)" for contract_expiration_date)
    if (zyIdx === -1 && field.zyAliases) {
      for (const alias of field.zyAliases) {
        zyIdx = findColFuzzy(zyData.headers, alias);
        if (zyIdx !== -1) { log(`Column "${field.zy}" not found — matched alias "${alias}" at col ${zyIdx}`, 'info'); break; }
      }
    }
    const zyAltIdx     = field.zyAlt ? findColFuzzy(zyData.headers, field.zyAlt) : -1;
    const ironMissing  = ironIdx === -1;
    // zy: null means intentionally iron-only — don't report as missing
    const zyIronOnly   = !field.zy;
    // For conditional fields, missing only if BOTH zy and zyAlt are absent
    const zyMissing    = field.zyAlt ? (zyIdx === -1 && zyAltIdx === -1) : (zyIdx === -1);
    if (ironMissing) result.missing.iron.push(field.iron);
    if (zyMissing && !zyIronOnly) result.missing.zy.push(field.zy);
    resolvedFields.push({ ...field, ironIdx, zyIdx, zyAltIdx, zyIsAmendmentIdx, ironMissing, zyMissing: zyMissing || zyIronOnly });
  }

  log(`Missing in Ironclad: [${result.missing.iron.join(', ') || 'none'}]`, result.missing.iron.length ? 'warn' : 'ok');
  log(`Missing in Zycus: [${result.missing.zy.join(', ') || 'none'}]`, result.missing.zy.length ? 'warn' : 'ok');

  onProgress(35, `Comparing ${result.matched.length} contracts…`);

  // Per-field aggregate accumulators
  const fieldAcc = resolvedFields.map(() => ({ total: 0, count: 0 }));

  // Compare each matched contract
  const batchSize = 100;
  for (let mi = 0; mi < result.matched.length; mi++) {
    if (mi % batchSize === 0) {
      const pct = 35 + Math.round((mi / result.matched.length) * 50);
      onProgress(pct, `Comparing contract ${mi+1} / ${result.matched.length}…`);
      await tick(); // yield to UI
    }

    const m = result.matched[mi];
    m.fieldScores = [];

    // Pair iron rows & zy rows (1-to-1 if multi, zip)
    const pairs = zipRows(m.ironRows, m.zyRows);
    // HNK = Zycus contract_number — store on the matched entry so the report can use it
    const hnk = m.zyRows[0] && zyHnkCol !== -1 ? norm(m.zyRows[0][zyHnkCol]) : '';
    m.hnk = hnk;

    for (const [ironRow, zyRow] of pairs) {
      const rowScores = [];
      for (let fi = 0; fi < resolvedFields.length; fi++) {
        const f = resolvedFields[fi];
        let ironVal = '', zyVal = '';
        // rawValue fields use normRaw to avoid Excel serial-date conversion (e.g. contract value numbers)
        const readVal = f.rawValue ? normRaw : norm;
        if (!f.ironMissing) ironVal = readVal(ironRow?.[f.ironIdx] ?? '');
        if (!f.zyMissing) {
          // zySelector: pick column index based on per-row Zycus data (e.g. is_amendment)
          const zyColIdx = f.zySelector ? f.zySelector(zyRow, f) : f.zyIdx;
          zyVal = readVal(zyRow?.[zyColIdx] ?? '');
        }
        // Apply field-specific value normalizer before scoring (e.g. Contract Type equivalences)
        let scoreA = ironVal, scoreB = zyVal;
        if (f.valueNorm) { scoreA = f.valueNorm(scoreA); scoreB = f.valueNorm(scoreB); }
        // scoreOverride can short-circuit similarity for known expected blanks/equivalences
        let score;
        if (f.ironMissing || f.zyMissing) {
          score = null;
        } else if (zyVal.trim() === '') {
          // Zycus has no value for this field — skip entirely (don't count in summary, not a mismatch)
          score = null;
        } else if (ironVal.trim() === '' && (zyVal.trim() === '0' || scoreB === 'no' || scoreB === '')) {
          // Ironclad blank + Zycus "0" / normalised-no / blank → neither side has a value → acceptable
          score = 100;
        } else if (f.scoreOverride) {
          const override = f.scoreOverride(ironVal, zyVal, zyRow, f);
          score = override !== null ? override : similarity(scoreA, scoreB);
        } else {
          score = similarity(scoreA, scoreB);
        }

        rowScores.push({ fieldLabel: f.label, iron: f.iron, zy: f.zy, ironVal, zyVal, score, ironMissing: f.ironMissing, zyMissing: f.zyMissing });

        if (score !== null) {
          fieldAcc[fi].total += score;
          fieldAcc[fi].count++;
        }

        if (score !== null && score < 100) {
          result.mismatches.push({ id: m.id, hnk, fieldLabel: f.label, iron: f.iron, zy: f.zy, ironVal, zyVal, score });
        }
      }
      m.fieldScores.push(rowScores);
    }
  }

  // Build field summary + compute true overall score across all field comparisons
  let grandTotal = 0, grandCount = 0;
  for (let fi = 0; fi < resolvedFields.length; fi++) {
    const f = resolvedFields[fi];
    const avg = fieldAcc[fi].count > 0 ? Math.round(fieldAcc[fi].total / fieldAcc[fi].count * 10) / 10 : null;
    result.fieldSummary.push({
      label: f.label, iron: f.iron, zy: f.zy,
      avgScore: avg,
      ironMissing: f.ironMissing, zyMissing: f.zyMissing,
    });
    grandTotal += fieldAcc[fi].total;
    grandCount += fieldAcc[fi].count;
  }
  result.overallScore = grandCount > 0 ? Math.round(grandTotal / grandCount * 10) / 10 : null;

  // Build Record Type breakdown from the full Ironclad dataset (all rows, not just matched)
  const rtCol = findColFuzzy(ironData.headers, 'Record Type');
  if (rtCol !== -1) {
    const rtMap = new Map();
    for (const row of ironData.rows) {
      const rt = norm(row[rtCol]) || '(blank)';
      rtMap.set(rt, (rtMap.get(rt) || 0) + 1);
    }
    result.recordTypeCounts = [...rtMap.entries()]
      .map(([recordType, count]) => ({ recordType, count }))
      .sort((a, b) => b.count - a.count);
  }

  onProgress(90, 'Finalizing results…');
  return { result, ironData, zyData };
}

function zipRows(as, bs) {
  const len = Math.max(as.length, bs.length);
  const pairs = [];
  for (let i = 0; i < len; i++) pairs.push([as[i] ?? null, bs[i] ?? null]);
  return pairs;
}

function tick() { return new Promise(r => setTimeout(r, 0)); }

/* ═══════════════════════════════════════════════════════════
   5.  XLSX REPORT GENERATOR
   ═══════════════════════════════════════════════════════════ */
/* ── Cell style helpers ── */
function cs(bgHex, bold = false, fgHex = 'FFFFFF', sz = 11) {
  return {
    font:  { bold, color: { rgb: fgHex }, sz },
    fill:  { fgColor: { rgb: bgHex } },
    alignment: { vertical: 'center', wrapText: true }
  };
}

function styleRange(ws, rowIdx, colCount, style) {
  // rowIdx is 0-based
  for (let c = 0; c < colCount; c++) {
    const addr = XLSX.utils.encode_cell({ r: rowIdx, c });
    if (!ws[addr]) ws[addr] = { t: 's', v: '' };
    ws[addr].s = style;
  }
}

function generateReport(ironData, zyData, result) {
  const wb = XLSX.utils.book_new();

  /* ────────────────────────────────────────────
     Sheet 1: Ironclad Data
  ──────────────────────────────────────────── */
  const ironAoa = [ironData.headers, ...ironData.rows];
  const wsIron = XLSX.utils.aoa_to_sheet(ironAoa);
  styleDataSheet(wsIron, ironData.headers.length, ironData.rows.length);
  XLSX.utils.book_append_sheet(wb, wsIron, 'Ironclad Data');

  /* ────────────────────────────────────────────
     Sheet 2: Zycus Data
  ──────────────────────────────────────────── */
  const zyAoa = [zyData.headers, ...zyData.rows];
  const wsZy = XLSX.utils.aoa_to_sheet(zyAoa);
  styleDataSheet(wsZy, zyData.headers.length, zyData.rows.length);
  XLSX.utils.book_append_sheet(wb, wsZy, 'Zycus Data');

  /* ────────────────────────────────────────────
     Sheet 3: Comparing
  ──────────────────────────────────────────── */
  const compHeaders = ['Contract ID', 'Zycus Contract Number'];
  for (const fs of result.fieldSummary) {
    compHeaders.push(`${fs.iron} (Ironclad)`);
    compHeaders.push(`${fs.zy} (Zycus)`);
  }

  const compRows = [compHeaders];
  // track [rowIdx, colIdx] of every mismatch cell (both iron + zy col for that field)
  const mismatchCells = [];

  for (const m of result.matched) {
    for (const rowScores of m.fieldScores) {
      const rowIdx = compRows.length; // 0-based, header is row 0
      const row = [m.id, m.hnk || ''];
      for (let fi = 0; fi < rowScores.length; fi++) {
        const fs = rowScores[fi];
        row.push(fs.ironVal);
        row.push(fs.zyVal);
        // score < 100 and not null = real mismatch → mark both iron & zy cells
        // offset by 2 now (Contract ID + Zycus Contract Number)
        if (fs.score !== null && fs.score < 100) {
          const ironCol = 2 + fi * 2;
          const zyCol   = 2 + fi * 2 + 1;
          mismatchCells.push({ r: rowIdx, c: ironCol });
          mismatchCells.push({ r: rowIdx, c: zyCol });
        }
      }
      compRows.push(row);
    }
  }

  const wsComp = XLSX.utils.aoa_to_sheet(compRows);
  styleDataSheet(wsComp, compHeaders.length, compRows.length - 1);

  // Apply light red highlight to mismatched cells
  const mismatchStyle = cs('FEE2E2', false, '991B1B', 11); // light red bg, dark red text
  for (const { r, c } of mismatchCells) {
    const addr = XLSX.utils.encode_cell({ r, c });
    if (!wsComp[addr]) wsComp[addr] = { t: 's', v: '' };
    wsComp[addr].s = mismatchStyle;
  }
  XLSX.utils.book_append_sheet(wb, wsComp, 'Comparing');

  /* ────────────────────────────────────────────
     Sheet 4: Summary
  ──────────────────────────────────────────── */
  const summaryRows = [];
  // track which row indices need which heading style
  const sectionHeadingRows = [];   // deep navy section banners
  const tableHeaderRows    = [];   // column header rows inside tables
  const mismatchSubHeaderRows = []; // per-field sub-headers
  let r = 0; // row counter (0-based)

  // ── Title block ──
  summaryRows.push(['CONTRACT COMPARISON SUMMARY', '', '', '']);
  sectionHeadingRows.push(r++);
  summaryRows.push(['Generated:', new Date().toLocaleString(), '', '']);
  r++;
  summaryRows.push(['Matched Contracts:', result.matched.length, 'Contracts only in Ironclad (not in Zycus):', result.ironOnly.length]);
  r++;
  summaryRows.push([]);
  r++;

  // ── Field Match Scores section ──
  summaryRows.push(['FIELD MATCH SCORES', '', '', '']);
  sectionHeadingRows.push(r++);
  summaryRows.push(['Attribute (Ironclad / Zycus)', 'Ironclad Column', 'Zycus Column', 'Match Score (% across all contracts)']);
  tableHeaderRows.push(r++);

  for (const fs of result.fieldSummary) {
    summaryRows.push([
      `${fs.iron} / ${fs.zy}`,
      fs.iron,
      fs.zy,
      fs.avgScore !== null ? `${fs.avgScore}%` : 'N/A'
    ]);
    r++;
  }

  summaryRows.push([]);
  r++;

  // ── Ironclad-only contracts ──
  summaryRows.push(['CONTRACTS ONLY IN IRONCLAD (NOT IN ZYCUS)', '', '', '']);
  sectionHeadingRows.push(r++);
  summaryRows.push(['Count:', result.ironOnly.length, '', '']);
  r++;

  summaryRows.push([]);
  r++;

  // ── Field Mismatches ──
  summaryRows.push(['FIELD MISMATCHES  (Score < 100)', '', '', '']);
  sectionHeadingRows.push(r++);

  const byField = new Map();
  for (const mm of result.mismatches) {
    if (!byField.has(mm.fieldLabel)) byField.set(mm.fieldLabel, []);
    byField.get(mm.fieldLabel).push(mm);
  }

  for (const [fieldLabel, mms] of byField) {
    summaryRows.push([]);
    r++;
    const sampleField = mms[0];
    summaryRows.push([`${fieldLabel}`, '', `${sampleField.iron} (Ironclad)`, `${sampleField.zy} (Zycus)`, '']);
    mismatchSubHeaderRows.push(r++);
    summaryRows.push(['Contract ID', 'Zycus Contract Number', `${sampleField.iron} (Ironclad)`, `${sampleField.zy} (Zycus)`, 'Score']);
    tableHeaderRows.push(r++);
    for (const mm of mms) {
      summaryRows.push([mm.id, mm.hnk || '', mm.ironVal, mm.zyVal, mm.score]);
      r++;
    }
  }

  const wsSum = XLSX.utils.aoa_to_sheet(summaryRows);
  const COL_COUNT = 5;

  // Set column widths
  wsSum['!cols'] = [{ wch: 38 }, { wch: 26 }, { wch: 28 }, { wch: 28 }, { wch: 10 }];

  // Apply section heading style: dark navy, white bold text
  const navyStyle  = cs('DBEAFE', true, '1E3A5F', 12);  // light blue bg, dark navy text
  // Table column headers: soft indigo bg, dark text
  const tealStyle  = cs('EDE9FE', true, '3730A3', 11);
  // Mismatch field sub-headers: pale lavender bg, dark text
  const slateStyle = cs('F1F5F9', true, '334155', 11);

  for (const rowIdx of sectionHeadingRows)    styleRange(wsSum, rowIdx, COL_COUNT, navyStyle);
  for (const rowIdx of tableHeaderRows)        styleRange(wsSum, rowIdx, COL_COUNT, tealStyle);
  for (const rowIdx of mismatchSubHeaderRows)  styleRange(wsSum, rowIdx, COL_COUNT, slateStyle);

  XLSX.utils.book_append_sheet(wb, wsSum, 'Summary');

  /* ────────────────────────────────────────────
     Sheet 5: Record Type Breakdown
  ──────────────────────────────────────────── */
  if (result.recordTypeCounts && result.recordTypeCounts.length > 0) {
    const rtRows = [['Record Type', 'Count']];
    for (const { recordType, count } of result.recordTypeCounts) {
      rtRows.push([recordType, count]);
    }
    rtRows.push([]);
    rtRows.push(['Total', result.recordTypeCounts.reduce((s, r) => s + r.count, 0)]);

    const wsRt = XLSX.utils.aoa_to_sheet(rtRows);
    wsRt['!cols'] = [{ wch: 40 }, { wch: 12 }];
    const rtHeaderStyle = cs('DBEAFE', true, '1E3A5F', 11);
    styleRange(wsRt, 0, 2, rtHeaderStyle);
    // Bold the total row
    const totalRowIdx = rtRows.length - 1;
    styleRange(wsRt, totalRowIdx, 2, cs('EDE9FE', true, '3730A3', 11));
    XLSX.utils.book_append_sheet(wb, wsRt, 'Record Type Breakdown');
  }

  return wb;
}

/* ═══════════════════════════════════════════════════════════
   5b.  PRE-VALIDATION XLSX REPORT
   ═══════════════════════════════════════════════════════════ */
function generatePreReport(zyData, preResult) {
  const wb = XLSX.utils.book_new();
  const { issues, fieldSummary, passCount, failCount, totalRows } = preResult;

  /* ── Sheet 1: Zycus Data ── */
  const zyAoa = [zyData.headers, ...zyData.rows];
  const wsZy = XLSX.utils.aoa_to_sheet(zyAoa);
  styleDataSheet(wsZy, zyData.headers.length, zyData.rows.length);
  XLSX.utils.book_append_sheet(wb, wsZy, 'Zycus Data');

  /* ── Sheet 2: Validation Summary ── */
  const sumRows = [];
  const sectionRows = [], tableHdrRows = [];
  let r = 0;

  sumRows.push(['PRE-VALIDATION SUMMARY', '', '', '', '']); sectionRows.push(r++);
  sumRows.push(['Generated:', new Date().toLocaleString(), '', '', '']); r++;
  sumRows.push(['Total Records:', totalRows, 'Pass:', passCount, '']); r++;
  sumRows.push(['Records with Issues:', failCount, 'Total Issues:', issues.length, '']); r++;
  sumRows.push([]); r++;

  sumRows.push(['FIELD CHECK OVERVIEW', '', '', '', '']); sectionRows.push(r++);
  sumRows.push(['Field', 'Zycus Column', 'Rule', 'Issues Found', 'Status']); tableHdrRows.push(r++);

  for (const fs of fieldSummary) {
    const status = !fs.colPresent ? 'Column Missing' : fs.issueCount === 0 ? 'PASS' : fs.severity === 'error' ? 'FAIL' : 'WARNING';
    sumRows.push([fs.label, fs.col, fs.rule, !fs.colPresent ? 'N/A' : fs.issueCount, status]);
    r++;
  }

  sumRows.push([]); r++;
  sumRows.push(['ALL ISSUES', '', '', '', '', '', '']); sectionRows.push(r++);
  sumRows.push(['Row #', 'Contract ID', 'Zycus Contract Number', 'Field', 'Value Found', 'Issue', 'Severity']); tableHdrRows.push(r++);

  for (const issue of issues) {
    sumRows.push([issue.rowNum, issue.contractId, issue.contractNumber, issue.label, issue.value, issue.message, issue.severity.toUpperCase()]);
    r++;
  }

  const wsSum = XLSX.utils.aoa_to_sheet(sumRows);
  wsSum['!cols'] = [{ wch: 8 }, { wch: 30 }, { wch: 26 }, { wch: 34 }, { wch: 20 }, { wch: 50 }, { wch: 12 }];

  const navyStyle  = cs('DBEAFE', true, '1E3A5F', 12);
  const tealStyle  = cs('EDE9FE', true, '3730A3', 11);
  for (const rowIdx of sectionRows)  styleRange(wsSum, rowIdx, 7, navyStyle);
  for (const rowIdx of tableHdrRows) styleRange(wsSum, rowIdx, 7, tealStyle);

  // Highlight error rows in light red, warning rows in light yellow
  const errorStyle   = cs('FEE2E2', false, '991B1B', 11);
  const warningStyle = cs('FFFBEB', false, '92400E', 11);
  let dataRowStart = tableHdrRows[tableHdrRows.length - 1] + 1;
  for (let i = dataRowStart; i < sumRows.length; i++) {
    const sev = (sumRows[i][6] || '').toString().toLowerCase();
    if (sev === 'error')   styleRange(wsSum, i, 7, errorStyle);
    if (sev === 'warning') styleRange(wsSum, i, 7, warningStyle);
  }

  XLSX.utils.book_append_sheet(wb, wsSum, 'Validation Report');

  /* ── Sheet 3: Issues by Field ── */
  const byFieldRows = [['Field', 'Zycus Column', 'Row #', 'Contract ID', 'Value Found', 'Issue', 'Severity']];
  const byFieldHdr = [0];
  const byFieldSectionRows = [];
  let bfr = 1;

  const byFieldMap = new Map();
  for (const issue of issues) {
    if (!byFieldMap.has(issue.col)) byFieldMap.set(issue.col, []);
    byFieldMap.get(issue.col).push(issue);
  }

  for (const [col, colIssues] of byFieldMap) {
    byFieldRows.push([`${colIssues[0].label} — ${col}`, '', '', '', '', '', '']);
    byFieldSectionRows.push(bfr++);
    for (const issue of colIssues) {
      byFieldRows.push(['', issue.col, issue.rowNum, issue.contractId, issue.value, issue.message, issue.severity.toUpperCase()]);
      bfr++;
    }
  }

  const wsByField = XLSX.utils.aoa_to_sheet(byFieldRows);
  wsByField['!cols'] = [{ wch: 30 }, { wch: 26 }, { wch: 8 }, { wch: 34 }, { wch: 20 }, { wch: 50 }, { wch: 10 }];
  for (const rowIdx of byFieldHdr)       styleRange(wsByField, rowIdx, 7, tealStyle);
  for (const rowIdx of byFieldSectionRows) styleRange(wsByField, rowIdx, 7, navyStyle);
  XLSX.utils.book_append_sheet(wb, wsByField, 'By Field');

  return wb;
}

function styleDataSheet(ws, colCount, rowCount) {
  const cols = Math.min(colCount, 50);
  ws['!cols'] = Array.from({ length: cols }, () => ({ wch: 22 }));
  // Auto-filter on header row
  ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rowCount || 1, c: cols - 1 } }) };
  // Light blue header row
  const headerStyle = cs('DBEAFE', true, '1E3A5F', 11);
  for (let c = 0; c < cols; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (!ws[addr]) ws[addr] = { t: 's', v: '' };
    ws[addr].s = headerStyle;
  }
}

/* ═══════════════════════════════════════════════════════════
   6.  UI RENDERING
   ═══════════════════════════════════════════════════════════ */
function renderScoreTable(fieldSummary) {
  const tbody = document.getElementById('score-tbody');
  tbody.innerHTML = '';
  for (const fs of fieldSummary) {
    const score = fs.avgScore;
    const missing = fs.ironMissing || fs.zyMissing;
    const color = missing ? '#64748b' : score >= 100 ? '#10b981' : score >= 80 ? '#3b82f6' : score >= 50 ? '#f59e0b' : '#ef4444';
    const pillClass = missing ? 'pill-missing' : score >= 100 ? 'pill-perfect' : score >= 80 ? 'pill-good' : score >= 50 ? 'pill-warn' : 'pill-bad';
    const pillText = missing ? (fs.ironMissing ? '⚠ Missing in Ironclad' : '⚠ Missing in Zycus') :
                     score >= 100 ? 'Perfect' : score >= 80 ? 'Good' : score >= 50 ? 'Review' : 'Poor';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${esc(fs.label)}</strong></td>
      <td style="color:var(--iron);font-size:11px">${esc(fs.iron)}</td>
      <td style="color:var(--zy);font-size:11px">${esc(fs.zy)}</td>
      <td>
        <div class="score-bar-wrap">
          <div class="score-bar-track">
            <div class="score-bar-fill" style="width:${missing?0:score}%;background:${color}"></div>
          </div>
          <span class="score-num" style="color:${color}">${missing ? 'N/A' : score % 1 === 0 ? score : score.toFixed(1)}</span>
        </div>
      </td>
      <td><span class="status-pill ${pillClass}">${pillText}</span></td>
    `;
    tbody.appendChild(tr);
  }
}

function renderComparisonTable(result) {
  const fields = result.fieldSummary;
  const thead = document.getElementById('comparison-thead');
  const tbody = document.getElementById('comparison-tbody');
  const fieldFilter = document.getElementById('comparison-field-filter');

  // Build headers
  let hHtml = '<tr><th>Contract ID</th>';
  for (const f of fields) {
    hHtml += `<th class="ironclad-h">${esc(f.iron)}</th><th class="zycus-h">${esc(f.zy)}</th>`;
  }
  hHtml += '</tr>';
  thead.innerHTML = hHtml;

  // Populate field filter
  fieldFilter.innerHTML = '<option value="">All Fields</option>';
  for (const f of fields) {
    fieldFilter.innerHTML += `<option value="${esc(f.iron)}">${esc(f.label)}</option>`;
  }

  renderComparisonRows(result, '', '');

  document.getElementById('comparison-search').addEventListener('input', e => {
    renderComparisonRows(result, e.target.value, document.getElementById('comparison-field-filter').value);
  });
  document.getElementById('comparison-field-filter').addEventListener('change', e => {
    renderComparisonRows(result, document.getElementById('comparison-search').value, e.target.value);
  });

  document.getElementById('comparison-count').textContent = `${result.matched.length} contracts`;
}

function renderComparisonRows(result, idFilter, fieldFilter) {
  const tbody = document.getElementById('comparison-tbody');
  const fields = result.fieldSummary;
  let html = '';
  let shown = 0;
  const maxDisplay = 500; // browser performance limit for preview

  for (const m of result.matched) {
    if (idFilter && !m.id.toLowerCase().includes(idFilter.toLowerCase())) continue;
    for (const rowScores of m.fieldScores) {
      if (shown >= maxDisplay) break;
      let row = `<tr><td class="td-id">${esc(m.id)}</td>`;
      for (let fi = 0; fi < fields.length; fi++) {
        const fs = rowScores[fi];
        if (fieldFilter && fields[fi].iron !== fieldFilter) {
          row += `<td class="td-empty">—</td><td class="td-empty">—</td>`;
          continue;
        }
        const match = fs.score === null ? '' : fs.score === 100 ? 'td-match' : 'td-mismatch';
        row += `<td class="${match}">${esc(fs.ironVal) || '<span class="td-empty">—</span>'}</td>`;
        row += `<td class="${match}">${esc(fs.zyVal) || '<span class="td-empty">—</span>'}</td>`;
      }
      row += '</tr>';
      html += row;
      shown++;
    }
  }
  if (shown >= maxDisplay) html += `<tr><td colspan="999" style="text-align:center;color:var(--muted);padding:16px">Showing first ${maxDisplay} rows — download full report for complete data</td></tr>`;
  tbody.innerHTML = html;
}

function renderMismatchTable(result) {
  const tbody = document.getElementById('mismatch-tbody');
  const fieldFilter = document.getElementById('mismatch-field-filter');

  // Populate field filter
  const uniqueFields = [...new Set(result.mismatches.map(m => m.fieldLabel))];
  fieldFilter.innerHTML = '<option value="">All Fields</option>';
  for (const f of uniqueFields) fieldFilter.innerHTML += `<option value="${esc(f)}">${esc(f)}</option>`;

  document.getElementById('mismatch-count').textContent = `${result.mismatches.length} mismatches`;
  renderMismatchRows(result.mismatches, '', '');

  document.getElementById('mismatch-search').addEventListener('input', e =>
    renderMismatchRows(result.mismatches, e.target.value, document.getElementById('mismatch-field-filter').value));
  document.getElementById('mismatch-field-filter').addEventListener('change', e =>
    renderMismatchRows(result.mismatches, document.getElementById('mismatch-search').value, e.target.value));
}

function renderMismatchRows(mismatches, search, fieldFilter) {
  const tbody = document.getElementById('mismatch-tbody');
  const maxDisplay = 1000;
  let html = '';
  let shown = 0;
  for (const mm of mismatches) {
    if (shown >= maxDisplay) break;
    if (search && !mm.id.toLowerCase().includes(search.toLowerCase()) &&
        !(mm.hnk || '').toLowerCase().includes(search.toLowerCase()) &&
        !mm.ironVal.toLowerCase().includes(search.toLowerCase()) &&
        !mm.zyVal.toLowerCase().includes(search.toLowerCase())) continue;
    if (fieldFilter && mm.fieldLabel !== fieldFilter) continue;
    html += `<tr>
      <td class="td-id">${esc(mm.id)}</td>
      <td style="color:var(--zy);font-weight:600">${esc(mm.hnk) || '<span class="val-empty">—</span>'}</td>
      <td style="color:var(--iron);font-size:11px">${esc(mm.iron)}</td>
      <td style="color:var(--zy);font-size:11px">${esc(mm.zy)}</td>
      <td class="val-iron">${esc(mm.ironVal) || '<span class="val-empty">—</span>'}</td>
      <td class="val-zy">${esc(mm.zyVal) || '<span class="val-empty">—</span>'}</td>
    </tr>`;
    shown++;
  }
  if (shown >= maxDisplay) html += `<tr><td colspan="6" style="text-align:center;color:var(--muted);padding:16px">Showing first ${maxDisplay} — download full report</td></tr>`;
  tbody.innerHTML = html;
}

function renderMissingCols(missing) {
  const container = document.getElementById('missing-cols-content');
  let html = '';
  if (missing.iron.length === 0 && missing.zy.length === 0) {
    html = '<div class="missing-card ok"><span class="missing-icon">✓</span><div class="missing-detail"><strong>All mapped columns found</strong><span>Both files contain all expected columns.</span></div></div>';
  }
  if (missing.iron.length > 0) {
    html += `<div class="missing-card error">
      <span class="missing-icon">✗</span>
      <div class="missing-detail">
        <strong>Missing in Ironclad file (${missing.iron.length} columns)</strong>
        <span>${missing.iron.map(c => `<code style="background:var(--surface3);padding:1px 6px;border-radius:4px;color:var(--red)">${esc(c)}</code>`).join(' ')}</span>
      </div>
    </div>`;
  }
  if (missing.zy.length > 0) {
    html += `<div class="missing-card warn">
      <span class="missing-icon">⚠</span>
      <div class="missing-detail">
        <strong>Missing in Zycus file (${missing.zy.length} columns)</strong>
        <span>${missing.zy.map(c => `<code style="background:var(--surface3);padding:1px 6px;border-radius:4px;color:var(--yellow)">${esc(c)}</code>`).join(' ')}</span>
      </div>
    </div>`;
  }
  container.innerHTML = html;
}

/* ── Pre-validation UI renders ── */
function renderPreStats(preResult) {
  document.getElementById('pre-stat-total').textContent = preResult.totalRows.toLocaleString();
  const passEl = document.getElementById('pre-stat-pass');
  passEl.textContent = preResult.passCount.toLocaleString();
  passEl.className = 'stat-val stat-val-green';
  const failEl = document.getElementById('pre-stat-fail');
  failEl.textContent = preResult.failCount.toLocaleString();
  failEl.className = preResult.failCount > 0 ? 'stat-val warn' : 'stat-val stat-val-green';
  document.getElementById('pre-stat-issues').textContent = preResult.issues.length.toLocaleString();
  document.getElementById('pre-stat-fields').textContent = PRE_VALIDATION_RULES.length;
}

function renderPreSummaryTable(fieldSummary) {
  const tbody = document.getElementById('pre-summary-tbody');
  let html = '';
  for (const fs of fieldSummary) {
    if (!fs.colPresent) {
      html += `<tr>
        <td><strong>${esc(fs.label)}</strong></td>
        <td style="color:var(--zy);font-size:11px">${esc(fs.col)}</td>
        <td style="color:var(--muted);font-size:11px">${esc(fs.rule)}</td>
        <td style="color:var(--muted)">N/A</td>
        <td><span class="status-pill pill-missing">⚠ Column Missing</span></td>
      </tr>`;
      continue;
    }
    const pillClass = fs.issueCount === 0 ? 'pill-perfect' :
                      fs.severity === 'error' ? 'pill-bad' : 'pill-warn';
    const pillText  = fs.issueCount === 0 ? 'PASS' :
                      fs.severity === 'error' ? `FAIL (${fs.issueCount})` : `WARNING (${fs.issueCount})`;
    html += `<tr>
      <td><strong>${esc(fs.label)}</strong></td>
      <td style="color:var(--zy);font-size:11px">${esc(fs.col)}</td>
      <td style="color:var(--muted);font-size:11px">${esc(fs.rule)}</td>
      <td style="font-weight:700;color:${fs.issueCount===0?'#34d399':fs.severity==='error'?'#f87171':'#fbbf24'}">${fs.issueCount}</td>
      <td><span class="status-pill ${pillClass}">${pillText}</span></td>
    </tr>`;
  }
  tbody.innerHTML = html;
  document.getElementById('pre-summary-chip').textContent = `${PRE_VALIDATION_RULES.length} checks`;
}

function renderPreIssuesTable(issues) {
  const fieldFilter = document.getElementById('pre-issues-field-filter');
  const uniqueFields = [...new Set(issues.map(i => i.label))];
  fieldFilter.innerHTML = '<option value="">All Fields</option>';
  for (const f of uniqueFields) fieldFilter.innerHTML += `<option value="${esc(f)}">${esc(f)}</option>`;

  document.getElementById('pre-issues-chip').textContent = `${issues.length} issues`;
  renderPreIssueRows(issues, '', '');

  document.getElementById('pre-issues-search').addEventListener('input', e =>
    renderPreIssueRows(issues, e.target.value, document.getElementById('pre-issues-field-filter').value));
  document.getElementById('pre-issues-field-filter').addEventListener('change', e =>
    renderPreIssueRows(issues, document.getElementById('pre-issues-search').value, e.target.value));
}

function renderPreIssueRows(issues, search, fieldFilter) {
  const tbody = document.getElementById('pre-issues-tbody');
  const maxDisplay = 1000;
  let html = '', shown = 0;
  for (const issue of issues) {
    if (shown >= maxDisplay) break;
    if (search && !issue.contractId.toLowerCase().includes(search.toLowerCase()) &&
        !issue.value.toLowerCase().includes(search.toLowerCase()) &&
        !issue.message.toLowerCase().includes(search.toLowerCase())) continue;
    if (fieldFilter && issue.label !== fieldFilter) continue;
    html += `<tr>
      <td style="color:var(--muted);font-size:11px">${esc(String(issue.rowNum))}</td>
      <td class="td-id">${esc(issue.contractId)}</td>
      <td style="color:var(--zy);font-size:11px">${esc(issue.label)}</td>
      <td class="val-iron">${esc(issue.value) || '<span class="val-empty">—</span>'}</td>
      <td style="color:var(--text2);font-size:11px">${esc(issue.message)}</td>
      <td><span class="status-pill ${issue.severity === 'error' ? 'pill-bad' : 'pill-warn'}">${issue.severity==='error'?'⛔ Error':'⚠ Warning'}</span></td>
    </tr>`;
    shown++;
  }
  if (shown >= maxDisplay) html += `<tr><td colspan="6" style="text-align:center;color:var(--muted);padding:16px">Showing first ${maxDisplay} — download full report</td></tr>`;
  tbody.innerHTML = html;
}

function renderPreByField(issues) {
  const container = document.getElementById('pre-by-field-content');
  const byFieldMap = new Map();
  for (const issue of issues) {
    if (!byFieldMap.has(issue.col)) byFieldMap.set(issue.col, { label: issue.label, issues: [] });
    byFieldMap.get(issue.col).issues.push(issue);
  }
  if (byFieldMap.size === 0) {
    container.innerHTML = '<div class="missing-card ok"><span class="missing-icon">✓</span><div class="missing-detail"><strong>No issues found!</strong><span>All records passed all validation checks.</span></div></div>';
    return;
  }
  let html = '';
  for (const [col, data] of byFieldMap) {
    html += `<div class="by-field-group">
      <div class="by-field-header">
        <strong>${esc(data.label)}</strong>
        <span style="color:var(--zy);font-size:11px;margin-left:8px">${esc(col)}</span>
        <span class="chip" style="margin-left:auto">${data.issues.length} issue${data.issues.length===1?'':'s'}</span>
      </div>
      <table class="mismatch-table" style="width:100%">
        <thead><tr><th>Row #</th><th>Contract ID</th><th>Value</th><th>Issue</th><th>Severity</th></tr></thead>
        <tbody>`;
    for (const issue of data.issues) {
      html += `<tr>
        <td style="color:var(--muted);font-size:11px">${esc(String(issue.rowNum))}</td>
        <td class="td-id">${esc(issue.contractId)}</td>
        <td class="val-iron">${esc(issue.value) || '<span class="val-empty">—</span>'}</td>
        <td style="font-size:11px;color:var(--text2)">${esc(issue.message)}</td>
        <td><span class="status-pill ${issue.severity==='error'?'pill-bad':'pill-warn'}">${issue.severity==='error'?'⛔ Error':'⚠ Warning'}</span></td>
      </tr>`;
    }
    html += `</tbody></table></div>`;
  }
  container.innerHTML = html;
}

function renderRecordTypeTab(recordTypeCounts) {
  const tbody = document.getElementById('record-type-tbody');
  const total = recordTypeCounts.reduce((s, r) => s + r.count, 0);
  let html = '';
  for (const { recordType, count } of recordTypeCounts) {
    html += `<tr>
      <td style="color:var(--text);font-weight:500">${esc(recordType)}</td>
      <td style="font-weight:700;color:var(--accent);text-align:right">${count.toLocaleString()}</td>
    </tr>`;
  }
  html += `<tr style="border-top:2px solid var(--border2)">
    <td style="font-weight:700;color:var(--text)">Total</td>
    <td style="font-weight:700;color:var(--accent);text-align:right">${total.toLocaleString()}</td>
  </tr>`;
  tbody.innerHTML = html;
  document.getElementById('record-type-chip').textContent = `${recordTypeCounts.length} type${recordTypeCounts.length === 1 ? '' : 's'}`;
}

function renderStats(result) {
  const overall = result.overallScore !== null ? result.overallScore : 0;
  document.getElementById('stat-matched').textContent = result.matched.length.toLocaleString();
  document.getElementById('stat-ironclad-only').textContent = result.ironOnly.length.toLocaleString();
  document.getElementById('stat-avg-score').textContent = overall + '%';
  const missingTotal = result.missing.iron.length + result.missing.zy.length;
  const missingEl = document.getElementById('stat-missing-cols');
  missingEl.textContent = missingTotal;
  if (missingTotal > 0) missingEl.className = 'stat-val warn';
}

/* ═══════════════════════════════════════════════════════════
   7.  THEME TOGGLE
   ═══════════════════════════════════════════════════════════ */
(function initTheme() {
  const btn   = document.getElementById('theme-toggle');
  const moon  = btn.querySelector('.icon-moon');
  const sun   = btn.querySelector('.icon-sun');
  const saved = localStorage.getItem('theme');

  function applyTheme(light) {
    document.body.classList.toggle('light', light);
    moon.style.display = light ? 'none'  : '';
    sun.style.display  = light ? ''      : 'none';
  }

  applyTheme(saved !== 'dark');

  btn.addEventListener('click', () => {
    const isLight = document.body.classList.toggle('light');
    moon.style.display = isLight ? 'none' : '';
    sun.style.display  = isLight ? ''     : 'none';
    localStorage.setItem('theme', isLight ? 'light' : 'dark');
  });
})();

/* ═══════════════════════════════════════════════════════════
   8.  BACKGROUND CANVAS ANIMATION
   ═══════════════════════════════════════════════════════════ */
(function initCanvas() {
  const canvas = document.getElementById('bg-canvas');
  const ctx = canvas.getContext('2d');
  let w, h, dots = [];

  const COLORS = [
    [96,  165, 250],   // blue
    [192, 132, 252],   // purple
    [103, 232, 249],   // cyan
  ];

  function resize() {
    w = canvas.width = window.innerWidth;
    h = canvas.height = window.innerHeight;
  }

  function initDots() {
    dots = [];
    const count = Math.floor((w * h) / 14000);
    for (let i = 0; i < count; i++) {
      const c = COLORS[Math.floor(Math.random() * COLORS.length)];
      dots.push({
        x:  Math.random() * w,
        y:  Math.random() * h,
        vx: (Math.random() - 0.5) * 0.25,
        vy: (Math.random() - 0.5) * 0.25,
        r:  Math.random() * 1.4 + 0.6,
        c:  c,
      });
    }
  }

  function draw() {
    ctx.clearRect(0, 0, w, h);
    // Draw connection lines
    for (let i = 0; i < dots.length; i++) {
      for (let j = i + 1; j < dots.length; j++) {
        const dx = dots[i].x - dots[j].x;
        const dy = dots[i].y - dots[j].y;
        const dist = Math.sqrt(dx*dx + dy*dy);
        if (dist < 130) {
          const alpha = 0.22 * (1 - dist / 130);
          const [r, g, b] = dots[i].c;
          ctx.beginPath();
          ctx.strokeStyle = `rgba(${r},${g},${b},${alpha})`;
          ctx.lineWidth = 0.5;
          ctx.moveTo(dots[i].x, dots[i].y);
          ctx.lineTo(dots[j].x, dots[j].y);
          ctx.stroke();
        }
      }
    }
    // Draw dots
    for (const d of dots) {
      const [r, g, b] = d.c;
      ctx.beginPath();
      ctx.arc(d.x, d.y, d.r, 0, Math.PI * 2);
      ctx.fillStyle = `rgba(${r},${g},${b},0.4)`;
      ctx.fill();
      // soft glow ring
      ctx.beginPath();
      ctx.arc(d.x, d.y, d.r * 2.0, 0, Math.PI * 2);
      ctx.fillStyle = `rgba(${r},${g},${b},0.04)`;
      ctx.fill();
      d.x += d.vx; d.y += d.vy;
      if (d.x < 0 || d.x > w) d.vx *= -1;
      if (d.y < 0 || d.y > h) d.vy *= -1;
    }
    requestAnimationFrame(draw);
  }

  window.addEventListener('resize', () => { resize(); initDots(); });
  resize(); initDots(); draw();
})();

/* ═══════════════════════════════════════════════════════════
   8.  LOG HELPER
   ═══════════════════════════════════════════════════════════ */
function log(msg, type = 'info') {
  const body = document.getElementById('log-body');
  const now = new Date().toTimeString().slice(0,8);
  const entry = document.createElement('span');
  entry.className = `log-entry ${type}`;
  entry.innerHTML = `<span class="log-time">${now}</span><span>${esc(msg)}</span>\n`;
  body.appendChild(entry);
  body.scrollTop = body.scrollHeight;
}

function esc(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* ═══════════════════════════════════════════════════════════
   9.  PROGRESS HELPERS
   ═══════════════════════════════════════════════════════════ */
function setBadge(state, label) {
  const badge = document.getElementById('header-badge');
  const badgeLabel = document.getElementById('badge-label');
  badge.className = 'header-badge';
  if (state !== 'ready') badge.classList.add('badge-' + state);
  badgeLabel.textContent = label;
}

function setProgress(pct, label) {
  document.getElementById('progress-bar').style.width = pct + '%';
  document.getElementById('progress-label').textContent = label;
  setBadge('processing', label);
}

/* ═══════════════════════════════════════════════════════════
   10.  DRAG-AND-DROP + FILE HANDLING
   ═══════════════════════════════════════════════════════════ */
function setupDropZone(dropId, inputId, stateKey, statusId) {
  const drop = document.getElementById(dropId);
  const input = document.getElementById(inputId);
  const status = document.getElementById(statusId);

  drop.addEventListener('click', () => input.click());
  input.addEventListener('change', e => handleFile(e.target.files[0]));

  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag-over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
  drop.addEventListener('drop', e => {
    e.preventDefault();
    drop.classList.remove('drag-over');
    handleFile(e.dataTransfer.files[0]);
  });

  function handleFile(file) {
    if (!file) return;
    STATE[stateKey] = file;
    status.textContent = `✓ ${file.name} (${(file.size/1024).toFixed(1)} KB)`;
    status.className = 'file-status ready';
    drop.classList.add('loaded');
    log(`Loaded ${stateKey}: ${file.name}`, 'ok');
    updateCompareBtn();
  }
}

function updateCompareBtn() {
  if (STATE.mode === 'pre') {
    document.getElementById('btn-compare').disabled = !STATE.zyFile;
  } else {
    document.getElementById('btn-compare').disabled = !(STATE.ironFile && STATE.zyFile);
  }
}

/**
 * Returns { country, extractType, env, suffix } or null if required fields are missing.
 * Pre-mode:  country + extractType required. env not required (empty string).
 * Post-mode: country + extractType + env all required.
 * Filename:
 *   Pre  → Country_ExtractType_Pre-Validation
 *   Post → Country_ExtractType_Environment_L1 Metadata Validation
 */
function getReportMeta() {
  const country     = document.getElementById('report-country').value.trim();
  const extractEl   = document.querySelector('input[name="extract-type"]:checked');
  const extractType = extractEl ? extractEl.value : '';
  const envEl       = document.querySelector('input[name="env-type"]:checked');
  const env         = envEl ? envEl.value : '';

  if (!country || !extractType) return null;
  if (STATE.mode === 'post' && !env) return null;

  return {
    country,
    extractType,
    env,
    suffix: STATE.mode === 'pre' ? 'Pre-Validation' : 'L1 Metadata Validation',
  };
}

/**
 * Enable the download button only when a result exists AND required meta fields are filled.
 * Shows a contextual hint / filename preview.
 */
function updateDownloadBtn() {
  const hasResult   = (STATE.mode === 'pre' && STATE.preResult) ||
                      (STATE.mode === 'post' && STATE.result);
  const meta        = getReportMeta();
  const hint        = document.getElementById('report-meta-hint');
  const country     = document.getElementById('report-country').value.trim();
  const extractType = (document.querySelector('input[name="extract-type"]:checked') || {}).value || '';
  const env         = (document.querySelector('input[name="env-type"]:checked')     || {}).value || '';

  document.getElementById('btn-download').disabled = !(hasResult && meta);

  if (hasResult && !meta) {
    let msg = '';
    if (!country)                              msg = 'Enter a Country name to enable download';
    else if (!extractType)                     msg = 'Select an Extract Type (Initial or Delta) to enable download';
    else                                       msg = 'Select an Environment (Acceptance or Production) to enable download';
    hint.textContent = msg;
    hint.classList.add('hint-error');
    document.getElementById('report-country').classList.toggle('input-error', !country);
  } else {
    if (meta) {
      const parts = [meta.country, meta.extractType];
      if (meta.env) parts.push(meta.env);
      parts.push(meta.suffix);
      hint.textContent = `Report will be saved as: ${parts.join('_')}.xlsx`;
    } else {
      hint.textContent = STATE.mode === 'pre'
        ? 'Fill in Country and select Extract Type to enable download'
        : 'Fill in Country, Extract Type and Environment to enable download';
    }
    hint.classList.remove('hint-error');
    document.getElementById('report-country').classList.remove('input-error');
  }
}

/* ═══════════════════════════════════════════════════════════
   11.  MAIN COMPARE FLOW
   ═══════════════════════════════════════════════════════════ */
document.getElementById('btn-compare').addEventListener('click', async () => {
  const btn = document.getElementById('btn-compare');
  btn.disabled = true;
  document.getElementById('btn-download').disabled = true;
  document.getElementById('progress-wrap').style.display = 'block';

  if (STATE.mode === 'pre') {
    // ── Pre-Validation flow ──
    document.getElementById('pre-stats-strip').style.display = 'none';
    document.getElementById('pre-results-section').style.display = 'none';

    try {
      setProgress(5, 'Parsing Zycus file…');
      log('Pre-Validation: parsing Zycus file…', 'info');
      STATE.zyData = await parseFile(STATE.zyFile);
      log(`Zycus: ${STATE.zyData.rows.length} rows, ${STATE.zyData.headers.length} columns`, 'ok');

      setProgress(15, 'Running pre-validation checks…');
      STATE.preResult = await runPreValidation(STATE.zyData, (pct, label) => setProgress(pct, label));

      setProgress(98, 'Rendering results…');
      await tick();

      renderPreStats(STATE.preResult);
      renderPreSummaryTable(STATE.preResult.fieldSummary);
      renderPreIssuesTable(STATE.preResult.issues);
      renderPreByField(STATE.preResult.issues);

      setProgress(100, 'Done!');
      setBadge(STATE.preResult.issues.length === 0 ? 'done' : 'done', `${STATE.preResult.issues.length} issue${STATE.preResult.issues.length === 1 ? '' : 's'} found`);
      document.getElementById('pre-stats-strip').style.display = 'grid';
      document.getElementById('pre-results-section').style.display = 'block';
      updateDownloadBtn();

      log(`Pre-validation complete. ${STATE.preResult.issues.length} issues found across ${STATE.preResult.totalRows} records.`, STATE.preResult.issues.length === 0 ? 'ok' : 'warn');

    } catch (err) {
      log(`Error: ${err.message}`, 'err');
      setProgress(0, 'Error — check log below');
      setBadge('error', 'Error');
      console.error(err);
    } finally {
      btn.disabled = false;
    }

  } else {
    // ── Post-Validation (comparison) flow ──
    document.getElementById('stats-strip').style.display = 'none';
    document.getElementById('results-section').style.display = 'none';

    try {
      setProgress(2, 'Parsing Ironclad file…');
      log('Parsing Ironclad file…', 'info');
      STATE.ironData = await parseFile(STATE.ironFile);
      log(`Ironclad: ${STATE.ironData.rows.length} rows, ${STATE.ironData.headers.length} columns`, 'ok');

      // Enrich Ironclad long IDs from Power BI map (if supplied)
      if (STATE.pbFile) {
        setProgress(6, 'Parsing Power BI ID map…');
        log('Parsing Power BI ID map…', 'info');
        STATE.pbData = await parseFile(STATE.pbFile);
        log(`Power BI map: ${STATE.pbData.rows.length} rows, ${STATE.pbData.headers.length} columns`, 'ok');

        // Locate the two columns we need in the Power BI file
        // Column called "ID" holds the 32-char hex long ID
        // Column called "Record Ironclad ID" holds the IC number (e.g. IC-141576)
        const pbLongCol = findColFuzzy(STATE.pbData.headers, 'ID');
        const pbIcCol   = findColFuzzy(STATE.pbData.headers, 'Record Ironclad ID');
        if (pbIcCol === -1 || pbLongCol === -1) {
          log('Power BI map: could not find "Record Ironclad ID" or "ID" columns — skipping enrichment', 'warn');
          log(`Power BI headers found: ${STATE.pbData.headers.slice(0,10).join(' | ')}`, 'info');
        } else {
          // Build IC-number → long-ID lookup (e.g. "ic-141576" → "8a54820e…")
          const pbMap = new Map();
          for (const row of STATE.pbData.rows) {
            const ic   = norm(row[pbIcCol]).toLowerCase();
            const long = norm(row[pbLongCol]);
            if (ic && long) pbMap.set(ic, long);
          }
          log(`Power BI map: ${pbMap.size} IC→LongID entries built`, 'info');

          // Find the IC number column in Ironclad (e.g. "Ironclad Id" / "Ironclad ID")
          // The IC number looks like "IC-141576"
          const ironIcIdCol = findColFuzzy(STATE.ironData.headers, 'Ironclad Id');
          if (ironIcIdCol === -1) {
            log('Ironclad: "Ironclad Id" column not found — cannot attach long IDs', 'warn');
            log(`Ironclad headers found: ${STATE.ironData.headers.slice(0,15).join(' | ')}`, 'info');
          } else {
            // Inject a new synthetic "__long_id__" column into Ironclad.
            // This becomes the hex ID column that detectIdColumn will prefer for matching.
            STATE.ironData.headers.push('__long_id__');
            let enriched = 0;
            for (const row of STATE.ironData.rows) {
              const icVal  = norm(row[ironIcIdCol]).toLowerCase();
              const longId = icVal ? (pbMap.get(icVal) || '') : '';
              row.push(longId);
              if (longId) enriched++;
            }
            log(`Power BI enrichment: ${enriched}/${STATE.ironData.rows.length} Ironclad rows received a long ID`, enriched > 0 ? 'ok' : 'warn');
          }
        }
      }

      setProgress(10, 'Parsing Zycus file…');
      log('Parsing Zycus file…', 'info');
      STATE.zyData = await parseFile(STATE.zyFile);
      log(`Zycus: ${STATE.zyData.rows.length} rows, ${STATE.zyData.headers.length} columns`, 'ok');

      const comparison = await runComparison(STATE.ironData, STATE.zyData, (pct, label) => {
        setProgress(pct, label);
      });
      STATE.result    = comparison.result;
      STATE.ironData  = comparison.ironData;
      STATE.zyData    = comparison.zyData;

      setProgress(95, 'Rendering results…');
      await tick();

      renderStats(STATE.result);
      renderScoreTable(STATE.result.fieldSummary);
      renderComparisonTable(STATE.result);
      renderMismatchTable(STATE.result);
      renderMissingCols(STATE.result.missing);
      if (STATE.result.recordTypeCounts && STATE.result.recordTypeCounts.length > 0) {
        renderRecordTypeTab(STATE.result.recordTypeCounts);
        document.getElementById('tab-record-type').style.display = '';
        document.querySelector('.tab[data-tab="record-type"]').style.display = '';
      } else {
        document.getElementById('tab-record-type').style.display = 'none';
        document.querySelector('.tab[data-tab="record-type"]').style.display = 'none';
      }

      setProgress(100, 'Done!');
      setBadge('done', `${STATE.result.matched.length} contracts compared`);
      document.getElementById('stats-strip').style.display = 'grid';
      document.getElementById('results-section').style.display = 'block';
      updateDownloadBtn();

      log(`Comparison complete. ${STATE.result.matched.length} contracts compared.`, 'ok');

    } catch (err) {
      log(`Error: ${err.message}`, 'err');
      setProgress(0, 'Error — check log below');
      setBadge('error', 'Error');
      console.error(err);
    } finally {
      btn.disabled = false;
    }
  }
});

/* ═══════════════════════════════════════════════════════════
   12.  DOWNLOAD HANDLER
   ═══════════════════════════════════════════════════════════ */
document.getElementById('btn-download').addEventListener('click', () => {
  const meta = getReportMeta();
  if (!meta) { updateDownloadBtn(); return; } // guard — shouldn't happen if button is properly gated
  // Build filename: Country_ExtractType[_Env]_Suffix
  const safeCountry = meta.country.replace(/[/\\?%*:|"<>]/g, '').replace(/\s+/g, '_');
  const fParts = [safeCountry, meta.extractType];
  if (meta.env) fParts.push(meta.env);
  fParts.push(meta.suffix);
  const safeName = fParts.join('_');
  log('Generating Excel report…', 'info');
  try {
    if (STATE.mode === 'pre' && STATE.preResult) {
      const wb = generatePreReport(STATE.zyData, STATE.preResult);
      XLSX.writeFile(wb, `${safeName}.xlsx`, { cellStyles: true });
    } else if (STATE.result) {
      const wb = generateReport(STATE.ironData, STATE.zyData, STATE.result);
      XLSX.writeFile(wb, `${safeName}.xlsx`, { cellStyles: true });
    }
    log(`Report downloaded as "${safeName}.xlsx".`, 'ok');
  } catch (err) {
    log(`Download error: ${err.message}`, 'err');
    console.error(err);
  }
});

/* ═══════════════════════════════════════════════════════════
   13.  TAB SWITCHER
   ═══════════════════════════════════════════════════════════ */
// Post-validation tabs
document.querySelectorAll('#results-section .tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('#results-section .tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('#results-section .tab-pane').forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById(`tab-${tab.dataset.tab}`).classList.add('active');
  });
});
// Pre-validation tabs
document.querySelectorAll('#pre-tab-bar .tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('#pre-tab-bar .tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('#pre-results-section .tab-pane').forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById(`tab-${tab.dataset.pretab}`).classList.add('active');
  });
});

/* ═══════════════════════════════════════════════════════════
   13b.  MODE SWITCHER
   ═══════════════════════════════════════════════════════════ */
document.querySelectorAll('.mode-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const mode = btn.dataset.mode;
    if (STATE.mode === mode) return;
    STATE.mode = mode;
    document.querySelectorAll('.mode-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');

    const uploadGrid = document.getElementById('upload-grid');
    const ironDrop   = document.getElementById('drop-ironclad');
    const btnLabel   = document.getElementById('btn-compare-label');
    const headerSub  = document.getElementById('header-sub');

    if (mode === 'pre') {
      ironDrop.style.display = 'none';
      document.getElementById('vs-divider-1').style.display = 'none';
      document.getElementById('drop-powerbi').style.display = 'none';
      document.getElementById('vs-divider-2').style.display = 'none';
      uploadGrid.style.gridTemplateColumns = '1fr';
      btnLabel.textContent = 'Run Pre-Validation';
      headerSub.textContent = 'Zycus — Pre-Upload Data Quality Check';
      // Hide post-validation results if shown
      document.getElementById('stats-strip').style.display = 'none';
      document.getElementById('results-section').style.display = 'none';
      // Pre-mode: hide Environment group, clear its selection
      document.getElementById('env-sep').style.display   = 'none';
      document.getElementById('env-field').style.display = 'none';
      document.querySelectorAll('input[name="env-type"]').forEach(r => { r.checked = false; });
    } else {
      ironDrop.style.display = '';
      document.getElementById('vs-divider-1').style.display = '';
      document.getElementById('drop-powerbi').style.display = '';
      document.getElementById('vs-divider-2').style.display = '';
      uploadGrid.style.gridTemplateColumns = '1fr auto 1fr auto 1fr';
      btnLabel.textContent = 'Run Comparison';
      headerSub.textContent = 'Ironclad × Zycus — Intelligent Field Matching Engine';
      // Hide pre-validation results if shown
      document.getElementById('pre-stats-strip').style.display = 'none';
      document.getElementById('pre-results-section').style.display = 'none';
      // Post-mode: show Environment group
      document.getElementById('env-sep').style.display   = '';
      document.getElementById('env-field').style.display = '';
    }
    // Reset files from the other mode
    STATE.ironFile  = null;
    STATE.zyFile    = null;
    STATE.pbFile    = null;
    STATE.pbData    = null;
    STATE.result    = null;
    STATE.preResult = null;
    ['status-ironclad','status-zycus','status-powerbi'].forEach(id => {
      const el = document.getElementById(id);
      el.textContent = 'No file selected';
      el.className = 'file-status';
    });
    ['drop-ironclad','drop-zycus','drop-powerbi'].forEach(id => document.getElementById(id).classList.remove('loaded'));
    updateDownloadBtn();
    updateCompareBtn();
  });
});

/* ═══════════════════════════════════════════════════════════
   14.  LOG CLEAR
   ═══════════════════════════════════════════════════════════ */
document.getElementById('log-clear').addEventListener('click', () => {
  document.getElementById('log-body').innerHTML = '';
});

/* ═══════════════════════════════════════════════════════════
   14b.  CLEAR ALL BUTTON
   ═══════════════════════════════════════════════════════════ */
document.getElementById('btn-clear').addEventListener('click', () => {
  // Reset state
  STATE.ironFile  = null;
  STATE.zyFile    = null;
  STATE.pbFile    = null;
  STATE.ironData  = null;
  STATE.zyData    = null;
  STATE.pbData    = null;
  STATE.result    = null;
  STATE.preResult = null;

  // Reset file drop zones
  ['drop-ironclad', 'drop-zycus', 'drop-powerbi'].forEach(id => {
    document.getElementById(id).classList.remove('loaded', 'drag-over');
  });
  document.getElementById('status-ironclad').textContent = 'No file selected';
  document.getElementById('status-ironclad').className   = 'file-status';
  document.getElementById('status-zycus').textContent    = 'No file selected';
  document.getElementById('status-zycus').className      = 'file-status';
  document.getElementById('status-powerbi').textContent  = 'No file selected';
  document.getElementById('status-powerbi').className    = 'file-status';
  // Reset file inputs so the same file can be re-selected
  document.getElementById('file-ironclad').value = '';
  document.getElementById('file-zycus').value    = '';
  document.getElementById('file-powerbi').value  = '';

  // Hide results & stats
  document.getElementById('progress-wrap').style.display      = 'none';
  document.getElementById('stats-strip').style.display        = 'none';
  document.getElementById('pre-stats-strip').style.display    = 'none';
  document.getElementById('results-section').style.display    = 'none';
  document.getElementById('pre-results-section').style.display = 'none';

  // Clear log
  document.getElementById('log-body').innerHTML = '';

  // Disable buttons
  updateDownloadBtn();
  updateCompareBtn();

  setBadge('ready', 'Ready');
  log('Cleared. Ready for new files.', 'ok');
});

/* ═══════════════════════════════════════════════════════════
   15.  INIT
   ═══════════════════════════════════════════════════════════ */
setupDropZone('drop-ironclad', 'file-ironclad', 'ironFile', 'status-ironclad');
setupDropZone('drop-zycus',    'file-zycus',    'zyFile',   'status-zycus');
setupDropZone('drop-powerbi',  'file-powerbi',  'pbFile',   'status-powerbi');

// Apply initial mode UI (Pre-Validation is default)
(function applyInitialMode() {
  document.getElementById('drop-ironclad').style.display = 'none';
  document.getElementById('vs-divider-1').style.display = 'none';
  document.getElementById('drop-powerbi').style.display = 'none';
  document.getElementById('vs-divider-2').style.display = 'none';
  document.getElementById('upload-grid').style.gridTemplateColumns = '1fr';
  document.getElementById('btn-compare-label').textContent = 'Run Pre-Validation';
  document.getElementById('header-sub').textContent = 'Zycus — Pre-Upload Data Quality Check';
  // Pre-validation is the default mode — hide Environment group
  document.getElementById('env-sep').style.display   = 'none';
  document.getElementById('env-field').style.display = 'none';
})();

// Meta bar inputs: re-evaluate download button whenever country or stage changes
document.getElementById('report-country').addEventListener('input', updateDownloadBtn);
document.querySelectorAll('input[name="extract-type"]').forEach(r => r.addEventListener('change', updateDownloadBtn));
document.querySelectorAll('input[name="env-type"]').forEach(r => r.addEventListener('change', updateDownloadBtn));

log('Contract Comparator ready. Drop your files to begin.', 'ok');
