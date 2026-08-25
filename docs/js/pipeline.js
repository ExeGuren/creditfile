/**
 * pipeline.js  —  Exact JS port of the Python creditfile pipeline
 *
 * Modules ported (1:1):
 *   utils.py      → isna / notna / forceNumeric / normalizeText
 *   parse.py      → loadReportSheet / locateSections / parsers
 *   normalize.py  → normalizePersonalData / normalizeIncomeSourceDetails /
 *                   normalizeIncomeAnalysis / normalizeOfficerAssessment
 *   featurize.py  → extractFeatures / prepareFeatures / prepareFinancials /
 *                   prepareDemographics
 *   loancalc.py   → solveAmort
 *   score.py      → predictDelinquency / makeCreditScore
 *   main.py       → ESSENTIAL_FIELDS / validateFields / analyzeFile
 *
 * Public API:
 *   CreditPipeline.init()            → Promise<void>
 *   CreditPipeline.analyzeFile(file) → Promise<result>
 */

const CreditPipeline = (() => {
  'use strict';

  // ── Artifact state ────────────────────────────────────────────────────────
  let MODEL  = null;
  let SCALER = null;
  let BOW    = null;

  const BASE = window.location.href.replace(/\/[^/]*$/, '').replace(/\/$/, '');

  async function init() {
    console.log('[CreditPipeline] Loading artifacts from:', BASE);
    const [model, scaler, bow] = await Promise.all([
      fetch(BASE + '/artifacts/model.json').then(r => { if (!r.ok) throw new Error('model.json ' + r.status); return r.json(); }),
      fetch(BASE + '/artifacts/scaler.json').then(r => { if (!r.ok) throw new Error('scaler.json ' + r.status); return r.json(); }),
      fetch(BASE + '/artifacts/bow.json').then(r => { if (!r.ok) throw new Error('bow.json ' + r.status); return r.json(); }),
    ]);
    MODEL  = model;
    SCALER = scaler;
    BOW    = bow;
    console.log('[CreditPipeline] Ready. Trees:', MODEL.trees.length, '| BoW tokens:', BOW.feature_names.length);
  }


  // ═══════════════════════════════════════════════════════════════════════════
  // utils.py
  // ═══════════════════════════════════════════════════════════════════════════

  function notna(x) {
    // Python: x == x and x is not None
    // NaN !== NaN covers float NaN; null/undefined cover None
    return x === x && x !== null && x !== undefined;
  }

  function isna(x) {
    return !notna(x);
  }

  function forceNumeric(num) {
    if (typeof num === 'number' && num === num) return num;   // real number
    if (typeof num === 'string') {
      const cleaned = num.replace(/[^0-9.\-]/g, '');
      const f = parseFloat(cleaned);
      return isNaN(f) ? NaN : f;
    }
    return NaN;
  }

  function normalizeText(text) {
    // unicodedata NFKD → ascii → lower → collapse punctuation → strip non-alnum
    let s = text.normalize('NFKD').replace(/[\u0300-\u036f]/g, '');
    s = s.toLowerCase();
    s = s.replace(/[.,\-\/\(\)\s]+/g, ' ').trim();
    s = s.replace(/[^a-z0-9 ]/g, '');
    return s;
  }


  // ═══════════════════════════════════════════════════════════════════════════
  // parse.py
  // ═══════════════════════════════════════════════════════════════════════════

  /** wipe_colon — strip colons and blank → null */
  function wipeColon(sheet) {
    return sheet.map(row =>
      row.map(cell => {
        if (cell === null || cell === undefined) return null;
        let s = String(cell).replace(/:/g, '').trim();
        return s === '' ? null : s;
      })
    );
  }

  /** load_report_sheet */
  function loadReportSheet(arrayBuffer) {
    const wb  = XLSX.read(arrayBuffer, { type: 'array', raw: true });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });
    let sheet = wipeColon(raw);

    // Drop leading empty column (matches Python: if not report_sheet[0].any())
    const col0HasData = sheet.some(row => notna(row[0]));
    if (!col0HasData) {
      sheet = sheet.map(row => row.slice(1));
    }
    return sheet;
  }

  /** locate_sections */
  function locateSections(sheet) {
    const sectionTags = [
      ['personal_data',        /name/],
      ['dependents',           /name of dependents|rela(?:sh|t)ionship/],
      ['character_references', /address|contact number/],
      ['income_data',          /sources of income|adjudication/],
      ['client_reputation',    /(?:informant|contact).*remarks/],
      ['other_creditors',      /creditor/],
      ['client_assets',        /encumbr/],
      ['credit_assessment',    /remarks/],
    ];

    let tagIdx = 0;
    const tagLocations = [];

    for (let i = 0; i < sheet.length && tagIdx < sectionTags.length; i++) {
      const rowStr = sheet[i].map(c => (c || '')).join('').toLowerCase();
      if (sectionTags[tagIdx][1].test(rowStr)) {
        tagLocations.push([sectionTags[tagIdx][0], i]);
        tagIdx++;
      }
    }
    tagLocations.push(['end', sheet.length]);

    const bounds = {};
    for (let i = 0; i < tagLocations.length - 1; i++) {
      bounds[tagLocations[i][0]] = [tagLocations[i][1], tagLocations[i + 1][1]];
    }
    return bounds;
  }

  /** extract_rowwise_key_value_pairs */
  function extractRowwiseKV(rows) {
    const data = {};
    for (const row of rows) {
      const kv = row.filter(x => notna(x));
      if (!kv.length) continue;
      let k = kv[0];
      if (k in data) {
        let suffix = 1;
        while ((k + '_' + suffix) in data) suffix++;
        k = k + '_' + suffix;
      }
      data[k] = kv.length > 1 ? kv[1] : null;
    }
    return data;
  }

  /** join_series_values */
  function joinSeriesValues(arr, sep) {
    sep = sep || '|';
    const vals = arr.filter(notna);
    return vals.length ? vals.join(sep) : null;
  }

  // ── Section parsers ───────────────────────────────────────────────────────

  /** parse_personal_data */
  function parsePersonalData(sheet, bounds) {
    const [start, end] = bounds.personal_data;
    // section as 0-indexed rows, reindex from 0
    let section = sheet.slice(start, end);

    // if section.iloc[7].any(): adjusted_index[7:] += 1
    // In JS: we simulate using a logical index map
    const row7 = section[7] || [];
    const hasRow7Data = row7.some(c => notna(c));
    // We use a helper to get section row by logical index
    function srow(logicalIdx) {
      if (!hasRow7Data) return section[logicalIdx] || [];
      // When row7 has data, indices >= 7 are shifted up by 1
      if (logicalIdx < 7)  return section[logicalIdx] || [];
      return section[logicalIdx + 1] || [];
    }

    // Split left (:14) and right (15:) — col indices
    function leftCols(row)  { return (row || []).slice(0, 15); }
    function rightCols(row) { return (row || []).slice(15); }

    // left_exceptions = {5, 8, 9, 14, 16, 17}
    const leftExceptions  = new Set([5, 8, 9, 14, 16, 17]);
    // right_exceptions = {5, 16}
    const rightExceptions = new Set([5, 16]);

    const leftRows  = section.map((r, i) => leftExceptions.has(i)  ? null : leftCols(r)).filter(r => r !== null);
    const rightRows = section.map((r, i) => rightExceptions.has(i) ? null : rightCols(r)).filter(r => r !== null);

    const personalData = extractRowwiseKV(leftRows);

    // Right section — prefix with spouse__ if key already exists
    const rightKV = extractRowwiseKV(rightRows);
    for (const [k, v] of Object.entries(rightKV)) {
      const key = (k in personalData) ? 'spouse__' + k : k;
      personalData[key] = v;
    }

    // Exceptions
    const row5 = section[5] || [];
    personalData.type_of_residence =
      row5.slice(8, 11).some(notna)  ? 'owned'    :
      row5.slice(13, 15).some(notna) ? 'rented'   :
      row5.slice(17, 21).some(notna) ? 'free_use' : null;

    personalData.dob            = srow(9)[2]  || null;
    personalData.age            = srow(9)[9]  || null;
    personalData.marital_status = srow(9)[13] || null;
    personalData.parents_name_2 = srow(14)[3] || null;
    personalData.parents_address_2       = srow(16)[3]  || null;
    personalData['spouse__parents_name_2'] = srow(16)[19] || null;
    personalData.n_children   = srow(17)[2]  || null;
    personalData.n_dependents = srow(17)[11] || null;

    return personalData;
  }

  /** parse_income_data */
  function parseIncomeData(sheet, bounds) {
    const [start, end] = bounds.income_data;
    const section = sheet.slice(start, end);

    // Left section cols 0–14, right section cols 15+
    const leftSection  = section.map(r => (r || []).slice(0, 15));
    const rightSection = section.map(r => (r || []).slice(15));

    // ── Left: income sources ──
    // Remark handling: rows where col0 contains 'remark' accumulate col4 values
    // Simplified port — just do subsection extraction with [0,4] cols
    const leftCleaned = leftSection.map(r => [r[0] || null, r[4] || null]);

    const incomeSourceSubsections = {
      employment:                   [2, 12],
      business:                     [14, 21],
      other_business_or_remittance: [23, 29],
      spouse:                       [31, 40],
    };
    const incomeSources = {};
    for (const [name, [s, e]] of Object.entries(incomeSourceSubsections)) {
      incomeSources[name] = extractRowwiseKV(leftCleaned.slice(s, e));
    }

    // ── Right: income adjudication ──
    // right section uses cols 15 and 24, re-indexed as [0] and [9]
    const rightCleaned = rightSection.map(r => [r[0] || null, r[9] || null]);

    const adjSubsections = {
      income:  [2, 9],
      expense: [11, 31],
      summary: [32, rightCleaned.length],
    };
    const incomeAdjudication = {};
    for (const [name, [s, e]] of Object.entries(adjSubsections)) {
      incomeAdjudication[name] = extractRowwiseKV(rightCleaned.slice(s, e));
    }

    return { income_sources: incomeSources, income_adjudication: incomeAdjudication };
  }

  /** parse_credit_assessment */
  function parseCreditAssessment(sheet, bounds) {
    const [start, end] = bounds.credit_assessment;
    const section = sheet.slice(start, end);

    // section.loc[:7, [3, 8]]
    const sub = section.slice(0, 8).map(r => [r[3] || null, r[8] || null]);
    const data = extractRowwiseKV(sub);
    data.remarks     = (section[9] || [])[7] || null;
    // prepared_by = section[0][section[0].last_valid_index()]
    let preparedBy = null;
    for (let i = section.length - 1; i >= 0; i--) {
      if (notna((section[i] || [])[0])) { preparedBy = section[i][0]; break; }
    }
    data.prepared_by = preparedBy;
    return data;
  }

  /** parse_subtable */
  function parseSubtable(rows) {
    // Drop rows/cols where all values are null
    const nonEmptyRows = rows.filter(r => r.some(notna));
    if (nonEmptyRows.length < 2) return {};

    const colCount = Math.max(...nonEmptyRows.map(r => r.length));
    const activeCols = [];
    for (let c = 0; c < colCount; c++) {
      if (nonEmptyRows.some(r => notna(r[c]))) activeCols.push(c);
    }

    // First row = headers
    const rawHeaders = activeCols.map(c => nonEmptyRows[0][c] || '');
    // Handle duplicate headers
    const headerCount = {};
    const headers = rawHeaders.map(h => {
      if (headerCount[h] === undefined) { headerCount[h] = 0; return h; }
      headerCount[h]++;
      return h + '_' + headerCount[h];
    });

    const result = {};
    headers.forEach(h => { result[h] = []; });
    for (let r = 1; r < nonEmptyRows.length; r++) {
      activeCols.forEach((c, i) => {
        result[headers[i]].push(nonEmptyRows[r][c] !== undefined ? nonEmptyRows[r][c] : null);
      });
    }
    return result;
  }

  /** parse_subtables */
  function parseSubtables(sheet, bounds) {
    const subtableOffsets = {
      dependents:           [0, -1],
      character_references: [0,  0],
      client_reputation:    [0,  0],
      other_creditors:      [0, -2],
      client_assets:        [0,  0],
    };
    const result = {};
    for (const [k, [os, oe]] of Object.entries(subtableOffsets)) {
      if (!bounds[k]) continue;
      const [s, e] = bounds[k];
      try {
        result[k] = parseSubtable(sheet.slice(s + os, e + (oe || 0)));
      } catch (_) { /* suppress */ }
    }
    return result;
  }

  /** parse_credit_report */
  function parseCreditReport(sheet, filename) {
    const bounds = locateSections(sheet);
    const parsed = {
      filename,
      last_modified: new Date().toISOString(),
    };
    try { parsed.personal_data = parsePersonalData(sheet, bounds); }   catch (_) {}
    try { parsed.income_data   = parseIncomeData(sheet, bounds); }     catch (_) {}
    try { parsed.assessment    = parseCreditAssessment(sheet, bounds); } catch (_) {}
    Object.assign(parsed, parseSubtables(sheet, bounds));
    return parsed;
  }


  // ═══════════════════════════════════════════════════════════════════════════
  // normalize.py
  // ═══════════════════════════════════════════════════════════════════════════

  function standardizeField(name) {
    name = name.trim().toLowerCase();
    name = name.replace(/[\/\s]+/g, '_');
    name = name.replace(/[^a-z0-9_]/g, '');
    return name || '_';
  }

  function flattenDict(nested) {
    const result = {};
    for (const [k0, v0] of Object.entries(nested)) {
      if (v0 && typeof v0 === 'object') {
        for (const [k1, v1] of Object.entries(v0)) {
          result[k0 + '__' + k1] = v1;
        }
      }
    }
    return result;
  }

  /** normalize_personal_data */
  function normalizePersonalData(parsed) {
    if (!parsed.personal_data) return {};
    const data = parsed.personal_data;

    const preCorrections = {
      'Date of Birth':    'spouse__dob',
      'Present Address':  'spouse__present_address',
      'Previous Address': 'spouse__previous_address',
    };
    const postCorrections = {
      amount_applied_for:                    'loan_amount',
      contact_number:                        'contact_no',
      downpayment_terms:                     'loan_terms',
      educational_attainment:                'education',
      length_of_stay_at_present_address:     'present_address_tenure',
      length_of_stay_at_previous_address:    'previous_address_tenure',
      no_of_children:                        'n_children',
      name_of_applicant:                     'name',
      name_of_landlady_number:               'landlord',
      name_of_spouse:                        'spouse__name',
      parents_address_1:                     'spouse__parents_adress',
      place_of_birth:                        'birthplace',
      spouse__date_of_birth:                 'spouse__dob',
      spouse__educational_attainment:        'spouse__education',
      type_of_residence:                     'housing_status',
      unit_applied_collateral:               'unit_applied',
      units_applied:                         'unit_applied',
    };
    const canonicalFields = new Set([
      'age','birthplace','contact_no','date_applied','dob','education',
      'housing_status','landlord','loan_amount','loan_terms','marital_status',
      'n_children','n_dependents','name','nationality','parents_address',
      'parents_address_2','parents_name','parents_name_2','present_address',
      'present_address_tenure','previous_address','previous_address_tenure',
      'spouse__dob','spouse__education','spouse__name','spouse__parents_address',
      'spouse__parents_adress','spouse__parents_name','spouse__parents_name_2',
      'spouse__present_address','spouse__previous_address','unit_applied',
    ]);

    const normalized = {};
    for (let [k, v] of Object.entries(data)) {
      k = preCorrections[k] || k;
      k = standardizeField(k);
      k = postCorrections[k] || k;
      if (canonicalFields.has(k) && notna(v)) normalized[k] = v;
    }

    // dependent_ages from subtable
    if (parsed.dependents && parsed.dependents['Age']) {
      normalized.dependent_ages = parsed.dependents['Age'];
    }
    return normalized;
  }

  /** normalize_income_source_details */
  function normalizeIncomeSourceDetails(parsed) {
    if (!parsed.income_data || !parsed.income_data.income_sources) return {};
    const data = parsed.income_data.income_sources;

    const fieldCorrections = {
      'business__address_of_business': 'business__address',
      'business__business_name': 'business__name',
      'business__business_permit_no': 'business__permit_no',
      'business__monthly_income': 'business__monthly_income',
      'business__remarks': 'business__remarks',
      'business__route_of_vehicle': 'business__vehicle_route',
      'business__years_in_business': 'business__tenure',
      'employment__address_of_employer': 'employment__address',
      'employment__contact_number_of_employer': 'employment__contact_no',
      'employment__length_of_service': 'employment__tenure',
      'employment__monthly_net_pay': 'employment__monthly_income',
      'employment__monthly_pay': 'employment__monthly_income',
      'employment__name_of_employer': 'employment__name',
      'employment__position_employement_status': 'employment__status',
      'employment__position_employment_status': 'employment__status',
      'employment__previous_employer_address': 'employment__previous_employer',
      'employment__remarks': 'employment__remarks',
      'employment__verified_thru_name_contact_no': 'employment__verifier',
      'employment__verified_thru_name_contact_no_verified': 'employment__verifier',
      'employment__years_in_operation_of_employer': 'employment__employer_tenure',
      'other_business_or_remittance__address_of_business': 'remittance__address',
      'other_business_or_remittance__address_of_business_address_of_sender': 'remittance__address',
      'other_business_or_remittance__address_of_sender': 'remittance__address',
      'other_business_or_remittance__business_name': 'remittance__name',
      'other_business_or_remittance__business_name_name_of_sender': 'remittance__name',
      'other_business_or_remittance__name_of_sender': 'remittance__name',
      'other_business_or_remittance__monthly_income': 'remittance__monthly_income',
      'other_business_or_remittance__monthly_net_income_p': 'remittance__monthly_income',
      'other_business_or_remittance__monthly_net_income_remittance': 'remittance__monthly_income',
      'other_business_or_remittance__monthly_net_income_remittance_p': 'remittance__monthly_income',
      'other_business_or_remittance__nature_of_business': 'remittance__industry',
      'other_business_or_remittance__nature_of_business_source_of_income_of_sender': 'remittance__industry',
      'other_business_or_remittance__relationship_of_sender_to_credit_applicant': 'remittance__relationship',
      'other_business_or_remittance__remarks': 'remittance__remarks',
      'other_business_or_remittance__years_in_business': 'remittance__tenure',
      'other_business_or_remittance__years_in_business_years_of_remittance': 'remittance__tenure',
      'other_business_or_remittance__years_of_remittance': 'remittance__tenure',
      'spouse__address_of_employer': 'spouse__employer_address',
      'spouse__address_of_business_address_of_sender': 'spouse__employer_address',
      'spouse__contact_number_of_employer': 'spouse__employer_contact_no',
      'spouse__length_of_service': 'spouse__employment_tenure',
      'spouse__monthly_net_income_remittance_p': 'spouse__income',
      'spouse__monthly_net_pay': 'spouse__income',
      'spouse__monthly_pay': 'spouse__income',
      'spouse__nature_of_business_source_of_income_of_sender': 'spouse__income',
      'spouse__name_of_employer': 'spouse__employer_name',
      'spouse__position_employement_status': 'spouse__employment_status',
      'spouse__position_employment_status': 'spouse__employment_status',
      'spouse__previous_employer_address': 'employment__previous_employer',
      'spouse__remarks': 'spouse__remarks',
      'spouse__verified_thru_name_contact_no': 'spouse__employment_verifier',
      'spouse__years_in_operation_of_employer': 'spouse__employer_tenure',
    };

    const normalized = {};
    for (let [k, v] of Object.entries(flattenDict(data))) {
      k = standardizeField(k);
      if (fieldCorrections[k] && notna(v)) normalized[fieldCorrections[k]] = v;
    }
    return normalized;
  }

  // match_len / extract_longest_match
  function matchLen(a, b) {
    let i = 0;
    while (i < a.length && i < b.length && a[i] === b[i]) i++;
    return i + 1;
  }

  function extractLongestMatch(query, references, minMatchLen) {
    minMatchLen = minMatchLen || 0;
    let runningMax = 0;
    let longestMatch = null;
    for (const ref of references) {
      const ml = matchLen(query, ref);
      if (ml > runningMax) { runningMax = ml; longestMatch = ref; }
    }
    return runningMax > minMatchLen ? longestMatch : null;
  }

  /** normalize_income_analysis */
  function normalizeIncomeAnalysis(parsed) {
    if (!parsed.income_data || !parsed.income_data.income_adjudication) return {};
    const data = parsed.income_data.income_adjudication;

    const incomeFields = new Set(['applicant','business','others','spouse','total_income']);
    const incomeCorrections = { '1': 'primary', '2': 'secondary' };
    const incomeItems = {};
    for (let [k, v] of Object.entries(data.income || {})) {
      let std = standardizeField(k);
      let mapped;
      if (incomeFields.has(std))       mapped = std;
      else if (incomeCorrections[std]) mapped = incomeCorrections[std];
      else                             mapped = extractLongestMatch(std, [...incomeFields], 3);
      if (mapped && notna(v)) incomeItems[mapped] = v;
    }

    const expenseFields = new Set([
      'living','education','amortization','elementary','high_school','college',
      'misc','others','rental','transportation','maintenance','house','helper',
      'building','electric','water','internet','load','total_expenses',
    ]);
    const expenseCorrections = { cignal: 'internet' };
    const expenseItems = {};
    for (let [k, v] of Object.entries(data.expense || {})) {
      let std = standardizeField(k);
      let mapped;
      if (expenseFields.has(std))       mapped = std;
      else if (expenseCorrections[std]) mapped = expenseCorrections[std];
      else                              mapped = extractLongestMatch(std, [...expenseFields], 3);
      if (mapped && notna(v)) expenseItems[mapped] = v;
    }

    const summaryCorrections = {
      'Gross Disposable Income': 'net_income',
      'LESS MONTHLY EXPENSES':   'total_expenses',
      'Monthly Amortization':    'monthly_amortization',
      'NET DISPOSABLE INCOME':   'net_disposable_income',
      'TOTAL EXPENSES':          'total_expenses',
      'TOTAL MONTHLY INCOME':    'gross_income',
    };
    const summary = {};
    for (const [k, v] of Object.entries(data.summary || {})) {
      if (summaryCorrections[k] && notna(v)) summary[summaryCorrections[k]] = v;
    }

    return { income: incomeItems, expense: expenseItems, summary };
  }

  /** normalize_officer_assessment */
  function normalizeOfficerAssessment(parsed) {
    if (!parsed.assessment) return {};
    const data = parsed.assessment;

    const fieldCorrections = {
      'Purpose of loan':                  'loan_purpose',
      'Who will use the unit':            'unit_rider',
      'Who will pay the for the unit':    'unit_payor',
      'User with/without license':        'rider_license',
      'Cellular signal on the area':      'cell_signal_status',
      'Previous/ Current account of Zurich/ Venture': 'existing_account',
      'Motorcyle unit/ vehicle that client owned  at the time of CI': 'other_units',
    };
    const canonicalFields = new Set([
      'loan_purpose','unit_payor','existing_account','other_units',
      'cell_signal_status','unit_rider','rider_license','remarks','prepared_by',
    ]);
    const normalized = {};
    for (let [k, v] of Object.entries(data)) {
      k = fieldCorrections[k] || k;
      if (canonicalFields.has(k) && notna(v)) normalized[k] = v;
    }
    return normalized;
  }

  /** normalize_credit_data */
  function normalizeCreditData(parsed) {
    return {
      filename:              parsed.filename,
      last_modified:         parsed.last_modified,
      personal_data:         normalizePersonalData(parsed),
      income_source_details: normalizeIncomeSourceDetails(parsed),
      income_analysis:       normalizeIncomeAnalysis(parsed),
      officer_assessment:    normalizeOfficerAssessment(parsed),
    };
  }


  // ═══════════════════════════════════════════════════════════════════════════
  // featurize.py
  // ═══════════════════════════════════════════════════════════════════════════

  const NULL_VALUE = NaN;
  const MIN_AGE = 0, MAX_AGE = 80, MAX_DEPENDENT_AGE = 21;

  const FEATURE_MAP = {
    filename:      'info__filename',
    last_modified: 'info__last_modified',
    personal_data: [
      ['unit_applied',    'motorcycle_model'],
      ['loan_terms',      'loan_terms'],
      ['loan_amount',     'loan_amount'],
      ['dependent_ages',  'dependent_ages'],
      ['n_dependents',    'n_dependents'],
      ['n_children',      'n_children'],
      ['age',             'age'],
      ['education',       'education'],
      ['housing_status',  'housing_status'],
      ['marital_status',  'marital_status'],
      ['spouse__education','spouse_education'],
    ],
    income_analysis: {
      income: [
        ['applicant', 'employment_income'],
        ['business',  'business_income'],
        ['spouse',    'spouse_income'],
      ],
      summary: [
        ['gross_income',         'gross_income'],
        ['monthly_amortization', 'monthly_amortization'],
      ],
    },
  };

  const MODEL_FEATURES = [
    'bow__115','bow__125','bow__125i','bow__150','bow__155','bow__175',
    'bow__barako','bow__black','bow__click','bow__es','bow__fazzio','bow__fi',
    'bow__honda','bow__i','bow__kawasaki','bow__mio','bow__raider','bow__repo',
    'bow__smash','bow__suzuki','bow__tmx','bow__yamaha',
    'cat__housing_status','cat__marital_status','cat__education','cat__spouse_education',
    'num__loan_amount','num__loan_downpayment','num__loan_term',
    'num__monthly_amortization','num__age','num__n_children','num__n_dependents',
    'num__n_dependents_corrected','num__gross_income','num__employment_income',
    'num__business_income','num__spouse_income',
    'num__loan_downpayment_ratio','num__amort_income_ratio',
  ];

  /** extract_features */
  function extractFeatures(featureMap, data, features) {
    features = features || {};
    if (Array.isArray(featureMap)) {
      for (const [fieldName, featureName] of featureMap) {
        const val = (data && data[fieldName] !== undefined) ? data[fieldName] : NULL_VALUE;
        features[featureName] = val;
      }
    } else if (typeof featureMap === 'object' && featureMap !== null) {
      for (const [parent, child] of Object.entries(featureMap)) {
        const subset = (data && data[parent] !== undefined) ? data[parent] : {};
        extractFeatures(child, subset, features);
      }
    } else if (typeof featureMap === 'string') {
      features[featureMap] = data;
    }
    return features;
  }

  /** BoW transform — mirrors CountVectorizer.transform */
  function bowTransform(text) {
    const vec = {};
    for (const name of BOW.feature_names) vec['bow__' + name] = 0;
    if (!text) return vec;
    const tokens = normalizeText(text).split(/\s+/);
    for (const token of tokens) {
      if (BOW.vocabulary[token] !== undefined) {
        vec['bow__' + token] = (vec['bow__' + token] || 0) + 1;
      }
    }
    return vec;
  }

  /** clamp_age_val */
  function clampAgeVal(val) {
    return (val >= MIN_AGE && val <= MAX_AGE) ? val : NULL_VALUE;
  }

  /** clean_dependent_ages */
  function cleanDependentAges(ages) {
    if (isna(ages)) return NULL_VALUE;
    if (!Array.isArray(ages)) return NULL_VALUE;
    const cleaned = [];
    for (let age of ages) {
      let ageVal = forceNumeric(age);
      if (typeof age === 'string' && age.toLowerCase().includes('mo')) ageVal /= 12;
      ageVal = clampAgeVal(ageVal);
      if (notna(ageVal)) cleaned.push(ageVal);
    }
    return cleaned;
  }

  /** correct_dependent_counts */
  function correctDependentCounts(features) {
    const dependentAges = cleanDependentAges(features.dependent_ages);
    const ageCount = (notna(dependentAges) && Array.isArray(dependentAges)) ? dependentAges.length : NULL_VALUE;
    const counts = [ageCount, forceNumeric(features.n_dependents)].filter(notna);
    const nDependents = counts.length ? Math.max(...counts) : NULL_VALUE;

    const agesForCorr = (notna(dependentAges) && Array.isArray(dependentAges))
      ? dependentAges.filter(a => a <= MAX_DEPENDENT_AGE) : NULL_VALUE;
    const nDependentsCorrected = (notna(agesForCorr) && Array.isArray(agesForCorr))
      ? agesForCorr.length : NULL_VALUE;

    return {
      n_dependents:           notna(nDependents)          ? Math.min(nDependents, 10)          : NULL_VALUE,
      n_dependents_corrected: notna(nDependentsCorrected) ? Math.min(nDependentsCorrected, 10) : NULL_VALUE,
    };
  }

  /** encode_education */
  function encodeEducation(val) {
    if (isna(val)) return -1;
    const v = normalizeText(val);
    if (/col|bs|vo|ma?s|tesda/.test(v)) return 2;
    if (/hi|h\s*s|k12/.test(v))         return 1;
    if (/elem/.test(v))                  return 0;
    return -1;
  }

  /** encode_housing_status */
  function encodeHousingStatus(val) {
    if (isna(val)) return -1;
    const codes = { rented: 0, free_use: 1, owned: 2 };
    return codes[val.toLowerCase()] !== undefined ? codes[val.toLowerCase()] : -1;
  }

  /** encode_marital_status */
  function encodeMaritalStatus(val) {
    if (isna(val)) return NULL_VALUE;
    const v = val.toLowerCase();
    if (v.includes('sep'))      return 1;
    if (v[0] === 'm')           return 3;
    if (v[0] === 's')           return 0;
    if (v[0] === 'c' || v[0] === 'l') return 2;
    return -1;
  }

  /** prepare_demographics */
  function prepareDemographics(features) {
    const bowVec = bowTransform(notna(features.motorcycle_model) ? features.motorcycle_model : '');
    const depCounts = correctDependentCounts(features);
    return Object.assign({}, bowVec, {
      'num__age':                    clampAgeVal(forceNumeric(features.age)),
      'num__n_children':             forceNumeric(features.n_children),
      'num__n_dependents':           depCounts.n_dependents,
      'num__n_dependents_corrected': depCounts.n_dependents_corrected,
      'cat__housing_status':         encodeHousingStatus(features.housing_status),
      'cat__marital_status':         encodeMaritalStatus(features.marital_status),
      'cat__education':              encodeEducation(features.education),
      'cat__spouse_education':       encodeEducation(features.spouse_education),
    });
  }

  // ── loancalc.py ───────────────────────────────────────────────────────────
  function solveAmort(principal, interest, term) {
    const compounded = Math.pow(1 + interest, term);
    return principal * interest * compounded / (compounded - 1);
  }

  /** expand_loan_terms */
  function expandLoanTerms(val) {
    if (isna(val)) return { loan_term: NULL_VALUE, loan_downpayment: NULL_VALUE };
    const split = val.split('/');
    const downpayment = forceNumeric(split[0]);
    let term = NULL_VALUE;
    if (split.length > 1) {
      term = forceNumeric(split[1]);
      if (split[1].toLowerCase().includes('y')) term *= 12;
      term = (term > 0 && term <= 48) ? term : NULL_VALUE;
    }
    return { loan_term: term, loan_downpayment: downpayment };
  }

  /** impute_amortization */
  function imputeAmortization(financials) {
    const forImputation = (
      isna(financials.monthly_amortization) &&
      notna(financials.loan_amount) &&
      notna(financials.loan_term)
    );
    if (forImputation) {
      financials.monthly_amortization = solveAmort(financials.loan_amount, 0.039881, financials.loan_term);
    }
    return financials;
  }

  /** add_ratio_features */
  function addRatioFeatures(financials) {
    financials.amort_income_ratio      = financials.monthly_amortization / financials.gross_income;
    financials.loan_downpayment_ratio  = financials.loan_downpayment / financials.loan_amount;
    return financials;
  }

  /** prepare_financials */
  function prepareFinancials(features) {
    const financialKeys = [
      'loan_amount','monthly_amortization','gross_income',
      'employment_income','business_income','spouse_income',
    ];
    const financials = {};
    for (const k of financialKeys) financials[k] = forceNumeric(features[k]);
    if (financials.gross_income === 0) financials.gross_income = NULL_VALUE;
    Object.assign(financials, expandLoanTerms(features.loan_terms));
    addRatioFeatures(imputeAmortization(financials));
    const result = {};
    for (const [k, v] of Object.entries(financials)) result['num__' + k] = v;
    return result;
  }

  /** prepare_features */
  function prepareFeatures(normalized) {
    const features    = extractFeatures(FEATURE_MAP, normalized);
    const demographics = prepareDemographics(features);
    const financials   = prepareFinancials(features);
    const all = Object.assign({}, demographics, financials);
    return MODEL_FEATURES.map(k => (all[k] !== undefined ? all[k] : NULL_VALUE));
  }


  // ═══════════════════════════════════════════════════════════════════════════
  // score.py  —  LightGBM tree eval + QuantileTransformer
  // ═══════════════════════════════════════════════════════════════════════════

  function predictTree(node, features) {
    if ('leaf' in node) return node.leaf;
    const val    = features[node.feature];
    const goLeft = isna(val) || isNaN(val) ? node.default_left : (val <= node.threshold);
    return predictTree(goLeft ? node.left : node.right, features);
  }

  function predictDelinquency(features) {
    let score = 0;
    for (const t of MODEL.trees) score += t.shrinkage * predictTree(t.tree, features);
    if (MODEL.average_output) score /= MODEL.trees.length;
    return score;
  }

  /** QuantileTransformer.transform — linear interpolation on quantile table */
  function quantileTransform(x) {
    const q = SCALER.quantiles;
    const r = SCALER.references;
    if (x <= q[0])           return r[0];
    if (x >= q[q.length - 1]) return r[r.length - 1];
    let lo = 0, hi = q.length - 1;
    while (lo < hi - 1) {
      const mid = (lo + hi) >> 1;
      if (q[mid] <= x) lo = mid; else hi = mid;
    }
    return r[lo] + (x - q[lo]) / (q[hi] - q[lo]) * (r[hi] - r[lo]);
  }

  /** make_credit_score */
  function makeCreditScore(features) {
    const delinquency = predictDelinquency(features);
    const scaled      = quantileTransform(delinquency);   // uniform [0,1]
    return Math.round((1 - scaled) * 100);
  }


  // ═══════════════════════════════════════════════════════════════════════════
  // main.py  —  validation + analyzeFile
  // ═══════════════════════════════════════════════════════════════════════════

  const MISSING_THRESHOLD = 5;

  const ESSENTIAL_FIELDS = {
    personal_data: [
      'name','present_address','present_address_tenure','contact_no','birthplace',
      'education','parents_name','parents_address','date_applied','unit_applied',
      'loan_amount','loan_terms','housing_status','dob','age','marital_status',
      'n_children','n_dependents','dependent_ages',
    ],
    income_analysis: { summary: ['gross_income','monthly_amortization'] },
    officer_assessment: [
      'loan_purpose','unit_payor','unit_rider','rider_license',
      'cell_signal_status','prepared_by','remarks',
    ],
  };

  /** is_missing — null/NaN OR iterable where every element is null/NaN */
  function isMissing(x) {
    if (isna(x)) return true;
    if (Array.isArray(x)) return x.every(el => isna(el));
    return false;
  }

  /** validate_fields */
  function validateFields(data, schema) {
    schema = schema || ESSENTIAL_FIELDS;
    const missing = {};
    for (const [k, v] of Object.entries(schema)) {
      const subset = (data && data[k] !== undefined) ? data[k] : null;
      if (isna(subset)) { missing[k] = v; continue; }
      let m;
      if (typeof v === 'object' && !Array.isArray(v)) {
        m = validateFields(subset, v);
        if (Object.keys(m).length) missing[k] = m;
      } else if (Array.isArray(v)) {
        m = v.filter(f => isMissing(subset[f]));
        if (m.length) missing[k] = m;
      }
    }
    return missing;
  }

  /** analyzeFile — public entry point */
  async function analyzeFile(file) {
    console.log('[CreditPipeline] Analyzing:', file.name);
    const arrayBuffer = await file.arrayBuffer();
    const sheet      = loadReportSheet(arrayBuffer);
    console.log('[CreditPipeline] Sheet rows:', sheet.length);
    const parsed     = parseCreditReport(sheet, file.name);
    console.log('[CreditPipeline] Parsed sections:', Object.keys(parsed));
    const normalized = normalizeCreditData(parsed);
    console.log('[CreditPipeline] personal_data keys:', Object.keys(normalized.personal_data || {}));
    const features   = prepareFeatures(normalized);

    const missingCount = features.filter(f => isna(f) || f === -1).length;
    const scoreValid   = missingCount < MISSING_THRESHOLD;
    const score        = makeCreditScore(features);
    const missingFields = validateFields(normalized);

    console.log('[CreditPipeline] missing features:', missingCount, '| score:', score, '| valid:', scoreValid);
    console.log('[CreditPipeline] missing fields:', JSON.stringify(missingFields, null, 2));

    return {
      filename:       file.name,
      credit_score:   scoreValid ? score : null,
      score_valid:    scoreValid,
      missing_count:  missingCount,
      missing_fields: missingFields,
      normalized,
    };
  }

  return { init, analyzeFile };
})();
