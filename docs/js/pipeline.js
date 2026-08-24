/**
 * pipeline.js
 * -----------
 * Pure-JS port of the Python creditfile pipeline.
 * Depends on: SheetJS (XLSX) loaded before this script.
 *
 * Public API (used by ui.js):
 *   CreditPipeline.init()              → Promise<void>   (loads JSON artifacts)
 *   CreditPipeline.analyzeFile(file)   → Promise<result>
 */

const CreditPipeline = (() => {
  'use strict';

  // ── Artifact state ───────────────────────────────────────────────────────
  let MODEL  = null;   // { trees, average_output }
  let SCALER = null;   // { quantiles, references }
  let BOW    = null;   // { vocabulary, feature_names }

  // Derive base URL from the page location, not the script tag
  // Works for http://localhost:8080, GitHub Pages, any subdirectory
  const BASE = window.location.href.replace(/\/[^/]*$/, '').replace(/\/$/, '');

  async function init() {
    console.log('[CreditPipeline] Loading artifacts from:', BASE);
    const [model, scaler, bow] = await Promise.all([
      fetch(`${BASE}/artifacts/model.json`).then(r => { if (!r.ok) throw new Error(`model.json ${r.status}`); return r.json(); }),
      fetch(`${BASE}/artifacts/scaler.json`).then(r => { if (!r.ok) throw new Error(`scaler.json ${r.status}`); return r.json(); }),
      fetch(`${BASE}/artifacts/bow.json`).then(r => { if (!r.ok) throw new Error(`bow.json ${r.status}`); return r.json(); }),
    ]);
    MODEL  = model;
    SCALER = scaler;
    BOW    = bow;
    console.log('[CreditPipeline] Artifacts loaded. Trees:', MODEL.trees.length, 'BoW tokens:', BOW.feature_names.length);
  }

  // ── Utilities ────────────────────────────────────────────────────────────
  const NaN_ = NaN;

  function isna(x) {
    if (x === null || x === undefined) return true;
    if (typeof x === 'number' && isNaN(x)) return true;
    return false;
  }
  function notna(x) { return !isna(x); }

  function forceNumeric(x) {
    if (typeof x === 'number' && !isNaN(x)) return x;
    if (typeof x === 'string') {
      const cleaned = x.replace(/[^0-9.\-]/g, '');
      const n = parseFloat(cleaned);
      return isNaN(n) ? NaN_ : n;
    }
    return NaN_;
  }

  function normalizeText(text) {
    if (!text) return '';
    // Basic Unicode → ASCII approximation
    let s = text.normalize('NFKD').replace(/[\u0300-\u036f]/g, '');
    s = s.toLowerCase();
    s = s.replace(/[.\,\-\/\(\)\s]+/g, ' ').trim();
    s = s.replace(/[^a-z0-9 ]/g, '');
    return s;
  }

  function standardizeField(name) {
    let s = name.trim().toLowerCase();
    s = s.replace(/[\/\s]+/g, '_');
    s = s.replace(/[^a-z0-9_]/g, '');
    return s || '_';
  }

  // ── Excel loading ────────────────────────────────────────────────────────
  function loadReportSheet(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: 'array', raw: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    // Convert to array-of-arrays (like pandas read_excel header=None, dtype=str)
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });
    // Wipe colons and strip
    return raw.map(row =>
      row.map(cell => {
        if (cell === null || cell === undefined) return null;
        let s = String(cell).replace(/:/g, '').trim();
        return s === '' ? null : s;
      })
    );
  }

  // ── Section location ─────────────────────────────────────────────────────
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
      const rowStr = sheet[i].map(c => c || '').join('').toLowerCase();
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

  // ── Parser utils ─────────────────────────────────────────────────────────
  function extractRowwiseKV(rows) {
    const data = {};
    for (const row of rows) {
      const kv = row.filter(x => notna(x));
      if (!kv.length) continue;
      let k = kv[0];
      if (k in data) {
        let suffix = 1;
        while (`${k}_${suffix}` in data) suffix++;
        k = `${k}_${suffix}`;
      }
      data[k] = kv.length > 1 ? kv[1] : null;
    }
    return data;
  }

  function joinSeriesValues(arr, sep = '|') {
    const vals = arr.filter(notna);
    return vals.length ? vals.join(sep) : null;
  }

  // ── Personal data parser ─────────────────────────────────────────────────
  function parsePersonalData(sheet, bounds) {
    const [start, end] = bounds.personal_data;
    let section = sheet.slice(start, end);

    // Adjust index 7+ offset like Python code
    // (Python: if section.iloc[7].any(): adjusted_index[7:] += 1)
    // We handle this by treating row indices as logical, not physical

    const leftCols  = section.map(r => r.slice(0, 15));
    const rightCols = section.map(r => r.slice(15));

    const leftExceptions  = new Set([5, 8, 9, 14, 16, 17]);
    const rightExceptions = new Set([5, 16]);

    const leftFiltered  = leftCols.filter((_, i) => !leftExceptions.has(i));
    const rightFiltered = rightCols.filter((_, i) => !rightExceptions.has(i));

    const personalData = extractRowwiseKV(leftFiltered);
    const rightData    = extractRowwiseKV(rightFiltered);
    for (const [k, v] of Object.entries(rightData)) {
      const key = k in personalData ? `spouse__${k}` : k;
      personalData[key] = v;
    }

    // Exceptions
    const typeRow = section[5] || [];
    personalData.type_of_residence =
      typeRow.slice(8, 11).some(notna) ? 'owned' :
      typeRow.slice(13, 15).some(notna) ? 'rented' :
      typeRow.slice(17, 21).some(notna) ? 'free_use' : null;

    const dobRow = section[9] || [];
    personalData.dob            = dobRow[2]  || null;
    personalData.age            = dobRow[9]  || null;
    personalData.marital_status = dobRow[13] || null;
    personalData.parents_name_2 = (section[14] || [])[3] || null;
    personalData.parents_address_2 = (section[16] || [])[3] || null;
    personalData.spouse__parents_name_2 = (section[16] || [])[19] || null;
    personalData.n_children   = (section[17] || [])[2]  || null;
    personalData.n_dependents = (section[17] || [])[11] || null;

    return personalData;
  }

  // ── Income data parser ───────────────────────────────────────────────────
  function parseIncomeData(sheet, bounds) {
    const [start, end] = bounds.income_data;
    const section = sheet.slice(start, end);

    const leftSection  = section.map(r => r.slice(0, 15));
    const rightSection = section.map(r => r.slice(15));

    // Income sources (left side)
    const incomeSrcBounds = {
      employment:                   [2, 12],
      business:                     [14, 21],
      other_business_or_remittance: [23, 29],
      spouse:                       [31, 40],
    };
    const incomeSources = {};
    for (const [name, [s, e]] of Object.entries(incomeSrcBounds)) {
      incomeSources[name] = extractRowwiseKV(
        leftSection.slice(s, e).map(r => [r[0], r[4]])
      );
    }

    // Income adjudication (right side)
    const rightCleaned = rightSection.map(r => [r[0], r[9]]);
    const adjBounds = {
      income:  [2, 9],
      expense: [11, 31],
      summary: [32, rightCleaned.length],
    };
    const incomeAdjudication = {};
    for (const [name, [s, e]] of Object.entries(adjBounds)) {
      incomeAdjudication[name] = extractRowwiseKV(rightCleaned.slice(s, e));
    }

    return { income_sources: incomeSources, income_adjudication: incomeAdjudication };
  }

  // ── Credit assessment parser ─────────────────────────────────────────────
  function parseCreditAssessment(sheet, bounds) {
    const [start, end] = bounds.credit_assessment;
    const section = sheet.slice(start, end);
    const sub = section.slice(0, 8).map(r => [r[3], r[8]]);
    const data = extractRowwiseKV(sub);
    data.remarks     = (section[9] || [])[7] || null;
    // prepared_by: last non-null value in column 0
    let preparedBy = null;
    for (let i = section.length - 1; i >= 0; i--) {
      if (notna(section[i][0])) { preparedBy = section[i][0]; break; }
    }
    data.prepared_by = preparedBy;
    return data;
  }

  // ── Subtable parser ──────────────────────────────────────────────────────
  function parseSubtable(rows) {
    // Drop empty rows and columns, use first row as headers
    const nonEmptyRows = rows.filter(r => r.some(notna));
    if (nonEmptyRows.length < 2) return {};
    // Find non-empty columns
    const colCount = Math.max(...nonEmptyRows.map(r => r.length));
    const activeCols = [];
    for (let c = 0; c < colCount; c++) {
      if (nonEmptyRows.some(r => notna(r[c]))) activeCols.push(c);
    }
    const headers = activeCols.map(c => nonEmptyRows[0][c] || '');
    const result = {};
    headers.forEach(h => { result[h] = []; });
    for (let r = 1; r < nonEmptyRows.length; r++) {
      activeCols.forEach((c, i) => {
        result[headers[i]].push(nonEmptyRows[r][c] || null);
      });
    }
    return result;
  }

  function parseSubtables(sheet, bounds) {
    const offsets = {
      dependents:           [0, -1],
      character_references: [0, 0],
      client_reputation:    [0, 0],
      other_creditors:      [0, -2],
      client_assets:        [0, 0],
    };
    const result = {};
    for (const [k, [os, oe]] of Object.entries(offsets)) {
      if (!bounds[k]) continue;
      const [s, e] = bounds[k];
      try {
        result[k] = parseSubtable(sheet.slice(s + os, e + oe));
      } catch (_) {}
    }
    return result;
  }

  // ── Full report parser ───────────────────────────────────────────────────
  function parseCreditReport(sheet, filename) {
    const bounds = locateSections(sheet);
    const parsed = { filename, last_modified: new Date().toISOString() };

    try { parsed.personal_data  = parsePersonalData(sheet, bounds); } catch (_) {}
    try { parsed.income_data    = parseIncomeData(sheet, bounds); }   catch (_) {}
    try { parsed.assessment     = parseCreditAssessment(sheet, bounds); } catch (_) {}
    Object.assign(parsed, parseSubtables(sheet, bounds));

    return parsed;
  }

  // ── Normalization ────────────────────────────────────────────────────────
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

    if (parsed.dependents && parsed.dependents['Age']) {
      normalized.dependent_ages = parsed.dependents['Age'];
    }
    return normalized;
  }

  function flattenDict(nested) {
    const result = {};
    for (const [k0, v0] of Object.entries(nested)) {
      if (v0 && typeof v0 === 'object') {
        for (const [k1, v1] of Object.entries(v0)) {
          result[`${k0}__${k1}`] = v1;
        }
      }
    }
    return result;
  }

  function normalizeIncomeSourceDetails(parsed) {
    if (!parsed.income_data || !parsed.income_data.income_sources) return {};
    const data = parsed.income_data.income_sources;

    const fieldCorrections = {
      business__address_of_business:            'business__address',
      business__business_name:                  'business__name',
      business__business_permit_no:             'business__permit_no',
      business__monthly_income:                 'business__monthly_income',
      business__remarks:                        'business__remarks',
      business__route_of_vehicle:               'business__vehicle_route',
      business__years_in_business:              'business__tenure',
      employment__address_of_employer:          'employment__address',
      employment__contact_number_of_employer:   'employment__contact_no',
      employment__length_of_service:            'employment__tenure',
      employment__monthly_net_pay:              'employment__monthly_income',
      employment__monthly_pay:                  'employment__monthly_income',
      employment__name_of_employer:             'employment__name',
      employment__position_employement_status:  'employment__status',
      employment__position_employment_status:   'employment__status',
      employment__previous_employer_address:    'employment__previous_employer',
      employment__remarks:                      'employment__remarks',
      employment__verified_thru_name_contact_no:'employment__verifier',
      employment__verified_thru_name_contact_no_verified:'employment__verifier',
      employment__years_in_operation_of_employer:'employment__employer_tenure',
      other_business_or_remittance__address_of_business:'remittance__address',
      other_business_or_remittance__address_of_business_address_of_sender:'remittance__address',
      other_business_or_remittance__address_of_sender:'remittance__address',
      other_business_or_remittance__business_name:'remittance__name',
      other_business_or_remittance__business_name_name_of_sender:'remittance__name',
      other_business_or_remittance__name_of_sender:'remittance__name',
      other_business_or_remittance__monthly_income:'remittance__monthly_income',
      other_business_or_remittance__monthly_net_income_p:'remittance__monthly_income',
      other_business_or_remittance__monthly_net_income_remittance:'remittance__monthly_income',
      other_business_or_remittance__monthly_net_income_remittance_p:'remittance__monthly_income',
      other_business_or_remittance__nature_of_business:'remittance__industry',
      other_business_or_remittance__nature_of_business_source_of_income_of_sender:'remittance__industry',
      other_business_or_remittance__relationship_of_sender_to_credit_applicant:'remittance__relationship',
      other_business_or_remittance__remarks:    'remittance__remarks',
      other_business_or_remittance__years_in_business:'remittance__tenure',
      other_business_or_remittance__years_in_business_years_of_remittance:'remittance__tenure',
      other_business_or_remittance__years_of_remittance:'remittance__tenure',
      spouse__address_of_employer:              'spouse__employer_address',
      spouse__address_of_business_address_of_sender:'spouse__employer_address',
      spouse__contact_number_of_employer:       'spouse__employer_contact_no',
      spouse__length_of_service:                'spouse__employment_tenure',
      spouse__monthly_net_income_remittance_p:  'spouse__income',
      spouse__monthly_net_pay:                  'spouse__income',
      spouse__monthly_pay:                      'spouse__income',
      spouse__nature_of_business_source_of_income_of_sender:'spouse__income',
      spouse__name_of_employer:                 'spouse__employer_name',
      spouse__position_employement_status:      'spouse__employment_status',
      spouse__position_employment_status:       'spouse__employment_status',
      spouse__previous_employer_address:        'employment__previous_employer',
      spouse__remarks:                          'spouse__remarks',
      spouse__verified_thru_name_contact_no:    'spouse__employment_verifier',
      spouse__years_in_operation_of_employer:   'spouse__employer_tenure',
    };

    const normalized = {};
    for (let [k, v] of Object.entries(flattenDict(data))) {
      k = standardizeField(k);
      if (fieldCorrections[k] && notna(v)) normalized[fieldCorrections[k]] = v;
    }
    return normalized;
  }

  function matchLen(a, b) {
    let i = 0;
    while (i < a.length && i < b.length && a[i] === b[i]) i++;
    return i;
  }

  function extractLongestMatch(query, references, minLen = 0) {
    let max = 0, best = null;
    for (const ref of references) {
      const len = matchLen(query, ref);
      if (len > max) { max = len; best = ref; }
    }
    return max > minLen ? best : null;
  }

  function normalizeIncomeAnalysis(parsed) {
    if (!parsed.income_data || !parsed.income_data.income_adjudication) return {};
    const data = parsed.income_data.income_adjudication;

    const incomeFields = new Set(['applicant','business','others','spouse','total_income']);
    const incomeCorrections = { '1': 'primary', '2': 'secondary' };
    const incomeItems = {};
    for (let [k, v] of Object.entries(data.income || {})) {
      const s = standardizeField(k);
      const mapped = incomeFields.has(s) ? s
        : (incomeCorrections[s] || extractLongestMatch(s, [...incomeFields], 3));
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
      const s = standardizeField(k);
      const mapped = expenseFields.has(s) ? s
        : (expenseCorrections[s] || extractLongestMatch(s, [...expenseFields], 3));
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

  function normalizeOfficerAssessment(parsed) {
    if (!parsed.assessment) return {};
    const data = parsed.assessment;
    const fieldCorrections = {
      'Purpose of loan':                 'loan_purpose',
      'Who will use the unit':           'unit_rider',
      'Who will pay the for the unit':   'unit_payor',
      'User with/without license':       'rider_license',
      'Cellular signal on the area':     'cell_signal_status',
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

  function normalizeCreditData(parsed) {
    return {
      filename:            parsed.filename,
      last_modified:       parsed.last_modified,
      personal_data:       normalizePersonalData(parsed),
      income_source_details: normalizeIncomeSourceDetails(parsed),
      income_analysis:     normalizeIncomeAnalysis(parsed),
      officer_assessment:  normalizeOfficerAssessment(parsed),
    };
  }

  // ── Feature engineering ──────────────────────────────────────────────────
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
        ['gross_income',          'gross_income'],
        ['monthly_amortization',  'monthly_amortization'],
      ],
    },
  };

  function extractFeatures(featureMap, data, features = {}) {
    if (Array.isArray(featureMap)) {
      for (const [fieldName, featureName] of featureMap) {
        features[featureName] = data ? (data[fieldName] !== undefined ? data[fieldName] : NaN_) : NaN_;
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

  function bowTransform(text) {
    const vec = {};
    for (const name of BOW.feature_names) vec[`bow__${name}`] = 0;
    if (!text) return vec;
    const tokens = normalizeText(text).split(/\s+/);
    for (const token of tokens) {
      if (BOW.vocabulary[token] !== undefined) {
        vec[`bow__${token}`] = (vec[`bow__${token}`] || 0) + 1;
      }
    }
    return vec;
  }

  function clampAge(val) {
    return (val >= 0 && val <= 80) ? val : NaN_;
  }

  function cleanDependentAges(ages) {
    if (isna(ages) || !Array.isArray(ages)) return NaN_;
    const cleaned = [];
    for (const age of ages) {
      let val = forceNumeric(age);
      if (typeof age === 'string' && age.toLowerCase().includes('mo')) val /= 12;
      val = clampAge(val);
      if (notna(val)) cleaned.push(val);
    }
    return cleaned;
  }

  function correctDependentCounts(features) {
    const depAges = cleanDependentAges(features.dependent_ages);
    const ageCount = (notna(depAges) && Array.isArray(depAges)) ? depAges.length : NaN_;
    const nDep = forceNumeric(features.n_dependents);
    const counts = [ageCount, nDep].filter(notna);
    const nDependents = counts.length ? Math.max(...counts) : NaN_;

    const correctedAges = (notna(depAges) && Array.isArray(depAges))
      ? depAges.filter(a => a <= 21) : NaN_;
    const nDependentsCorrected = (notna(correctedAges) && Array.isArray(correctedAges))
      ? correctedAges.length : NaN_;

    return {
      n_dependents:           isna(nDependents) ? NaN_ : Math.min(nDependents, 10),
      n_dependents_corrected: isna(nDependentsCorrected) ? NaN_ : Math.min(nDependentsCorrected, 10),
    };
  }

  function encodeEducation(val) {
    if (isna(val)) return -1;
    const s = normalizeText(val);
    if (/col|bs|vo|ma?s|tesda/.test(s)) return 2;
    if (/hi|h\s*s|k12/.test(s)) return 1;
    if (/elem/.test(s)) return 0;
    return -1;
  }

  function encodeHousingStatus(val) {
    if (isna(val)) return -1;
    const codes = { rented: 0, free_use: 1, owned: 2 };
    return codes[val.toLowerCase()] !== undefined ? codes[val.toLowerCase()] : -1;
  }

  function encodeMaritalStatus(val) {
    if (isna(val)) return NaN_;
    const s = val.toLowerCase();
    if (s.includes('sep')) return 1;
    if (s[0] === 'm') return 3;
    if (s[0] === 's') return 0;
    if (s[0] === 'c' || s[0] === 'l') return 2;
    return -1;
  }

  function prepareDemographics(features) {
    const bowVec = bowTransform(features.motorcycle_model);
    const depCounts = correctDependentCounts(features);
    return {
      ...bowVec,
      num__age:                    clampAge(forceNumeric(features.age)),
      num__n_children:             forceNumeric(features.n_children),
      num__n_dependents:           depCounts.n_dependents,
      num__n_dependents_corrected: depCounts.n_dependents_corrected,
      cat__housing_status:         encodeHousingStatus(features.housing_status),
      cat__marital_status:         encodeMaritalStatus(features.marital_status),
      cat__education:              encodeEducation(features.education),
      cat__spouse_education:       encodeEducation(features.spouse_education),
    };
  }

  function expandLoanTerms(val) {
    if (isna(val)) return { loan_term: NaN_, loan_downpayment: NaN_ };
    const split = val.split('/');
    const downpayment = forceNumeric(split[0]);
    let term = NaN_;
    if (split.length > 1) {
      term = forceNumeric(split[1]);
      if (split[1].toLowerCase().includes('y')) term *= 12;
      if (!(term > 0 && term <= 48)) term = NaN_;
    }
    return { loan_term: term, loan_downpayment: downpayment };
  }

  function solveAmort(principal, interest, term) {
    const c = Math.pow(1 + interest, term);
    return principal * interest * c / (c - 1);
  }

  function impute_amortization(f) {
    if (isna(f.monthly_amortization) && notna(f.loan_amount) && notna(f.loan_term)) {
      f.monthly_amortization = solveAmort(f.loan_amount, 0.039881, f.loan_term);
    }
    return f;
  }

  function prepareFinancials(features) {
    const keys = ['loan_amount','monthly_amortization','gross_income',
                  'employment_income','business_income','spouse_income'];
    const fin = {};
    for (const k of keys) fin[k] = forceNumeric(features[k]);
    if (fin.gross_income === 0) fin.gross_income = NaN_;
    const { loan_term, loan_downpayment } = expandLoanTerms(features.loan_terms);
    fin.loan_term = loan_term;
    fin.loan_downpayment = loan_downpayment;
    impute_amortization(fin);
    fin.amort_income_ratio       = fin.monthly_amortization / fin.gross_income;
    fin.loan_downpayment_ratio   = fin.loan_downpayment / fin.loan_amount;
    const result = {};
    for (const [k, v] of Object.entries(fin)) result[`num__${k}`] = v;
    return result;
  }

  function prepareFeatures(normalized) {
    const raw = extractFeatures(FEATURE_MAP, normalized);
    const demo = prepareDemographics(raw);
    const fin  = prepareFinancials(raw);
    const all  = { ...demo, ...fin };
    return MODEL_FEATURES.map(k => {
      const v = all[k];
      return (v === undefined || v === null) ? NaN_ : v;
    });
  }

  // ── LightGBM inference ───────────────────────────────────────────────────
  function predictTree(node, features) {
    if ('leaf' in node) return node.leaf;
    const val = features[node.feature];
    const goLeft = isna(val) || isNaN(val)
      ? node.default_left
      : (val <= node.threshold);
    return predictTree(goLeft ? node.left : node.right, features);
  }

  function predictRawScore(features) {
    let score = 0;
    for (const t of MODEL.trees) {
      score += t.shrinkage * predictTree(t.tree, features);
    }
    if (MODEL.average_output) score /= MODEL.trees.length;
    return score;
  }

  // ── QuantileTransformer ──────────────────────────────────────────────────
  function quantileTransform(x) {
    // Interpolate: find where x falls in quantiles, return corresponding reference
    const q = SCALER.quantiles;
    const r = SCALER.references;
    if (x <= q[0])  return r[0];
    if (x >= q[q.length - 1]) return r[r.length - 1];
    // Binary search
    let lo = 0, hi = q.length - 1;
    while (lo < hi - 1) {
      const mid = (lo + hi) >> 1;
      if (q[mid] <= x) lo = mid; else hi = mid;
    }
    const t = (x - q[lo]) / (q[hi] - q[lo]);
    return r[lo] + t * (r[hi] - r[lo]);
  }

  function makeCreditScore(features) {
    const raw = predictRawScore(features);
    const scaled = quantileTransform(raw);   // uniform [0,1]
    return Math.round((1 - scaled) * 100);
  }

  // ── Validation ───────────────────────────────────────────────────────────
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

  function validateFields(data, schema = ESSENTIAL_FIELDS) {
    const missing = {};
    for (const [k, v] of Object.entries(schema)) {
      const subset = data ? data[k] : null;
      if (isna(subset) || subset === undefined) { missing[k] = v; continue; }
      let m;
      if (typeof v === 'object' && !Array.isArray(v)) {
        m = validateFields(subset, v);
      } else if (Array.isArray(v)) {
        m = v.filter(f => isna(subset[f]));
      }
      if (m && (Array.isArray(m) ? m.length : Object.keys(m).length)) missing[k] = m;
    }
    return missing;
  }

  // ── Public analyzeFile ───────────────────────────────────────────────────
  async function analyzeFile(file) {
    console.log('[CreditPipeline] Analyzing:', file.name);
    const arrayBuffer = await file.arrayBuffer();
    const sheet   = loadReportSheet(arrayBuffer);
    console.log('[CreditPipeline] Sheet rows:', sheet.length);
    const parsed  = parseCreditReport(sheet, file.name);
    console.log('[CreditPipeline] Parsed sections:', Object.keys(parsed));
    const normalized = normalizeCreditData(parsed);
    console.log('[CreditPipeline] Normalized personal_data keys:', Object.keys(normalized.personal_data || {}));
    const features   = prepareFeatures(normalized);
    console.log('[CreditPipeline] Features:', features);

    const MISSING_THRESHOLD = 5;
    const missingCount = features.filter(f => isna(f) || f === -1).length;
    const scoreValid   = missingCount < MISSING_THRESHOLD;
    const score        = scoreValid ? makeCreditScore(features) : null;
    const missingFields = validateFields(normalized);

    return {
      filename:      file.name,
      credit_score:  score,
      score_valid:   scoreValid,
      missing_count: missingCount,
      missing_fields: missingFields,
      normalized,
    };
  }

  return { init, analyzeFile };
})();
