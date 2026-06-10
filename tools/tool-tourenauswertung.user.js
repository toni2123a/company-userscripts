// ==UserScript==
// @name         DPD Dispatcher – Tourenauswertung
// @namespace    bodo.dpd.custom
// @version      3.1.3
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-tourenauswertung.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-tourenauswertung.user.js
// @description  Tourenauswertung mit Datum von/bis, Systempartner/Touren, Fahrername, Stopps/Paketen, automatischen Fällen und Zustellhindernissen inkl. Klick-Details.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const MODULE_ID = 'touren-auswertung';
  const NS = 'spx-';
  const STORE_DB = 'fvpr_db';
  const STORE_TOURMAP = 'tourMap';

  const TOURMAP_CACHE_MS = 10 * 60 * 1000;
  const SUMMARY_CACHE_MS = 60 * 1000;
  const DETAIL_CACHE_MS = 60 * 1000;

  const PAGE_SIZE = 500;
  const HARD_MAX_PAGES = 300;

  const moduleDef = {
    id: MODULE_ID,
    label: 'Tourenauswertung',
    run: () => startModuleOnce()
  };

  function registerTourenauswertungModule() {
    try {
      if (window.TM && typeof window.TM.register === 'function') {
        window.TM.register(moduleDef);
        return true;
      }

      window.__tmQueue = window.__tmQueue || [];
      if (!window.__tmQueue.some(m => m && m.id === MODULE_ID)) {
        window.__tmQueue.push(moduleDef);
      }
    } catch (e) {
      console.warn('Tourenauswertung konnte noch nicht registriert werden.', e);
    }
    return false;
  }

  registerTourenauswertungModule();
  setTimeout(registerTourenauswertungModule, 500);
  setTimeout(registerTourenauswertungModule, 1500);
  window.addEventListener('load', registerTourenauswertungModule, { once: true });

  let started = false;
  let isBusy = false;
  let lastOkRequest = null;

  let tourPartnerMap = new Map();
  let tourDriverMap = new Map();
  let tourPartnerLookupCache = new Map();
  let tourDriverLookupCache = new Map();
  let tourMapLoadedAt = 0;

  const summaryCache = new Map();
  const detailCache = new Map();

  const collator = new Intl.Collator('de', { numeric: true, sensitivity: 'base' });


  const NEW_DIRECT_COLUMNS_FROM = '2026-05-06';
  const EXTRA_COLUMNS_STORAGE_KEY = NS + 'extra-detail-columns-v1';

  const FIELD_CATALOG = [
    ['id', 'ID', ['id']],
    ['date', 'Auftragsdatum', ['date']],
    ['scan_time', 'Scan-Zeit', ['scan_time', 'scanTime']],
    ['delivery_status', 'Status Zustellung', ['delivery_status', 'deliveryStatus']],
    ['pickup_status', 'Status Abholung', ['pickup_status', 'pickupStatus']],
    ['planned_real_distance_deviation_meters', 'Distanzabweichung', ['planned_real_distance_deviation_meters', 'plannedRealDistanceDeviationMeters']],
    ['priority', 'Prio', ['priority']],
    ['tour', 'Tour', ['tour', 'round']],
    ['parcel_number', 'Paketnummer', ['parcel_number', 'parcelNumber']],
    ['name', 'Name', ['name']],
    ['street', 'Straße', ['street']],
    ['houseNo', 'Hausnummer', ['houseNo', 'houseno', 'houseNumber']],
    ['city', 'Stadt', ['city']],
    ['country_postal', 'Land und Postleitzahl', ['country', 'postal_code', 'postalCode']],
    ['predict_timeframe', 'Predict-Zeitfenster', ['timeframe_from_predict', 'timeframe_to_predict']],
    ['difference', 'ETA Differenz', ['difference']],
    ['stop', 'Stopp', ['stop']],
    ['order_type', 'Auftragstyp', ['order_type', 'orderType']],
    ['service_code', 'Service-Code', ['service_code', 'serviceCode']],
    ['pickup_type', 'Abholtyp', ['pickup_type', 'pickupType']],
    ['customer_number', 'Kundennummer', ['customer_number', 'customerNumber']],
    ['customer_name', 'Kundenname', ['customer_name', 'customerName']],
    ['changed_consignee', 'Geänderter Empfänger', ['changed_consignee', 'changedConsignee']],
    ['name2', 'Name 2', ['name2']],
    ['phone', 'Telefon', ['phone']],
    ['scanned_planned_parcels', 'Gescannte / Geplante Pakete', ['completed_parcel', 'estimated_parcels', 'parcels']],
    ['time_critical', 'zeitkritisch', ['time_critical', 'timeCritical']],
    ['depot', 'Depot', ['depot']],
    ['standard_timeframe', 'Standard-Zeitfenster', ['default_time_from', 'default_time_to']],
    ['new_deliverydate', 'Neues Lieferdatum', ['new_deliverydate', 'newDeliverydate']],
    ['delay', 'Verspätung', ['delay']],
    ['timeframe_type', 'Zeitfenster-Typ', ['timeframe_type', 'timeframeType']],
    ['scan_date', 'Scan-Datum', ['scan_date', 'scanDate']],
    ['cod_cop', 'COD / COP', ['cash_amount', 'cash_currency']],
    ['receipt_id', 'Quittung ID', ['receipt_id', 'receiptId']],
    ['loading_unloading', 'Laden / Entladen', ['loading', 'unloading']],
    ['waiting', 'Warten', ['waiting']],
    ['additional_code', 'Zusatzcode', ['additional_code', 'additionalCode', 'additionalCodes']],
    ['yellow_card', 'PIK-Nummer', ['yellow_card', 'yellowCard']],
    ['freetext_dc', 'Freetext Dc', ['freetext_dc', 'freetextDc']],
    ['contact_name', 'Kontaktperson', ['contact_name', 'contactName']],
    ['pudo_id', 'Paketshop ID', ['pudo_id', 'pudoId']],
    ['in_eta', 'in ETA', ['in_eta', 'inEta']],
    ['hazardous_goods', 'GG', ['hazardous_goods', 'hazardousGoods']],
    ['tour_area', 'DPD PLZ', ['tour_area', 'tourArea']],
    ['delis_id', 'Delis ID', ['delis_id', 'delisId']],
    ['eta', 'ETA', ['eta']],
    ['ga_1', 'Zeitfenster 1 Z / A', ['good_acceptance_time_from_1', 'good_acceptance_time_to_1']],
    ['ga_2', 'Zeitfenster 2 Z / A', ['good_acceptance_time_from_2', 'good_acceptance_time_to_2']],
    ['extra', 'Präferenz', ['extra']],
    ['service_category', 'Service Kategorie', ['service_category', 'serviceCategory']],
    ['eta_scantime', 'ETA-Scantime', ['eta_scantime', 'etaScantime']],
    ['proposed_tour', 'Ermittelte Tour', ['proposed_tour', 'proposedTour']],
    ['signer', 'Unterzeichner', ['signer']],
    ['note', 'Interne Notiz', ['note']],
    ['depot_tour_change_counter', 'Tour Änderungen', ['depot_tour_change_counter', 'depotTourChangeCounter']],
    ['service_type', 'Service', ['service_type', 'serviceType']],
    ['real_coordinate_lat', 'Real Coordinate Lat', ['real_coordinate_lat', 'realCoordinateLat']],
    ['real_coordinate_long', 'Real Coordinate Long', ['real_coordinate_long', 'realCoordinateLong']],
    ['planned_coordinate_lat', 'Planned Coordinate Lat', ['planned_coordinate_lat', 'plannedCoordinateLat']],
    ['planned_coordinate_long', 'Planned Coordinate Long', ['planned_coordinate_long', 'plannedCoordinateLong']],
    ['insert_user', 'Ersteller', ['insert_user', 'insertUser']],
    ['modify_date', 'Änderungsdatum', ['modify_date', 'modifyDate']],
    ['modify_user', 'Verändert durch', ['modify_user', 'modifyUser']],
    ['as_code', 'AS code', ['as_code', 'asCode']],
    ['elements', 'Service Elemente', ['elements']],
    ['swap', 'Austausch', ['swap']],
    ['id_check', 'ID Check', ['id_check', 'idCheck']],
    ['department_delivery', 'Abteilung Zustellung', ['department_delivery', 'departmentDelivery']],
    ['delivered_time', 'Zustellzeit', ['delivered_time', 'deliveredTime']],
    ['free_comment', 'Freier Kommentar', ['free_comment', 'freeComment']],
    ['event_id', 'Event ID', ['event_id', 'eventId']],
    ['source_system_id', 'Quellsystem ID', ['source_system_id', 'sourceSystemId']],
    ['complaint_id', 'Beschwerde ID', ['complaint_id', 'complaintId']],
    ['complaint_status', 'Beschwerde Status', ['complaint_status', 'complaintStatus']],
    ['planned_pallets', 'Anzahl der geplanten Paletten', ['planned_pallets', 'plannedPallets']],
    ['problem_reason', 'PROBLEM Check Grund', ['problem_reason', 'problemReason']],
    ['problem_comment', 'PROBLEM Check Kommentar', ['problem_comment', 'problemComment']],
    ['courier_name', 'Zustellername', ['courier_name', 'courierName']],
    ['subcontractor', 'Systempartner', ['subcontractor', 'subcontractor_name', 'subcontractorName', 'systempartner', 'systemPartner']]
  ].map(([key, label, paths]) => ({ key: 'extra_' + key, rawKey: key, label, paths }));

  const BASE_DETAIL_COLUMN_KEYS = new Set([
    'parcel', 'systempartner', 'driver', 'tour', 'serviceCode', 'type', 'address', 'packages',
    'parcelsList', 'status', 'reason', 'time', 'additionalCode'
  ]);

  const state = {
    summaryRows: [],
    standText: '',
    detectedDateParam: '',
    detectedDateValue: '',
    detectedDateFrom: '',
    detectedDateTo: '',
    selectedExtraColumns: [],
    modal: {
      type: '',
      title: '',
      rows: [],
      columns: [],
      totals: null,
      sortCol: -1,
      sortDir: 'asc'
    }
  };

  const STATUS_LABEL_DE = new Map([
    ['DELIVERED', 'ZUGESTELLT'],
    ['DELIVERED TO PUDO', 'ZUGESTELLT PUDO'],
    ['DELIVERED_TO_PUDO', 'ZUGESTELLT PUDO'],
    ['ZUGESTELLT', 'ZUGESTELLT'],
    ['PICKED UP', 'ABGEHOLT'],
    ['PICKED_UP', 'ABGEHOLT'],
    ['ABGEHOLT', 'ABGEHOLT'],
    ['COMPLETED', 'ABGEHOLT'],
    ['DELIVERY PROBLEM', 'ZUSTELLUNG PROBLEM'],
    ['DELIVERY_PROBLEM', 'ZUSTELLUNG PROBLEM'],
    ['DELIVERY ISSUE', 'ZUSTELLUNG PROBLEM'],
    ['ZUSTELLUNG PROBLEM', 'ZUSTELLUNG PROBLEM'],
    ['PROBLEM', 'PROBLEM'],
    ['PROBLEM_ROLLOVER', 'PROBLEM'],
    ['NOT DELIVERED', 'NICHT ZUGESTELLT'],
    ['NICHT ZUGESTELLT', 'NICHT ZUGESTELLT'],
    ['DELIVERY CANCELLED AUTOMATICALLY', 'ZUSTELLUNG STORNIERT (AUTOMATISCH)'],
    ['DELIVERY_CANCELLED_AUTOMATICALLY', 'ZUSTELLUNG STORNIERT (AUTOMATISCH)'],
    ['CANCELLED_AUTOMATICALLY', 'STORNIERT (AUTOMATISCH)'],
    ['PICKUP CANCELLED AUTOMATICALLY', 'ABHOLUNG STORNIERT (AUTOMATISCH)'],
    ['PICKUP_CANCELLED_AUTOMATICALLY', 'ABHOLUNG STORNIERT (AUTOMATISCH)'],
    ['CANCELLED', 'STORNIERT'],
    ['CANCELLED_ROLLOVER', 'STORNIERT']
  ]);

  const DELIVERY_CANCEL_MATCHES = [
    'DELIVERY CANCELLED AUTOMATICALLY',
    'DELIVERY_CANCELLED_AUTOMATICALLY',
    'CANCELLED_AUTOMATICALLY',
    'ZUSTELLUNG STORNIERT (AUTOMATISCH)',
    'AUTO CANCEL',
    'AUTOMATICALLY CANCELLED'
  ];

  const PICKUP_CANCEL_MATCHES = [
    'PICKUP CANCELLED AUTOMATICALLY',
    'PICKUP_CANCELLED_AUTOMATICALLY',
    'CANCELLED_AUTOMATICALLY',
    'ABHOLUNG STORNIERT (AUTOMATISCH)',
    'AUTO CANCEL',
    'AUTOMATICALLY CANCELLED'
  ];

  const DELIVERY_HINDRANCE_MATCHES = [
    'DELIVERY PROBLEM',
    'DELIVERY_PROBLEM',
    'DELIVERY ISSUE',
    'NOT DELIVERED',
    'NICHT ZUGESTELLT',
    'ZUSTELLUNG PROBLEM',
    'DELIVERY OBSTACLE',
    'ZUSTELLHINDERNIS',
    'CONSIGNEE ABSENT',
    'EMPFÄNGER NICHT ANGETROFFEN',
    'REFUSED',
    'ANNAHME VERWEIGERT',
    'NO ACCESS',
    'KEIN ZUTRITT',
    'ADDRESS INCORRECT',
    'ADDRESS UNKNOWN',
    'ADRESSFEHLER',
    'UNCLAIMED',
    'FAILED DELIVERY',
    'UNDELIVERABLE',
    'RETURNED',
    'RETOUR',
    'CLOSED',
    'GESCHLOSSEN'
  ];

  const esc = s => String(s ?? '').replace(/[&<>"']/g, ch => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
  }[ch]));

  const norm = s => String(s || '').replace(/\s+/g, ' ').trim();

  function toIsoDateLocal(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function parseDateLike(val) {
    const s = String(val || '').trim();
    if (!s) return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const m1 = s.match(/^(\d{4}-\d{2}-\d{2})T/);
    if (m1) return m1[1];
    const m2 = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
    if (m2) return `${m2[3]}-${m2[2]}-${m2[1]}`;
    return '';
  }

  function formatDateDE(iso) {
    const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return iso || '—';
    return `${m[3]}.${m[2]}.${m[1]}`;
  }

  function detectDateParamsFromUrl(u) {
    if (!u) return { from: '', to: '', key: '' };

    const dateFrom = parseDateLike(u.searchParams.get('dateFrom'));
    const dateTo = parseDateLike(u.searchParams.get('dateTo'));

    if (dateFrom || dateTo) {
      return { from: dateFrom, to: dateTo || dateFrom, key: 'dateFrom/dateTo' };
    }

    const preferred = [
      'date', 'deliveryDate', 'pickupDate', 'tourDate', 'dispatchDate', 'plannedDate',
      'executionDate', 'serviceDate', 'day'
    ];

    for (const k of preferred) {
      const v = u.searchParams.get(k);
      const iso = parseDateLike(v);
      if (iso) return { from: iso, to: iso, key: k };
    }

    for (const [k, v] of u.searchParams.entries()) {
      if (!/date|day/i.test(k)) continue;
      const iso = parseDateLike(v);
      if (iso) return { from: iso, to: iso, key: k };
    }

    return { from: '', to: '', key: '' };
  }

  function getDefaultDateFrom() {
    return state.detectedDateFrom || state.detectedDateValue || toIsoDateLocal(new Date());
  }

  function getDefaultDateTo() {
    return state.detectedDateTo || state.detectedDateValue || toIsoDateLocal(new Date());
  }

  function getSelectedDateFrom() {
    const el = document.getElementById(NS + 'date-from');
    return el ? String(el.value || '') : '';
  }

  function getSelectedDateTo() {
    const el = document.getElementById(NS + 'date-to');
    return el ? String(el.value || '') : '';
  }

  function ensureValidDateRange() {
    const fromEl = document.getElementById(NS + 'date-from');
    const toEl = document.getElementById(NS + 'date-to');
    if (!fromEl || !toEl) return;
    if (fromEl.value && toEl.value && fromEl.value > toEl.value) {
      toEl.value = fromEl.value;
    }
  }

  function getSummaryCacheKey() {
    return `${getSelectedDateFrom()}__${getSelectedDateTo()}`;
  }

  function tourKey(t) {
    let s = String(t || '').trim();
    if (!s) return '';
    s = s.replace(/\s+/g, '');
    s = s.replace(',', '.');

    const numMatch = s.match(/(\d+(?:\.\d+)?)/);
    if (numMatch) {
      const n = Number(numMatch[1]);
      if (!Number.isNaN(n)) return String(Math.trunc(n));
    }

    s = s.replace(/[^\dA-Za-z]/g, '');
    if (!s) return '';

    const onlyDigits = s.match(/\d+/);
    if (onlyDigits) {
      const n = Number(onlyDigits[0]);
      if (!Number.isNaN(n)) return String(Math.trunc(n));
    }

    return s.toUpperCase();
  }

  function normalizeStatusForMatch(s) {
    return norm(String(s || ''))
      .toUpperCase()
      .replace(/\u00A0/g, ' ')
      .replace(/[_/\\|]+/g, ' ')
      .replace(/[-]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function containsAnyToken(text, tokens) {
    const x = normalizeStatusForMatch(text);
    return tokens.some(t => x.includes(normalizeStatusForMatch(t)));
  }

  function statusDe(s) {
    const n = normalizeStatusForMatch(s);
    return STATUS_LABEL_DE.get(n) || norm(s || '—') || '—';
  }

  function isCanceledDeliveryStatus(status, row) {
    return containsAnyToken(status, DELIVERY_CANCEL_MATCHES) || automaticHint(row, 'DELIVERY');
  }

  function isCanceledPickupStatus(status, row) {
    return containsAnyToken(status, PICKUP_CANCEL_MATCHES) || automaticHint(row, 'PICKUP');
  }

  function isDeliveryHindranceStatus(status, row) {
    const x = normalizeStatusForMatch(status);
    if (!x) return false;

    if (isCanceledDeliveryStatus(x, row)) return false;

    if (x === 'DELIVERED' || x === 'ZUGESTELLT' || x === 'DELIVERED TO PUDO' || x === 'DELIVERED_TO_PUDO') {
      return false;
    }

    if (x === 'DELIVERY PROBLEM' || x === 'DELIVERY_PROBLEM') return true;
    if (x === 'NOT DELIVERED' || x === 'NICHT ZUGESTELLT') return true;

    const hindrance = additionalCodeOf(row);
    const reason = norm(row?.problem_reason || row?.problemReason || row?.problem_reason_text || '');

    if (containsAnyToken(x, DELIVERY_HINDRANCE_MATCHES)) return true;
    if (hindrance) return true;
    if (reason) return true;

    return false;
  }

  function automaticHint(row, type) {
    const text = flattenPrimitiveStrings(row, 4).join(' | ').toUpperCase();
    if (!text) return false;
    const auto = text.includes('AUTOMAT') || text.includes('AUTO CANCEL') || text.includes('AUTOMATICALLY');
    const cancel = text.includes('STORNI') || text.includes('CANCEL');
    if (!(auto && cancel)) return false;
    if (type === 'PICKUP') return text.includes('PICKUP') || text.includes('ABHOL') || !text.includes('DELIVERY');
    return text.includes('DELIVERY') || text.includes('ZUSTELL') || !text.includes('PICKUP');
  }

  function flattenPrimitiveStrings(obj, maxDepth = 4) {
    const out = [];
    const seen = new WeakSet();

    function walk(v, depth) {
      if (v == null || depth > maxDepth) return;
      if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') {
        const s = norm(String(v));
        if (s) out.push(s);
        return;
      }
      if (Array.isArray(v)) {
        v.forEach(x => walk(x, depth + 1));
        return;
      }
      if (typeof v === 'object') {
        if (seen.has(v)) return;
        seen.add(v);
        Object.values(v).forEach(val => walk(val, depth + 1));
      }
    }

    walk(obj, 0);
    return out;
  }

  function getObjectPath(obj, path) {
    try {
      return path.split('.').reduce((acc, part) => acc?.[part], obj);
    } catch {
      return undefined;
    }
  }

  function parsePossibleDateFromRows(rows) {
    const candidatePaths = [
      'date',
      'tourDate',
      'deliveryDate',
      'pickupDate',
      'serviceDate',
      'plannedDate',
      'dispatchDate',
      'executionDate',
      'day',
      'createdDate',
      'createdAt',
      'delivery.date',
      'pickup.date',
      'tour.date'
    ];
    for (const row of rows) {
      for (const p of candidatePaths) {
        const v = getObjectPath(row, p);
        const iso = parseDateLike(v);
        if (iso) return iso;
      }
    }
    return '';
  }

  function statusOf(r) {
    const directCandidates = [
      r?.delivery_status,
      r?.pickup_status,
      r?.deliveryStatus,
      r?.pickupStatus,
      r?.statusDisplay,
      r?.statusLabel,
      r?.statusDescription,
      r?.statusName,
      r?.statusText,
      r?.parcelStatus,
      typeof r?.status === 'string' ? r.status : '',
      typeof r?.stopStatus === 'string' ? r.stopStatus : '',
      typeof r?.orderStatus === 'string' ? r.orderStatus : '',
      r?.processingStatus,
      r?.scanStatus,
      r?.tourStatus,
      r?.reasonText,
      r?.statusReason,
      r?.problemReason,
      r?.problem_reason
    ];

    for (const c of directCandidates) {
      const s = norm(c);
      if (s) return s;
    }

    const flat = flattenPrimitiveStrings(r, 3);
    const hit = flat.find(x =>
      /delivered|zugestellt|picked up|abgeholt|problem|issue|not delivered|nicht zugestellt|cancel/i.test(x)
    );
    return hit || '';
  }

  function orderTypeOf(r) {
    const candidates = [
      r.order_type,
      r.orderType,
      r.type,
      r.stopType,
      r.jobType,
      r.missionType,
      r.serviceType,
      r.shipmentType,
      r.deliveryOrPickup,
      r.kind,
      r.category,
      r?.order?.type,
      r?.stop?.type,
      r?.tour?.type
    ];

    for (const c of candidates) {
      const s = String(c ?? '').trim().toUpperCase();
      if (!s) continue;
      if (s.includes('DELIVERY') || s.includes('ZUSTELL')) return 'DELIVERY';
      if (s.includes('PICKUP') || s.includes('ABHOL')) return 'PICKUP';
    }

    const text = flattenPrimitiveStrings(r, 3).join(' | ').toUpperCase();
    if (text.includes('PICKUP') || text.includes('ABHOL')) return 'PICKUP';
    return 'DELIVERY';
  }

  function extractTour(r) {
    const candidates = [
      r.tour,
      r.round,
      r.route,
      r.tourNo,
      r.tourNumber,
      r.routeNumber,
      r.routeNo,
      r.roundNo,
      r.tripNo,
      r.tripNumber,
      r?.tour?.number,
      r?.tour?.tourNumber,
      r?.tour?.routeNumber,
      r?.tour?.name,
      r?.route?.number,
      r?.route?.name,
      r?.stop?.tour,
      r?.stop?.tourNumber
    ];

    for (const c of candidates) {
      const s = String(c ?? '').trim();
      if (s) return s;
    }
    return '';
  }


  function dataDateOfRow(r) {
    return parseDateLike(
      r?.date || r?.scan_date || r?.scanDate || r?.deliveryDate || r?.pickupDate ||
      r?.tourDate || r?.plannedDate || r?.executionDate || r?.createdDate || r?.modify_date || ''
    );
  }

  function useDirectColumnsForRow(r) {
    const rowDate = dataDateOfRow(r) || getSelectedDateFrom() || getDefaultDateFrom();
    return !!rowDate && rowDate >= NEW_DIRECT_COLUMNS_FROM;
  }

  function getFirstValueByPaths(obj, paths) {
    for (const p of paths || []) {
      const v = getObjectPath(obj, p);
      if (v == null || v === '') continue;
      if (Array.isArray(v) && !v.length) continue;
      return v;
    }
    return '';
  }

  function formatGenericValue(v) {
    if (v == null || v === '') return '—';
    if (Array.isArray(v)) return v.map(formatGenericValue).filter(x => x !== '—').join(', ') || '—';
    if (typeof v === 'object') {
      try { return JSON.stringify(v); } catch { return String(v); }
    }
    return String(v).trim() || '—';
  }

  function combinedFieldValue(r, def) {
    const values = [];
    for (const p of def.paths || []) {
      const v = getObjectPath(r, p);
      const txt = formatGenericValue(v);
      if (txt !== '—') values.push(txt);
    }
    return [...new Set(values)].join(' / ') || '—';
  }

  function extraFieldMapOf(rawRow) {
    const out = Object.create(null);
    for (const def of FIELD_CATALOG) out[def.key] = combinedFieldValue(rawRow, def);
    return out;
  }

  function getSelectedExtraColumns() {
    const valid = new Set(FIELD_CATALOG.map(f => f.key));
    let arr = state.selectedExtraColumns;
    if (!Array.isArray(arr) || !arr.length) {
      try { arr = JSON.parse(localStorage.getItem(EXTRA_COLUMNS_STORAGE_KEY) || '[]'); } catch { arr = []; }
    }
    return arr.filter(k => valid.has(k));
  }

  function setSelectedExtraColumns(keys) {
    const valid = new Set(FIELD_CATALOG.map(f => f.key));
    state.selectedExtraColumns = Array.from(new Set((keys || []).filter(k => valid.has(k))));
    try { localStorage.setItem(EXTRA_COLUMNS_STORAGE_KEY, JSON.stringify(state.selectedExtraColumns)); } catch {}
  }

  function extraColumnsForDetails() {
    const selected = getSelectedExtraColumns();
    return FIELD_CATALOG
      .filter(f => selected.includes(f.key))
      .filter(f => !BASE_DETAIL_COLUMN_KEYS.has(f.key))
      .map(f => ({ key: f.key, label: f.label }));
  }

  function withExtraFields(row, normalizedRow) {
    return Object.assign(row, normalizedRow?.__extraFields || {});
  }

  function rawPartnerOfRow(r) {
    const candidates = [
      r.systemPartner,
      r.systempartner,
      r.systemPartnerName,
      r.partner,
      r.partnerName,
      r.servicePartner,
      r.servicePartnerName,
      r.carrierName,
      r.transportPartnerName,
      r.contractorName,
      r.subcontractor,
      r.subcontractorName,
      r.subcontractor_name,
      r?.tour?.partnerName,
      r?.tour?.systemPartner,
      r?.tour?.servicePartner,
      r?.route?.partnerName
    ];

    for (const c of candidates) {
      const s = norm(c);
      if (s) return s;
    }
    return '';
  }

  function extractNumericCode(val) {
    if (val == null) return '';
    const s = String(val).trim();
    if (!s) return '';

    const m = s.match(/\d{1,3}/);
    return m ? m[0].padStart(3, '0') : '';
  }

  function additionalCodeRawOf(r) {
    // Exakt dieselbe Quelle wie im funktionierenden Prio/Express-Script:
    // pickup-delivery -> additionalCodes
    if (Array.isArray(r?.additionalCodes) && r.additionalCodes.length) {
      const cleaned = r.additionalCodes
        .map(extractNumericCode)
        .filter(Boolean);

      if (cleaned.length) return cleaned.join(', ');
    }

    return '';
  }

  function additionalCodeOf(r) {
    return additionalCodeRawOf(r) || '';
  }

  function serviceCodeOf(r) {
    if (!r) return '';

    const set = new Set();

    const addFromVal = v => {
      if (v == null) return;
      String(v)
        .split(/[^\dA-Za-z]+/)
        .map(s => s.trim())
        .filter(Boolean)
        .forEach(code => set.add(code));
    };

    const addFromArr = arr => {
      if (!Array.isArray(arr)) return;
      arr.forEach(addFromVal);
    };

    addFromVal(r.serviceCode);
    addFromVal(r.servicecode);
    addFromVal(r.service_code);
    addFromArr(r.serviceCodes);

    if (r.service && typeof r.service === 'object') {
      addFromVal(r.service.code);
      addFromVal(r.service.serviceCode);
      addFromVal(r.service.id);
      addFromArr(r.service.serviceCodes);
    }

    if (r.product && typeof r.product === 'object') {
      addFromVal(r.product.serviceCode);
      addFromVal(r.product.code);
      addFromVal(r.product.id);
      addFromArr(r.product.serviceCodes);
    }

    const arr = Array.from(set);
    arr.sort((a, b) => collator.compare(a, b));
    return arr.join(' ');
  }

  function serviceCodeKind(val) {
    const s = normalizeStatusForMatch(val);
    if (!s) return '';
    if (s.includes('EXPRESS')) return 'EXPRESS';
    if (s.includes('PRIO')) return 'PRIO';
    return '';
  }

  function cleanAdditionalCode(val) {
    if (val == null) return '';

    const parts = String(val)
      .split(',')
      .map(s => extractNumericCode(s))
      .filter(Boolean);

    return parts.length ? parts[0] : '';
  }

  function hindranceItemCount(r) {
    if (!r) return 1;

    const psnCount = Array.isArray(r.__parcelList) ? r.__parcelList.filter(Boolean).length : 0;
    if (psnCount > 0) return psnCount;

    const pkg = Number(r.__pkgCount || 0);
    if (Number.isFinite(pkg) && pkg > 0) return Math.trunc(pkg);

    return 1;
  }

  function hasUsableAdditionalCode(r) {
    return !!cleanAdditionalCode(r?.__additionalCode || '');
  }

  function driverNameOf(r) {
    return norm(
      r.driverName ||
      r.driver ||
      r.courierName ||
      r.courier_name ||
      r.courier ||
      r.employeeName ||
      r.driverFullName ||
      r.courierFullName ||
      r.vehicleDriverName ||
      r.vehicle_driver_name ||
      r.chauffeurName ||
      r.deliveryDriverName ||
      r.pickupDriverName ||
      r?.driver?.name ||
      r?.courier?.name ||
      r?.employee?.name ||
      r?.vehicle?.driverName ||
      r?.vehicle?.driver?.name ||
      r?.tour?.driverName ||
      r?.tour?.driver ||
      r?.route?.driverName ||
      ''
    );
  }

  function parcelListOf(r) {
    const directLists = [
      r?.parcel_number,
      r?.parcelNumbers,
      r?.parcels,
      r?.parcelNumberList,
      r?.parcelsList,
      r?.shipmentNumbers,
      r?.labels,
      r?.barcodes,
      r?.packages,
      r?.consignments,
      r?.shipments,
      r?.completed_parcel,
      r?.removed_parcel_numbers
    ];

    const out = [];

    function pushMaybe(val) {
      if (val == null || val === '') return;

      if (typeof val === 'string' || typeof val === 'number') {
        const txt = String(val).trim();

        if (txt.includes(',') || txt.includes(';') || txt.includes('|')) {
          txt.split(/[;,|]/).forEach(part => pushMaybe(part));
          return;
        }

        let x = txt.replace(/\D+/g, '');
        if (!x) return;
        if (x.length === 13) x = '0' + x;
        if (x.length >= 8) out.push(x);
        return;
      }

      if (Array.isArray(val)) {
        val.forEach(item => pushMaybe(item));
        return;
      }

      if (typeof val === 'object') {
        const cands = [
          val.parcelNumber, val.number, val.psn, val.shipmentNumber,
          val.barcode, val.labelNumber, val.consignmentNumber
        ];
        cands.forEach(c => pushMaybe(c));
      }
    }

    directLists.forEach(v => pushMaybe(v));
    return [...new Set(out)];
  }

  function parcelCountOf(r) {
    const list = parcelListOf(r);
    if (list.length) return list.length;

    const candidates = [
      r?.estimated_parcels,
      r?.completed_parcel,
      r?.parcelCount,
      r?.parcelsCount,
      r?.numberOfParcels,
      r?.numberOfPackages,
      r?.packageCount,
      r?.packagesCount,
      r?.shipmentCount,
      r?.quantity,
      r?.qty,
      r?.pieces,
      r?.pieceCount,
      r?.itemCount,
      r?.count,
      r?.totalParcels,
      r?.totalPackages,
      r?.consignmentCount
    ];

    for (const c of candidates) {
      const n = Number(c);
      if (Number.isFinite(n) && n > 0) return Math.trunc(n);
    }
    return 1;
  }

  function addrOf(r) {
    const street = norm(
      r.street ||
      r.addressLine1 ||
      r.address ||
      r?.address?.street ||
      r?.recipient?.street ||
      ''
    );
    const house = norm(
      r.houseno ||
      r.houseNo ||
      r.houseNumber ||
      r?.address?.houseNumber ||
      r?.recipient?.houseNumber ||
      ''
    );
    const postal = norm(
      r.postalCode ||
      r.zipCode ||
      r.zip ||
      r.postal_code ||
      r?.address?.postalCode ||
      r?.recipient?.postalCode ||
      ''
    );
    const city = norm(
      r.city ||
      r.town ||
      r?.address?.city ||
      r?.recipient?.city ||
      ''
    );

    const a1 = [street, house].filter(Boolean).join(' ');
    const a2 = [postal, city].filter(Boolean).join(' ');
    return [a1, a2].filter(Boolean).join(' · ') || '—';
  }

  function formatTime(v) {
    if (!v) return '';
    const d = new Date(v);
    if (isNaN(d)) return '';
    return d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' });
  }

  function createEmptySummaryRow(systempartner, tour = '', driver = '') {
    return {
      systempartner,
      tour,
      driver,
      driverSet: new Set(driver ? [driver] : []),
      deliveryStops: 0,
      deliveryParcels: 0,
      pickupStops: 0,
      pickupParcels: 0,
      canceledDeliveryStops: 0,
      canceledPickupStops: 0,
      hindranceStops: 0,
      hindranceCodes: Object.create(null)
    };
  }

  function finalizeDriverText(set) {
    const arr = Array.from(set || [])
      .map(x => norm(x))
      .filter(Boolean)
      .sort((a, b) => collator.compare(a, b));
    return arr.join(', ');
  }

  function totalsOfSummary(rows) {
    return rows.reduce((acc, r) => {
      acc.deliveryStops += r.deliveryStops;
      acc.deliveryParcels += r.deliveryParcels;
      acc.pickupStops += r.pickupStops;
      acc.pickupParcels += r.pickupParcels;
      acc.canceledDeliveryStops += r.canceledDeliveryStops;
      acc.canceledPickupStops += r.canceledPickupStops;
      acc.hindranceStops += r.hindranceStops || 0;

      if (r.hindranceCodes) {
        for (const [code, cnt] of Object.entries(r.hindranceCodes)) {
          acc.hindranceCodes[code] = (acc.hindranceCodes[code] || 0) + cnt;
        }
      }

      if (r.driver) {
        r.driver.split(/\s*,\s*/).filter(Boolean).forEach(name => acc.driverSet.add(name));
      }

      return acc;
    }, {
      driverSet: new Set(),
      deliveryStops: 0,
      deliveryParcels: 0,
      pickupStops: 0,
      pickupParcels: 0,
      canceledDeliveryStops: 0,
      canceledPickupStops: 0,
      hindranceStops: 0,
      hindranceCodes: Object.create(null)
    });
  }

  function aggregateSummaryFromRows(rows, mode) {
    const map = new Map();

    for (const r of rows) {
      const partnerName = norm(r.__partner) || 'Ohne Zuordnung';

      const key = mode === 'partner'
        ? partnerName
        : `${partnerName}|||${String(r.__tour || '—')}`;

      let g = map.get(key);
      if (!g) {
        g = createEmptySummaryRow(
          partnerName,
          mode === 'tour' ? String(r.__tour || '—') : '',
          mode === 'tour' ? (r.__driver || '') : ''
        );
        map.set(key, g);
      }

      if (r.__driver) g.driverSet.add(r.__driver);
      g.driver = finalizeDriverText(g.driverSet);

      if (r.__type === 'DELIVERY') {
        g.deliveryStops += 1;
        g.deliveryParcels += r.__pkgCount;

        if (isCanceledDeliveryStatus(r.__statusNorm || r.__status, r.__raw)) {
          g.canceledDeliveryStops += 1;
        }

        if (isDeliveryHindranceStatus(r.__statusNorm || r.__status, r.__raw)) {
          const code = cleanAdditionalCode(r.__additionalCode);
          if (!code) continue;

          const itemCount = hindranceItemCount(r);
          g.hindranceStops += itemCount;
          g.hindranceCodes[code] = (g.hindranceCodes[code] || 0) + itemCount;
        }
      } else {
        g.pickupStops += 1;
        g.pickupParcels += r.__pkgCount;

        if (isCanceledPickupStatus(r.__statusNorm || r.__status, r.__raw)) {
          g.canceledPickupStops += 1;
        }
      }
    }

    const arr = Array.from(map.values()).map(r => ({
      ...r,
      driver: finalizeDriverText(r.driverSet),
      driverSet: undefined
    }));

    if (mode === 'partner') {
      arr.sort((a, b) => collator.compare(a.systempartner, b.systempartner));
    } else {
      arr.sort((a, b) => {
        const c1 = collator.compare(a.systempartner, b.systempartner);
        if (c1 !== 0) return c1;
        return collator.compare(String(a.tour || ''), String(b.tour || ''));
      });
    }
    return arr;
  }

  function getCachedSummary(key) {
    const hit = summaryCache.get(key);
    if (!hit) return null;
    if ((Date.now() - hit.ts) > SUMMARY_CACHE_MS) {
      summaryCache.delete(key);
      return null;
    }
    return hit.data;
  }

  function setCachedSummary(key, data) {
    summaryCache.set(key, { ts: Date.now(), data });
  }

  function getCachedDetailRows(key) {
    const hit = detailCache.get(key);
    if (!hit) return null;
    if ((Date.now() - hit.ts) > DETAIL_CACHE_MS) {
      detailCache.delete(key);
      return null;
    }
    return hit.rows;
  }

  function setCachedDetailRows(key, rows) {
    detailCache.set(key, { ts: Date.now(), rows });
  }

  function clearRuntimeCaches() {
    summaryCache.clear();
    detailCache.clear();
    tourPartnerLookupCache = new Map();
    tourDriverLookupCache = new Map();
    tourPartnerMap = new Map();
    tourDriverMap = new Map();
    tourMapLoadedAt = 0;
  }

  function tpIdbOpen() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(STORE_DB);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function tpIdbAll(store) {
    try {
      const db = await tpIdbOpen();
      if (!db.objectStoreNames.contains(store)) return [];
      return await new Promise((resolve) => {
        const r = db.transaction(store, 'readonly').objectStore(store).getAll();
        r.onsuccess = () => resolve(r.result || []);
        r.onerror = () => resolve([]);
      });
    } catch {
      return [];
    }
  }

  async function loadTourPartnerMap(force = false) {
    const now = Date.now();
    if (!force && tourPartnerMap.size && (now - tourMapLoadedAt) < TOURMAP_CACHE_MS) return;

    const mergedPartner = new Map();
    const mergedDriver = new Map();

    function addKey(rawTour, partner, driver) {
      const p = norm(partner || '');
      const d = norm(driver || '');

      const raw = String(rawTour || '').trim();
      if (!raw) return;

      const variants = new Set();
      variants.add(raw);
      variants.add(raw.replace(/\s+/g, ''));
      variants.add(raw.replace(',', '.'));
      variants.add(raw.replace(/[^\dA-Za-z]/g, ''));

      const k = tourKey(raw);
      if (k) {
        variants.add(k);
        const n = Number(k);
        if (!Number.isNaN(n)) {
          variants.add(String(Math.trunc(n)));
          variants.add(String(Math.trunc(n)).padStart(3, '0'));
          variants.add(String(Math.trunc(n)).padStart(4, '0'));
        }
      }

      for (const v of variants) {
        const key = tourKey(v);
        if (!key) continue;
        if (p && !mergedPartner.has(key)) mergedPartner.set(key, p);
        if (d && !mergedDriver.has(key)) mergedDriver.set(key, d);
      }
    }

    if (lastOkRequest) {
      try {
        const base = new URL('/dispatcher/api/vehicle-overview', location.origin);
        base.searchParams.set('pageSize', String(PAGE_SIZE));
        base.searchParams.set('dateFrom', getSelectedDateFrom() || getDefaultDateFrom());
        base.searchParams.set('dateTo', getSelectedDateTo() || getDefaultDateTo());
        const headers = buildHeaders(lastOkRequest.headers);

        for (let p = 1; p <= HARD_MAX_PAGES; p++) {
          const u = new URL(base.toString());
          u.searchParams.set('page', String(p));
          u.searchParams.set('_ts', String(Date.now() + p));
          const j = await fetchJsonPage(u, headers);
          const arr = pickArray(j);

          for (const r of arr) {
            const tour =
              r.tour || r.round || r.route || r.tourNumber || r.routeNumber ||
              r.tourNo || r.routeNo || r?.tour?.number || r?.route?.number || '';
            const partner =
              r.subcontractor_name || r.subcontractorName || r.partnerName || r.systemPartner || r.systempartner || '';
            const driver =
              r.driverName || r.driver || r.courierName || r.employeeName ||
              r.driverFullName || r.courierFullName || r.vehicleDriverName ||
              r?.driver?.name || r?.employee?.name || r?.vehicle?.driverName ||
              r?.vehicle?.driver?.name || '';
            addKey(tour, partner, driver);
          }

          if (!arr.length || arr.length < PAGE_SIZE) break;
        }
      } catch (e) {
        console.warn('vehicle-overview konnte nicht geladen werden, nutze Fallback.', e);
      }
    }

    try {
      const rows = await tpIdbAll(STORE_TOURMAP);
      for (const r of rows) {
        addKey(
          r.tour || r.route || r.tourNo || r.tourNumber || r.routeNumber,
          r.partner,
          r.driver || r.driverName || ''
        );
      }
    } catch {}

    tourPartnerMap = mergedPartner;
    tourDriverMap = mergedDriver;
    tourPartnerLookupCache = new Map();
    tourDriverLookupCache = new Map();
    tourMapLoadedAt = now;
  }

  function partnerOfTour(tour) {
    const key = tourKey(tour);
    if (!key) return '';
    if (tourPartnerLookupCache.has(key)) return tourPartnerLookupCache.get(key);
    const p = norm(tourPartnerMap.get(key) || '');
    tourPartnerLookupCache.set(key, p);
    return p;
  }

  function driverOfTour(tour) {
    const key = tourKey(tour);
    if (!key) return '';
    if (tourDriverLookupCache.has(key)) return tourDriverLookupCache.get(key);
    const d = norm(tourDriverMap.get(key) || '');
    tourDriverLookupCache.set(key, d);
    return d;
  }

  function resolvePartner(rawRow, tour) {
    const direct = rawPartnerOfRow(rawRow);
    if (useDirectColumnsForRow(rawRow) && direct) return direct;
    return partnerOfTour(tour) || direct || 'Ohne Zuordnung';
  }

  function resolveDriver(rawRow, tour) {
    const direct = driverNameOf(rawRow);
    if (useDirectColumnsForRow(rawRow) && direct) return direct;
    return driverOfTour(tour) || direct || '';
  }

  function normalizeRows(rows) {
    const out = [];

    for (let idx = 0; idx < rows.length; idx++) {
      const r = rows[idx];
      const tour = extractTour(r);
      const driver = resolveDriver(r, tour);
      const partner = resolvePartner(r, tour);

      const type = orderTypeOf(r);
      const status = statusOf(r);
      const statusNorm = normalizeStatusForMatch(status);
      const parcels = parcelListOf(r);
      const pkgCount = Math.max(1, parcelCountOf(r));
      const addr = addrOf(r);
      const additionalCode = additionalCodeOf(r);
      const serviceCode = serviceCodeOf(r);

      const deliveredAt =
        r.deliveredTime ||
        r.delivered_time ||
        r.deliveryTime ||
        r.deliveryDateTime ||
        r.deliveryTimestamp ||
        r?.delivery?.time ||
        r?.delivery?.dateTime ||
        r?.statusTime ||
        '';

      const pickupAt =
        r.pickupTime ||
        r.pickedUpTime ||
        r.pickupDateTime ||
        r.pickupTimestamp ||
        r?.pickup?.time ||
        r?.pickup?.dateTime ||
        r?.statusTime ||
        '';

      out.push({
        __raw: r,
        __partner: partner,
        __tour: tour || '—',
        __status: statusDe(status || '—'),
        __statusNorm: statusNorm,
        __type: type,
        __addr: addr,
        __pkgCount: pkgCount,
        __parcelList: parcels,
        __stopKey: String(r.stopId ?? r.id ?? `${tour}#${addr}#${idx}`),
        __deliveredAt: deliveredAt,
        __pickupAt: pickupAt,
        __driver: driver,
        __additionalCode: additionalCode,
        __serviceCode: serviceCode,
        __extraFields: extraFieldMapOf(r)
      });
    }

    return out;
  }

  function buildHeaders(h) {
    const H = new Headers();
    try {
      if (h) {
        Object.entries(h).forEach(([k, v]) => {
          const key = String(k).toLowerCase();
          if (['authorization', 'accept', 'x-xsrf-token', 'x-csrf-token'].includes(key)) {
            H.set(key === 'accept' ? 'Accept' : key.replace(/(^.|-.)/g, s => s.toUpperCase()), v);
          }
        });
      }
      if (!H.has('Accept')) H.set('Accept', 'application/json, text/plain, */*');
    } catch {}
    return H;
  }

  function pickArray(payload) {
    if (Array.isArray(payload)) return payload;
    if (payload && Array.isArray(payload.items)) return payload.items;
    if (payload && Array.isArray(payload.content)) return payload.content;
    if (payload && payload.data) {
      if (Array.isArray(payload.data)) return payload.data;
      if (Array.isArray(payload.data.items)) return payload.data.items;
      if (Array.isArray(payload.data.content)) return payload.data.content;
    }
    if (payload && Array.isArray(payload.results)) return payload.results;
    if (payload && payload._embedded) {
      const v = Object.values(payload._embedded).find(Array.isArray);
      if (Array.isArray(v)) return v;
    }
    return [];
  }

  function getPickupDeliveryBaseUrl() {
    if (lastOkRequest?.url) {
      const u = new URL(lastOkRequest.url.href);
      return new URL(`${u.origin}/dispatcher/api/pickup-delivery`);
    }
    return new URL('/dispatcher/api/pickup-delivery', location.origin);
  }

  function buildUrlAll(page) {
    const u = getPickupDeliveryBaseUrl();
    const q = u.searchParams;

    const fromDate = getSelectedDateFrom() || state.detectedDateFrom || state.detectedDateValue || toIsoDateLocal(new Date());
    const toDate = getSelectedDateTo() || state.detectedDateTo || state.detectedDateValue || toIsoDateLocal(new Date());

    q.set('page', String(page));
    q.set('pageSize', String(PAGE_SIZE));
    q.set('dateFrom', fromDate);
    q.set('dateTo', toDate);
    q.set('_ts', String(Date.now() + page));

    [
      'priority',
      'elements',
      'parcelNumber',
      'parcel_number',
      'orderType',
      'order_type',
      'type',
      'status',
      'deliveryStatus',
      'delivery_status',
      'pickupStatus',
      'pickup_status',
      'tab',
      'activeTab',
      'selectedTab',
      'mode'
    ].forEach(k => q.delete(k));

    return u;
  }

  async function fetchJsonPage(url, headers) {
    const res = await fetch(url.toString(), {
      credentials: 'include',
      headers,
      cache: 'no-store'
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  }

  async function fetchAllRawRows(onProgress) {
    if (!lastOkRequest) throw new Error('Bitte zuerst einmal die normale Pickup-/Delivery-Liste laden.');
    const headers = buildHeaders(lastOkRequest.headers);
    const rows = [];
    const seen = new Set();

    let page = 1;
    while (page <= HARD_MAX_PAGES) {
      const url = buildUrlAll(page);
      const json = await fetchJsonPage(url, headers);
      const arr = pickArray(json);

      if (!arr.length) break;

      for (const item of arr) {
        const key = String(item?.id ?? `${item?.tour || ''}|${item?.parcel_number || item?.parcelNumber || ''}|${page}|${rows.length}`);
        if (seen.has(key)) continue;
        seen.add(key);
        rows.push(item);
      }

      if (onProgress) {
        onProgress({
          loadedPages: page,
          totalPages: '?',
          rowsLoaded: rows.length
        });
      }

      if (arr.length < PAGE_SIZE) break;
      page += 1;
    }

    return rows;
  }

  async function fetchSummaryOnly(onProgress) {
    await loadTourPartnerMap(true);
    const rawRows = await fetchAllRawRows(onProgress);
    const normalized = normalizeRows(rawRows);
    const summaryRows = aggregateSummaryFromRows(normalized, 'partner');
    const loadedDate = parsePossibleDateFromRows(rawRows) || getSelectedDateFrom() || '';
    return { summaryRows, loadedDate, normalized };
  }

  async function fetchDetailRows(onProgress, forceReload = false) {
    const cacheKey = getSummaryCacheKey();
    const cached = !forceReload ? getCachedDetailRows(cacheKey) : null;
    if (cached) return cached;

    await loadTourPartnerMap(forceReload);

    const rawRows = await fetchAllRawRows(onProgress);
    const normalized = normalizeRows(rawRows);

    setCachedDetailRows(cacheKey, normalized);
    return normalized;
  }

  function setLoadedStandFromDateRange(fromDate, toDate) {
    const now = new Date();
    state.standText =
      `Daten vom ${formatDateDE(fromDate || '—')} bis ${formatDateDE(toDate || '—')} · geladen am ${now.toLocaleDateString('de-DE')} ${now.toLocaleTimeString('de-DE', {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      })}`;
    renderStand();
  }

  function ensureStyles() {
    if (document.getElementById(NS + 'style')) return;

    const style = document.createElement('style');
    style.id = NS + 'style';
    style.textContent = `
      .${NS}panel{
        position:fixed; top:72px; left:50%; transform:translateX(-50%);
        width:min(1800px,98vw); max-height:84vh; overflow:visible;
        background:#fff; border:1px solid rgba(0,0,0,.14);
        box-shadow:0 12px 28px rgba(0,0,0,.20); border-radius:12px;
        z-index:100000; display:none; font:12px system-ui;
      }
      .${NS}head{
        display:flex; justify-content:space-between; align-items:center; gap:8px;
        padding:8px 10px; border-bottom:1px solid rgba(0,0,0,.08);
        position:sticky; top:0; background:#fff; z-index:5;
      }
      .${NS}title{font:700 14px system-ui}
      .${NS}sub{opacity:.75; font:600 12px system-ui; margin-top:2px}
      .${NS}stand{opacity:.85; font:700 12px system-ui; margin-top:4px; color:#0f3f75}
      .${NS}controls{display:flex; gap:8px; align-items:center; flex-wrap:wrap; justify-content:flex-end}
      .${NS}label{font:600 12px system-ui; opacity:.9}
      .${NS}select{
        border:1px solid rgba(0,0,0,.14); background:#fff;
        padding:6px 10px; border-radius:8px; cursor:pointer;
        font:600 12px system-ui; min-width:150px;
      }
      .${NS}btn{
        border:1px solid rgba(0,0,0,.14); background:#f7f7f7;
        padding:6px 10px; border-radius:8px; cursor:pointer;
        font:600 12px system-ui;
      }
      .${NS}btn:hover{background:#efefef}
      .${NS}body{padding:0; max-height:calc(84vh - 86px); overflow:auto}
      .${NS}loading{
        display:none; padding:8px 12px; background:#fffbe6;
        border-bottom:1px solid rgba(0,0,0,.08); font:600 12px system-ui;
      }
      .${NS}loading.on{display:block}
      .${NS}tbl{
        width:100%; border-collapse:separate; border-spacing:0; table-layout:fixed;
        font:12px system-ui;
      }
      .${NS}tbl th, .${NS}tbl td{
        border-bottom:1px solid rgba(0,0,0,.08);
        border-right:1px solid rgba(0,0,0,.08);
        padding:5px 6px;
        text-align:right;
        overflow:hidden;
        text-overflow:ellipsis;
        vertical-align:top;
      }
      .${NS}tbl th:first-child, .${NS}tbl td:first-child{
        text-align:left;
        border-left:1px solid rgba(0,0,0,.08);
      }
      .${NS}tbl thead th{
        position:sticky; top:0; z-index:2;
        background:#ead0d0; color:#8a1f1f; font-weight:700;
      }
      .${NS}tbl tfoot td{
        background:#d7e6f3; color:#0b3d6b; font-weight:700;
      }
      .${NS}tbl tbody tr:hover{background:#f8fafc}
      .${NS}link,.${NS}numlink{
        color:#0f3f75; text-decoration:none; cursor:pointer; font-weight:700;
      }
      .${NS}muted{opacity:.7; font-size:11px; margin-left:4px}
      .${NS}psn-wrap{display:flex; flex-wrap:wrap; gap:4px; justify-content:flex-start}
      .${NS}psn-btn{
        border:1px solid rgba(0,0,0,.14); background:#f7f7f7; border-radius:6px;
        padding:2px 6px; cursor:pointer; font:600 11px system-ui; color:#0f3f75;
      }
      .${NS}psn-btn:hover{background:#efefef}
      .${NS}empty{padding:16px; text-align:center; opacity:.75; font:600 12px system-ui}
      .${NS}modal{
        position:fixed; inset:0; background:rgba(0,0,0,.36);
        display:none; align-items:flex-start; justify-content:center;
        z-index:100001; box-sizing:border-box;
      }
      .${NS}modal-inner{
        background:#fff; width:min(1760px,97vw); max-height:min(88vh,1050px);
        margin-top:50px; border-radius:12px; box-shadow:0 12px 28px rgba(0,0,0,.22);
        border:1px solid rgba(0,0,0,.12); overflow:hidden; display:flex; flex-direction:column;
      }
      .${NS}modal-head{
        display:flex; justify-content:space-between; align-items:center; gap:8px;
        padding:10px 12px; border-bottom:1px solid rgba(0,0,0,.08); background:#fff;
        z-index:5; flex:0 0 auto;
      }
      .${NS}modal-title{font:700 13px system-ui}
      .${NS}copy-note{
        padding:0 12px 8px 12px; opacity:.75; font:600 12px system-ui;
        background:#fff; flex:0 0 auto;
      }
      .${NS}modal-body{padding:0 12px 12px 12px; overflow:auto; flex:1 1 auto; min-height:0}
      .${NS}modal-scroll{overflow:visible; max-height:none; min-height:0}
      .${NS}sort-asc::after{content:" ▲";font-size:11px}
      .${NS}sort-desc::after{content:" ▼";font-size:11px}
      .${NS}status{
        display:inline-block; max-width:100%; padding:2px 8px; border-radius:999px;
        font:700 11px system-ui; border:1px solid transparent; white-space:nowrap;
        overflow:hidden; text-overflow:ellipsis;
      }
      .${NS}status.gray{background:#f3f4f6;color:#374151;border-color:#d1d5db}
      .${NS}status.red{background:#fee2e2;color:#991b1b;border-color:#fca5a5}
      .${NS}status.green{background:#dcfce7;color:#166534;border-color:#86efac}
      .${NS}status.yellow{background:#fef3c7;color:#92400e;border-color:#fcd34d}
      .${NS}status.blue{background:#dbeafe;color:#1d4ed8;border-color:#93c5fd}
      .${NS}status.purple{background:#ede9fe;color:#6d28d9;border-color:#c4b5fd}

      .${NS}col-partner{
        width:22%;
        min-width:160px;
      }
      .${NS}col-driver{
        width:16%;
        min-width:150px;
      }
      .${NS}partner-cell{
        white-space:nowrap;
      }
      .${NS}driver-cell{
        white-space:normal;
        line-height:1.2;
        display:-webkit-box;
        -webkit-box-orient:vertical;
        -webkit-line-clamp:2;
        line-clamp:2;
        overflow:hidden;
        text-overflow:ellipsis;
        word-break:break-word;
      }

      @media (max-width:1200px){
        .${NS}col-partner{ width:20%; min-width:130px; }
        .${NS}col-driver{ width:14%; min-width:120px; }
      }

      @media (max-width:900px){
        .${NS}col-partner{ width:18%; min-width:100px; }
        .${NS}col-driver{ width:12%; min-width:100px; }
      }
    `;
    document.head.appendChild(style);
  }

  function renderStand() {
    const el = document.getElementById(NS + 'stand');
    if (el) el.textContent = state.standText || 'Daten vom —';
  }

  function renderDateOptions() {
    const fromEl = document.getElementById(NS + 'date-from');
    const toEl = document.getElementById(NS + 'date-to');
    if (!fromEl || !toEl) return;

    fromEl.value = fromEl.value || getDefaultDateFrom();
    toEl.value = toEl.value || getDefaultDateTo();

    const title = state.detectedDateParam
      ? `Erkanntes Datumsfeld: ${state.detectedDateParam}`
      : 'Kein Datumsfeld sicher erkannt – Bereich wird bestmöglich gesetzt';

    fromEl.title = title;
    toEl.title = title;

    fromEl.disabled = !lastOkRequest;
    toEl.disabled = !lastOkRequest;

    ensureValidDateRange();
  }

  function setLoading(on, text) {
    const el = document.getElementById(NS + 'loading');
    if (!el) return;
    el.classList.toggle('on', !!on);
    if (text) el.textContent = text;
    else if (!on) el.textContent = 'Lade Daten …';
  }

  function mountUI() {
    ensureStyles();
    if (document.getElementById(NS + 'panel')) return;

    const panel = document.createElement('div');
    panel.id = NS + 'panel';
    panel.className = NS + 'panel';
    panel.innerHTML = `
      <div class="${NS}head">
        <div>
          <div class="${NS}title">Tourenauswertung</div>
          <div class="${NS}sub">Systempartner / Touren · Fahrername · Stopps und Pakete · Klick auf Zahlen öffnet Listen</div>
          <div class="${NS}stand" id="${NS}stand">Daten vom —</div>
        </div>
        <div class="${NS}controls">
          <label class="${NS}label" for="${NS}date-from">Datum von:</label>
          <input id="${NS}date-from" class="${NS}select" type="date" style="min-width:160px">
          <label class="${NS}label" for="${NS}date-to">Datum bis:</label>
          <input id="${NS}date-to" class="${NS}select" type="date" style="min-width:160px">
          <button class="${NS}btn" data-action="refresh">Aktualisieren</button>
          <button class="${NS}btn" data-action="chooseColumns">Spalten</button>
          <button class="${NS}btn" data-action="copyMain">Tabelle kopieren</button>
        </div>
      </div>
      <div id="${NS}loading" class="${NS}loading">Lade Daten …</div>
      <div class="${NS}body">
        <div id="${NS}note" class="${NS}empty">Bitte einmal die normale Pickup-/Delivery-Liste laden, damit der API-Request geklont werden kann.</div>
        <div id="${NS}table-wrap"></div>
      </div>
    `;
    document.body.appendChild(panel);

    const fromEl = panel.querySelector('#' + NS + 'date-from');
    const toEl = panel.querySelector('#' + NS + 'date-to');
    if (fromEl) {
      fromEl.value = getDefaultDateFrom();
      fromEl.addEventListener('change', ensureValidDateRange);
    }
    if (toEl) {
      toEl.value = getDefaultDateTo();
      toEl.addEventListener('change', ensureValidDateRange);
    }

    const modal = document.createElement('div');
    modal.id = NS + 'modal';
    modal.className = NS + 'modal';
    modal.innerHTML = `
      <div class="${NS}modal-inner">
        <div class="${NS}modal-head">
          <div class="${NS}modal-title" id="${NS}modal-title">Liste</div>
          <div style="display:flex;gap:8px;align-items:center">
            <button class="${NS}btn" data-action="copyModal">Tabelle kopieren</button>
            <button class="${NS}btn" data-action="closeModal">Schließen</button>
          </div>
        </div>
        <div class="${NS}copy-note">Kopie erfolgt als formatierte HTML-Tabelle und zusätzlich als Text für Excel.</div>
        <div class="${NS}modal-body" id="${NS}modal-body"></div>
      </div>
    `;
    document.body.appendChild(modal);

    panel.addEventListener('click', async (e) => {
      const btn = e.target.closest('button[data-action]');
      const link = e.target.closest('[data-kind]');
      if (btn) {
        const action = btn.dataset.action;
        if (action === 'refresh') {
          await fullRefresh({ forceReload: true }).catch(console.error);
          return;
        }
        if (action === 'copyMain') {
          await copyMainTable().catch(console.error);
          return;
        }
        if (action === 'chooseColumns') {
          openColumnChooser();
          return;
        }
      }
      if (link) await handleMainTableClick(link).catch(console.error);
    });

    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        hideModal();
        return;
      }

      const btn = e.target.closest('button[data-action]');
      const psnBtn = e.target.closest('button[data-psn]');
      const th = e.target.closest('th[data-col]');
      const link = e.target.closest('[data-kind]');

      if (btn) {
        const action = btn.dataset.action;
        if (action === 'closeModal') {
          hideModal();
          return;
        }
        if (action === 'copyModal') {
          copyModalTable().catch(console.error);
          return;
        }
        if (action === 'saveColumns') {
          saveColumnChooser();
          return;
        }
        if (action === 'resetColumns') {
          setSelectedExtraColumns([]);
          hideModal();
          return;
        }
      }

      if (psnBtn) {
        openParcelTracking(psnBtn.dataset.psn || '');
        return;
      }

      if (th && state.modal.type === 'table') {
        sortModalByColumn(Number(th.dataset.col || 0));
        return;
      }

      if (link && link.dataset.kind === 'hindranceCodeDetails') {
        openHindranceCodeDetails({
          partner: String(link.dataset.partner || ''),
          tour: String(link.dataset.tour || ''),
          code: String(link.dataset.code || '')
        }).catch(console.error);
        return;
      }

      if (link) handleMainTableClick(link).catch(console.error);
    });



    // Klick außerhalb schließt Modal bzw. Hauptfenster
    document.addEventListener('mousedown', function(e) {

      const modal = document.getElementById(NS + 'modal');
      const modalInner = modal?.querySelector('.' + NS + 'modal-inner');

      if (modal && modal.style.display === 'flex') {
        if (modalInner && !modalInner.contains(e.target)) {
          hideModal();
        }
        return;
      }

      const panel = document.getElementById(NS + 'panel');

      if (
        panel &&
        getComputedStyle(panel).display !== 'none' &&
        !panel.contains(e.target)
      ) {
        panel.style.setProperty('display', 'none', 'important');
      }

    }, true);

renderDateOptions();
    renderStand();
    renderMainTable();
  }

  function togglePanel(force) {
    const panel = document.getElementById(NS + 'panel');
    if (!panel) {
      mountUI();
      return;
    }
    const hidden = getComputedStyle(panel).display === 'none';
    const show = force != null ? !!force : hidden;
    panel.style.setProperty('display', show ? 'block' : 'none', 'important');
  }

function renderMainTable() {
  const wrap = document.getElementById(NS + 'table-wrap');
  if (!wrap) return;

  if (!state.summaryRows.length) {
    wrap.innerHTML = '';
    const note = document.getElementById(NS + 'note');
    if (note && !lastOkRequest) note.style.display = '';
    return;
  }

  const total = totalsOfSummary(state.summaryRows);
  const totalDriverText = finalizeDriverText(total.driverSet);

  wrap.innerHTML = `
    <div style="overflow:auto">
      <table class="${NS}tbl" id="${NS}main-table">
        <thead>
          <tr>
            <th class="${NS}col-partner">Systempartner</th>
            <th class="${NS}col-driver">Fahrername</th>
            <th>Stopps Zustellung</th>
            <th>Pakete Zustellung</th>
            <th>Stopps Abholung</th>
            <th>Pakete Abholung</th>
            <th>Nicht bearb. Zustellstopps</th>
            <th>Nicht bearb. Abholstopps</th>
            <th>Zustellhindernisse</th>
          </tr>
        </thead>
        <tbody>
          ${state.summaryRows.map(r => `
            <tr>
              <td class="${NS}partner-cell">
                <span class="${NS}link" data-kind="partner" data-partner="${esc(r.systempartner)}">${esc(r.systempartner)}</span>
              </td>
              <td class="${NS}driver-cell" title="${esc(r.driver || '—')}">${esc(r.driver || '—')}</td>
              <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(r.systempartner)}" data-metric="deliveryStops">${r.deliveryStops}</span></td>
              <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(r.systempartner)}" data-metric="deliveryParcels">${r.deliveryParcels}</span></td>
              <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(r.systempartner)}" data-metric="pickupStops">${r.pickupStops}</span></td>
              <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(r.systempartner)}" data-metric="pickupParcels">${r.pickupParcels}</span></td>
              <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(r.systempartner)}" data-metric="canceledDeliveryStops">${r.canceledDeliveryStops}</span></td>
              <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(r.systempartner)}" data-metric="canceledPickupStops">${r.canceledPickupStops}</span></td>
              <td><span class="${NS}numlink" data-kind="hindranceCodes" data-level="partner" data-partner="${esc(r.systempartner)}">${r.hindranceStops || 0}</span></td>
            </tr>
          `).join('')}
        </tbody>
        <tfoot>
          <tr>
            <td>
              <span class="${NS}link" data-kind="partnerAll">Gesamt</span>
            </td>
            <td class="${NS}driver-cell" title="${esc(totalDriverText || '—')}">${esc(totalDriverText || '—')}</td>
            <td><span class="${NS}numlink" data-kind="metric" data-level="all" data-metric="deliveryStops">${total.deliveryStops}</span></td>
            <td><span class="${NS}numlink" data-kind="metric" data-level="all" data-metric="deliveryParcels">${total.deliveryParcels}</span></td>
            <td><span class="${NS}numlink" data-kind="metric" data-level="all" data-metric="pickupStops">${total.pickupStops}</span></td>
            <td><span class="${NS}numlink" data-kind="metric" data-level="all" data-metric="pickupParcels">${total.pickupParcels}</span></td>
            <td><span class="${NS}numlink" data-kind="metric" data-level="all" data-metric="canceledDeliveryStops">${total.canceledDeliveryStops}</span></td>
            <td><span class="${NS}numlink" data-kind="metric" data-level="all" data-metric="canceledPickupStops">${total.canceledPickupStops}</span></td>
            <td><span class="${NS}numlink" data-kind="hindranceCodes" data-level="all">${total.hindranceStops || 0}</span></td>
          </tr>
        </tfoot>
      </table>
    </div>
  `;
}

async function handleMainTableClick(el) {
  const kind = el.dataset.kind;

  if (kind === 'partnerAll') {
    setLoading(true, 'Lade alle Touren …');
    try {
      const rows = await fetchDetailRows(({ loadedPages, totalPages, rowsLoaded }) => {
        setLoading(true, `Lade Detaildaten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
      });

      const grouped = aggregateSummaryFromRows(rows, 'tour');

      const html = `
        <div style="overflow:auto">
          <table class="${NS}tbl" id="${NS}tour-table-all">
            <thead>
              <tr>
                <th class="${NS}col-partner">Systempartner</th>
                <th>Tour</th>
                <th class="${NS}col-driver">Fahrername</th>
                <th>Stopps Zustellung</th>
                <th>Pakete Zustellung</th>
                <th>Stopps Abholung</th>
                <th>Pakete Abholung</th>
                <th>Nicht bearb. Zustellstopps</th>
                <th>Nicht bearb. Abholstopps</th>
                <th>Zustellhindernisse</th>
              </tr>
            </thead>
            <tbody>
              ${grouped.map(r => `
                <tr>
                  <td class="${NS}partner-cell">${esc(r.systempartner)}</td>
                  <td>
                    <span class="${NS}link"
                          data-kind="tourAll"
                          data-partner="${esc(r.systempartner)}"
                          data-tour="${esc(r.tour)}">${esc(r.tour || '—')}</span>
                  </td>
                  <td class="${NS}driver-cell" title="${esc(r.driver || '—')}">${esc(r.driver || '—')}</td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(r.systempartner)}" data-tour="${esc(r.tour)}" data-metric="deliveryStops">${r.deliveryStops}</span></td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(r.systempartner)}" data-tour="${esc(r.tour)}" data-metric="deliveryParcels">${r.deliveryParcels}</span></td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(r.systempartner)}" data-tour="${esc(r.tour)}" data-metric="pickupStops">${r.pickupStops}</span></td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(r.systempartner)}" data-tour="${esc(r.tour)}" data-metric="pickupParcels">${r.pickupParcels}</span></td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(r.systempartner)}" data-tour="${esc(r.tour)}" data-metric="canceledDeliveryStops">${r.canceledDeliveryStops}</span></td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(r.systempartner)}" data-tour="${esc(r.tour)}" data-metric="canceledPickupStops">${r.canceledPickupStops}</span></td>
                  <td><span class="${NS}numlink" data-kind="hindranceCodes" data-level="tour" data-partner="${esc(r.systempartner)}" data-tour="${esc(r.tour)}">${r.hindranceStops || 0}</span></td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      `;

      openModalHtml('Alle Touren', html);
    } finally {
      setLoading(false);
    }
    return;
  }

  if (kind === 'partner') {
    const partner = String(el.dataset.partner || '');
    await openPartnerTours(partner);
    return;
  }

  if (kind === 'tourAll') {
    const partner = String(el.dataset.partner || '');
    const tour = String(el.dataset.tour || '');
    await openTourAllList({ partner, tour });
    return;
  }

  if (kind === 'metric') {
    const level = String(el.dataset.level || 'partner');
    const partner = String(el.dataset.partner || '');
    const tour = String(el.dataset.tour || '');
    const metric = String(el.dataset.metric || '');
    await openMetricList({ level, partner, tour, metric });
    return;
  }

  if (kind === 'hindranceCodes') {
    const level = String(el.dataset.level || 'partner');
    const partner = String(el.dataset.partner || '');
    const tour = String(el.dataset.tour || '');
    await openHindranceCodeSummary({ level, partner, tour });
  }
}
  async function openPartnerTours(partner) {
    setLoading(true, `Lade Detaildaten für ${partner} …`);
    try {
      const rows = await fetchDetailRows(({ loadedPages, totalPages, rowsLoaded }) => {
        setLoading(true, `Lade Detaildaten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
      });

      const source = rows.filter(r => r.__partner === partner);
      const grouped = aggregateSummaryFromRows(source, 'tour');
      const total = totalsOfSummary(grouped);
      const totalDriverText = finalizeDriverText(total.driverSet);

      const html = `
        <div style="overflow:auto">
          <table class="${NS}tbl" id="${NS}tour-table">
            <thead>
              <tr>
                <th style="width:14%">Tour</th>
                <th class="${NS}col-driver">Fahrername</th>
                <th>Stopps Zustellung</th>
                <th>Pakete Zustellung</th>
                <th>Stopps Abholung</th>
                <th>Pakete Abholung</th>
                <th>Nicht bearb. Zustellstopps</th>
                <th>Nicht bearb. Abholstopps</th>
                <th>Zustellhindernisse</th>
              </tr>
            </thead>
            <tbody>
              ${grouped.map(r => `
                <tr>
                  <td>
                    <span class="${NS}link"
                          data-kind="tourAll"
                          data-partner="${esc(partner)}"
                          data-tour="${esc(r.tour)}">${esc(r.tour || '—')}</span>
                  </td>
                  <td class="${NS}driver-cell" title="${esc(r.driver || '—')}">${esc(r.driver || '—')}</td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(partner)}" data-tour="${esc(r.tour)}" data-metric="deliveryStops">${r.deliveryStops}</span></td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(partner)}" data-tour="${esc(r.tour)}" data-metric="deliveryParcels">${r.deliveryParcels}</span></td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(partner)}" data-tour="${esc(r.tour)}" data-metric="pickupStops">${r.pickupStops}</span></td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(partner)}" data-tour="${esc(r.tour)}" data-metric="pickupParcels">${r.pickupParcels}</span></td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(partner)}" data-tour="${esc(r.tour)}" data-metric="canceledDeliveryStops">${r.canceledDeliveryStops}</span></td>
                  <td><span class="${NS}numlink" data-kind="metric" data-level="tour" data-partner="${esc(partner)}" data-tour="${esc(r.tour)}" data-metric="canceledPickupStops">${r.canceledPickupStops}</span></td>
                  <td><span class="${NS}numlink" data-kind="hindranceCodes" data-level="tour" data-partner="${esc(partner)}" data-tour="${esc(r.tour)}">${r.hindranceStops || 0}</span></td>
                </tr>
              `).join('')}
            </tbody>
            <tfoot>
              <tr>
                <td>Gesamt</td>
                <td class="${NS}driver-cell" title="${esc(totalDriverText || '—')}">${esc(totalDriverText || '—')}</td>
                <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(partner)}" data-metric="deliveryStops">${total.deliveryStops}</span></td>
                <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(partner)}" data-metric="deliveryParcels">${total.deliveryParcels}</span></td>
                <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(partner)}" data-metric="pickupStops">${total.pickupStops}</span></td>
                <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(partner)}" data-metric="pickupParcels">${total.pickupParcels}</span></td>
                <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(partner)}" data-metric="canceledDeliveryStops">${total.canceledDeliveryStops}</span></td>
                <td><span class="${NS}numlink" data-kind="metric" data-level="partner" data-partner="${esc(partner)}" data-metric="canceledPickupStops">${total.canceledPickupStops}</span></td>
                <td><span class="${NS}numlink" data-kind="hindranceCodes" data-level="partner" data-partner="${esc(partner)}">${total.hindranceStops || 0}</span></td>
              </tr>
            </tfoot>
          </table>
        </div>
      `;

      openModalHtml(`Touren – ${partner}`, html);
    } finally {
      setLoading(false);
    }
  }

  function metricTitle(metric) {
    switch (metric) {
      case 'deliveryStops': return 'Stopps Zustellung';
      case 'deliveryParcels': return 'Pakete Zustellung';
      case 'pickupStops': return 'Stopps Abholung';
      case 'pickupParcels': return 'Pakete Abholung';
      case 'canceledDeliveryStops': return 'Nicht bearb. Zustellstopps';
      case 'canceledPickupStops': return 'Nicht bearb. Abholstopps';
      case 'hindranceParcels': return 'Zustellhindernisse';
      default: return metric;
    }
  }

  function metricToColumns(metric) {
    if (metric === 'deliveryParcels' || metric === 'pickupParcels' || metric === 'hindranceParcels') {
      return [
        { key: 'parcel', label: 'Paketscheinnummer' },
        { key: 'systempartner', label: 'Systempartner' },
        { key: 'driver', label: 'Fahrername' },
        { key: 'tour', label: 'Tour' },
        { key: 'serviceCode', label: 'Servicecode' },
        { key: 'type', label: 'Typ' },
        { key: 'address', label: 'Adresse' },
        { key: 'status', label: 'Status' },
        { key: 'reason', label: 'Grund' },
        { key: 'time', label: metric === 'pickupParcels' ? 'Abholzeit' : 'Zustellzeit' }
      ].concat(extraColumnsForDetails());
    }

    return [
      { key: 'systempartner', label: 'Systempartner' },
      { key: 'driver', label: 'Fahrername' },
      { key: 'tour', label: 'Tour' },
      { key: 'serviceCode', label: 'Servicecode' },
      { key: 'type', label: 'Typ' },
      { key: 'address', label: 'Adresse' },
      { key: 'packages', label: 'Pakete' },
      { key: 'parcelsList', label: 'Paketscheine' },
      { key: 'status', label: 'Status' },
      { key: 'reason', label: 'Grund' },
      { key: 'time', label: (metric === 'pickupStops' || metric === 'canceledPickupStops') ? 'Abholzeit' : 'Zustellzeit' }
    ].concat(extraColumnsForDetails());
  }

  function metricToListRows(rows, metric) {
    switch (metric) {
      case 'deliveryStops':
        return rows
          .filter(r => r.__type === 'DELIVERY')
          .map(r => withExtraFields({
            systempartner: r.__partner,
            driver: r.__driver || '—',
            tour: r.__tour || '—',
            serviceCode: r.__serviceCode || '—',
            type: 'Zustellung',
            address: r.__addr,
            packages: r.__pkgCount,
            parcelsList: r.__parcelList.slice(),
            status: r.__status || '—',
            reason: cleanAdditionalCode(r.__additionalCode) || '—',
            time: formatTime(r.__deliveredAt) || '—'
          }, r));

      case 'pickupStops':
        return rows
          .filter(r => r.__type === 'PICKUP')
          .map(r => withExtraFields({
            systempartner: r.__partner,
            driver: r.__driver || '—',
            tour: r.__tour || '—',
            serviceCode: r.__serviceCode || '—',
            type: 'Abholung',
            address: r.__addr,
            packages: r.__pkgCount,
            parcelsList: r.__parcelList.slice(),
            status: r.__status || '—',
            reason: '—',
            time: formatTime(r.__pickupAt) || '—'
          }, r));

      case 'canceledDeliveryStops':
        return rows
          .filter(r => r.__type === 'DELIVERY' && isCanceledDeliveryStatus(r.__statusNorm || r.__status, r.__raw))
          .map(r => withExtraFields({
            systempartner: r.__partner,
            driver: r.__driver || '—',
            tour: r.__tour || '—',
            serviceCode: r.__serviceCode || '—',
            type: 'Zustellung',
            address: r.__addr,
            packages: r.__pkgCount,
            parcelsList: r.__parcelList.slice(),
            status: r.__status || '—',
            reason: cleanAdditionalCode(r.__additionalCode) || '—',
            time: formatTime(r.__deliveredAt) || '—'
          }, r));

      case 'canceledPickupStops':
        return rows
          .filter(r => r.__type === 'PICKUP' && isCanceledPickupStatus(r.__statusNorm || r.__status, r.__raw))
          .map(r => withExtraFields({
            systempartner: r.__partner,
            driver: r.__driver || '—',
            tour: r.__tour || '—',
            serviceCode: r.__serviceCode || '—',
            type: 'Abholung',
            address: r.__addr,
            packages: r.__pkgCount,
            parcelsList: r.__parcelList.slice(),
            status: r.__status || '—',
            reason: '—',
            time: formatTime(r.__pickupAt) || '—'
          }, r));

      case 'deliveryParcels':
        return rows
          .filter(r => r.__type === 'DELIVERY')
          .flatMap(r => {
            const psns = r.__parcelList.length ? r.__parcelList : ['—'];
            return psns.map(psn => withExtraFields({
              parcel: psn,
              systempartner: r.__partner,
              driver: r.__driver || '—',
              tour: r.__tour || '—',
              serviceCode: r.__serviceCode || '—',
              type: 'Zustellung',
              address: r.__addr,
              status: r.__status || '—',
              reason: cleanAdditionalCode(r.__additionalCode) || '—',
              time: formatTime(r.__deliveredAt) || '—'
            }, r));
          });

      case 'pickupParcels':
        return rows
          .filter(r => r.__type === 'PICKUP')
          .flatMap(r => {
            const psns = r.__parcelList.length ? r.__parcelList : ['—'];
            return psns.map(psn => withExtraFields({
              parcel: psn,
              systempartner: r.__partner,
              driver: r.__driver || '—',
              tour: r.__tour || '—',
              serviceCode: r.__serviceCode || '—',
              type: 'Abholung',
              address: r.__addr,
              status: r.__status || '—',
              reason: '—',
              time: formatTime(r.__pickupAt) || '—'
            }, r));
          });

      case 'hindranceParcels':
        return rows
          .filter(r =>
            r.__type === 'DELIVERY' &&
            isDeliveryHindranceStatus(r.__statusNorm || r.__status, r.__raw) &&
            hasUsableAdditionalCode(r)
          )
          .flatMap(r => {
            const psns = r.__parcelList.length ? r.__parcelList : ['—'];
            const code = cleanAdditionalCode(r.__additionalCode);

            return psns.map(psn => withExtraFields({
              parcel: psn,
              systempartner: r.__partner,
              driver: r.__driver || '—',
              tour: r.__tour || '—',
              serviceCode: r.__serviceCode || '—',
              type: 'Zustellhindernis',
              address: r.__addr,
              status: r.__status || '—',
              reason: code || '—',
              time: formatTime(r.__deliveredAt) || '—'
            }, r));
          });

      default:
        return [];
    }
  }

  function metricTotals(metric, rows) {
    if (metric === 'deliveryParcels' || metric === 'pickupParcels' || metric === 'hindranceParcels') {
      return {
        parcel: 'Gesamt',
        systempartner: rows.length,
        driver: '',
        tour: '',
        serviceCode: '',
        type: '',
        address: '',
        status: '',
        reason: '',
        time: '',
        ...Object.fromEntries(extraColumnsForDetails().map(c => [c.key, '']))
      };
    }

    const pkgSum = rows.reduce((sum, r) => sum + Number(r.packages || 0), 0);
    return {
      systempartner: 'Gesamt',
      driver: '',
      tour: rows.length,
      serviceCode: '',
      type: '',
      address: '',
      packages: pkgSum,
      parcelsList: '',
      status: '',
      reason: '',
      time: '',
      ...Object.fromEntries(extraColumnsForDetails().map(c => [c.key, '']))
    };
  }

  async function openMetricList({ level, partner, tour, metric }) {
    setLoading(true, 'Lade Detaildaten …');
    try {
      const rows = await fetchDetailRows(({ loadedPages, totalPages, rowsLoaded }) => {
        setLoading(true, `Lade Detaildaten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
      }, false);

      let baseRows = rows;
      if (level === 'partner') {
        baseRows = baseRows.filter(r => r.__partner === partner);
      } else if (level === 'tour') {
        baseRows = baseRows.filter(r => r.__partner === partner && String(r.__tour) === String(tour));
      }

      const titleParts = [];
      if (partner) titleParts.push(partner);
      if (tour) titleParts.push(`Tour ${tour}`);
      titleParts.push(metricTitle(metric));

      const listRows = metricToListRows(baseRows, metric);
      const columns = metricToColumns(metric);
      const totals = metricTotals(metric, listRows);

      openModalTable(titleParts.join(' – '), listRows, columns, totals);
    } finally {
      setLoading(false);
    }
  }

  async function openHindranceCodeSummary({ level, partner, tour }) {
    setLoading(true, 'Lade Zustellhindernisse …');
    try {
      const rows = await fetchDetailRows(({ loadedPages, totalPages, rowsLoaded }) => {
        setLoading(true, `Lade Detaildaten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
      }, false);

      let baseRows = rows.filter(r =>
        r.__type === 'DELIVERY' &&
        isDeliveryHindranceStatus(r.__statusNorm || r.__status, r.__raw) &&
        hasUsableAdditionalCode(r)
      );

      if (level === 'partner') {
        baseRows = baseRows.filter(r => r.__partner === partner);
      } else if (level === 'tour') {
        baseRows = baseRows.filter(r => r.__partner === partner && String(r.__tour) === String(tour));
      }

      const groups = new Map();

      for (const r of baseRows) {
        const code = cleanAdditionalCode(r.__additionalCode);
        const key = `${r.__partner}|||${r.__driver || ''}|||${r.__tour}|||${code}`;

        if (!groups.has(key)) {
          groups.set(key, {
            systempartner: r.__partner,
            driver: r.__driver || '—',
            tour: r.__tour || '—',
            additionalCode: code,
            count: 0
          });
        }

        groups.get(key).count += hindranceItemCount(r);
      }

      const listRows = Array.from(groups.values()).sort((a, b) => {
        const c1 = collator.compare(a.systempartner, b.systempartner);
        if (c1 !== 0) return c1;
        const c2 = collator.compare(String(a.tour || ''), String(b.tour || ''));
        if (c2 !== 0) return c2;
        return collator.compare(a.additionalCode, b.additionalCode);
      });

      const titleParts = [];
      if (partner) titleParts.push(partner);
      if (tour) titleParts.push(`Tour ${tour}`);
      titleParts.push('Zustellhindernisse');

      const html = `
        <div style="overflow:auto">
          <table class="${NS}tbl" id="${NS}hindrance-code-table" style="min-width:1100px">
            <thead>
              <tr>
                <th>Systempartner</th>
                <th>Fahrername</th>
                <th>Tour</th>
                <th>Grund</th>
                <th>Anzahl</th>
              </tr>
            </thead>
            <tbody>
              ${listRows.map(r => `
                <tr>
                  <td>${esc(r.systempartner)}</td>
                  <td class="${NS}driver-cell" title="${esc(r.driver || '—')}">${esc(r.driver || '—')}</td>
                  <td>${esc(r.tour)}</td>
                  <td>${esc(r.additionalCode)}</td>
                  <td>
                    <span class="${NS}numlink"
                          data-kind="hindranceCodeDetails"
                          data-partner="${esc(r.systempartner)}"
                          data-tour="${esc(r.tour)}"
                          data-code="${esc(r.additionalCode)}">${r.count}</span>
                  </td>
                </tr>
              `).join('')}
            </tbody>
            <tfoot>
              <tr>
                <td>Gesamt</td>
                <td></td>
                <td></td>
                <td></td>
                <td>${listRows.reduce((s, r) => s + r.count, 0)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      `;

      openModalHtml(titleParts.join(' – '), html);
    } finally {
      setLoading(false);
    }
  }

  async function openHindranceCodeDetails({ partner, tour, code }) {
    setLoading(true, 'Lade Zustellhindernis-Details …');
    try {
      const rows = await fetchDetailRows(({ loadedPages, totalPages, rowsLoaded }) => {
        setLoading(true, `Lade Detaildaten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
      }, false);

      const wantedCode = cleanAdditionalCode(code);

      const baseRows = rows.filter(r =>
        r.__type === 'DELIVERY' &&
        isDeliveryHindranceStatus(r.__statusNorm || r.__status, r.__raw) &&
        r.__partner === partner &&
        String(r.__tour || '—') === String(tour || '—') &&
        cleanAdditionalCode(r.__additionalCode) === wantedCode
      );

      const listRows = baseRows.flatMap(r => {
        const psns = r.__parcelList.length ? r.__parcelList : ['—'];
        const realCode = cleanAdditionalCode(r.__additionalCode);

        return psns.map(psn => ({
          parcel: psn,
          systempartner: r.__partner,
          driver: r.__driver || '—',
          tour: r.__tour || '—',
          serviceCode: r.__serviceCode || '—',
          additionalCode: realCode || '—',
          address: r.__addr,
          status: r.__status || '—',
          time: formatTime(r.__deliveredAt) || '—'
        }));
      });

      const html = `
        <div style="overflow:auto">
          <table class="${NS}tbl" id="${NS}hindrance-detail-table" style="min-width:1500px">
            <thead>
              <tr>
                <th>Paketscheinnummer</th>
                <th>Systempartner</th>
                <th>Fahrername</th>
                <th>Tour</th>
                <th>Servicecode</th>
                <th>Grund</th>
                <th>Adresse</th>
                <th>Status</th>
                <th>Zeit</th>
              </tr>
            </thead>
            <tbody>
              ${listRows.map(r => `
                <tr>
                  <td>${r.parcel !== '—' ? `<button class="${NS}psn-btn" data-psn="${esc(String(r.parcel))}">${esc(String(r.parcel))}</button>` : '—'}</td>
                  <td>${esc(r.systempartner)}</td>
                  <td class="${NS}driver-cell" title="${esc(r.driver || '—')}">${esc(r.driver || '—')}</td>
                  <td>${esc(r.tour)}</td>
                  <td>${formatModalCell('serviceCode', r.serviceCode)}</td>
                  <td>${esc(r.additionalCode)}</td>
                  <td>${esc(r.address)}</td>
                  <td>${formatModalCell('status', r.status)}</td>
                  <td>${esc(r.time)}</td>
                </tr>
              `).join('')}
            </tbody>
            <tfoot>
              <tr>
                <td>Gesamt</td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td>${listRows.length}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      `;

      openModalHtml(`Zustellhindernisse – ${partner} – Tour ${tour} – Grund ${wantedCode}`, html);
    } finally {
      setLoading(false);
    }
  }


  function openColumnChooser() {
    const selected = new Set(getSelectedExtraColumns());
    const html = `
      <div style="padding:12px;max-height:65vh;overflow:auto">
        <div style="font:700 13px system-ui;margin-bottom:8px">Zusätzliche Detailspalten auswählen</div>
        <div style="font:600 12px system-ui;opacity:.75;margin-bottom:10px">
          Standard bleibt unverändert. Die Auswahl wird nur zusätzlich in den Detailtabellen angezeigt.
        </div>
        <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:6px 14px">
          ${FIELD_CATALOG.map(f => `
            <label style="display:flex;gap:7px;align-items:center;font:12px system-ui">
              <input type="checkbox" data-extra-column="${esc(f.key)}" ${selected.has(f.key) ? 'checked' : ''}>
              <span>${esc(f.label)}</span>
            </label>
          `).join('')}
        </div>
        <div style="margin-top:14px;display:flex;gap:8px;justify-content:flex-end">
          <button class="${NS}btn" data-action="resetColumns">Standard</button>
          <button class="${NS}btn" data-action="saveColumns">Übernehmen</button>
        </div>
      </div>
    `;
    openModalHtml('Spaltenauswahl', html);
  }

  function saveColumnChooser() {
    const boxes = Array.from(document.querySelectorAll('#' + NS + 'modal input[data-extra-column]'));
    setSelectedExtraColumns(boxes.filter(b => b.checked).map(b => b.dataset.extraColumn));
    hideModal();
  }

  function openModalHtml(title, html) {
    const m = document.getElementById(NS + 'modal');
    const t = document.getElementById(NS + 'modal-title');
    const b = document.getElementById(NS + 'modal-body');
    if (t) t.textContent = title || '';
    if (b) b.innerHTML = `<div class="${NS}modal-scroll">${html || ''}</div>`;
    state.modal = { type: 'html', title, rows: [], columns: [], totals: null, sortCol: -1, sortDir: 'asc' };
    if (m) m.style.display = 'flex';
  }

  function openModalTable(title, rows, columns, totals) {
    state.modal = {
      type: 'table',
      title: title || '',
      rows: rows.slice(),
      columns: columns.slice(),
      totals: totals || null,
      sortCol: -1,
      sortDir: 'asc'
    };
    renderModalTable();
    const m = document.getElementById(NS + 'modal');
    if (m) m.style.display = 'flex';
  }

  function getStatusBadgeClass(value) {
    const s = normalizeStatusForMatch(value);
    if (!s) return 'gray';
    if (s.includes('AUTOMATISCH') || s.includes('STORNIERT') || s.includes('CANCEL') || s.includes('FAILED') || s.includes('PROBLEM')) return 'red';
    if (s.includes('ZUGESTELLT') || s.includes('DELIVERED') || s.includes('ABGEHOLT') || s.includes('PICKED UP') || s.includes('COMPLETED')) return 'green';
    if (s.includes('UNTERWEGS') || s.includes('IN TOUR') || s.includes('ON ROAD')) return 'blue';
    if (s.includes('OFFEN') || s.includes('PENDING') || s.includes('GEPLANT')) return 'yellow';
    if (s.includes('RETOUR') || s.includes('RETURN')) return 'purple';
    return 'gray';
  }

  function renderServiceCodeBadge(value) {
    let codes = [];
    let kind = '';

    if (Array.isArray(value)) {
      codes = value.map(v => String(v || '').trim()).filter(Boolean);
    } else if (value && typeof value === 'object') {
      codes = Array.isArray(value.codes)
        ? value.codes.map(v => String(v || '').trim()).filter(Boolean)
        : String(value.codes || '')
            .split(/[^\dA-Za-z]+/)
            .map(s => s.trim())
            .filter(Boolean);
      kind = String(value.kind || '').trim().toUpperCase();
    } else {
      codes = String(value || '')
        .split(/[^\dA-Za-z]+/)
        .map(s => s.trim())
        .filter(Boolean);
    }

    if (!codes.length) {
      return `<span style="
        display:inline-block;
        max-width:100%;
        padding:2px 8px;
        border-radius:999px;
        font:700 11px system-ui;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
        background:#f3f4f6;
        color:#374151;
        border:1px solid #d1d5db;
      ">—</span>`;
    }

    let bg = '#f3f4f6';
    let fg = '#374151';
    let border = '#d1d5db';

    if (kind === 'EXPRESS') {
      bg = '#fee2e2';
      fg = '#991b1b';
      border = '#fca5a5';
    } else if (kind === 'PRIO') {
      bg = '#fef2f2';
      fg = '#b91c1c';
      border = '#fecaca';
    }

    return `<div style="display:flex;flex-wrap:wrap;gap:4px;justify-content:flex-start;align-items:flex-start;">${
      codes.map(code => `<span style="
        display:inline-block;
        padding:2px 8px;
        border-radius:999px;
        font:700 11px system-ui;
        white-space:nowrap;
        background:${bg};
        color:${fg};
        border:1px solid ${border};
      ">${esc(code)}</span>`).join('')
    }</div>`;
  }

  async function openTourAllList({ partner, tour }) {
    setLoading(true, `Lade komplette Tour ${tour} …`);
    try {
      const rows = await fetchDetailRows(({ loadedPages, totalPages, rowsLoaded }) => {
        setLoading(true, `Lade Detaildaten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
      }, false);

      const baseRows = rows.filter(r =>
        r.__partner === partner &&
        String(r.__tour || '—') === String(tour || '—')
      );

      const deliveryDetailCache = window.__spxDeliveryDetailCache || (window.__spxDeliveryDetailCache = new Map());
      const headers = buildHeaders(lastOkRequest?.headers || {});
      const origin = (lastOkRequest?.url?.origin) || location.origin;

      async function fetchDeliveryDetail(deliveryId) {
        const key = String(deliveryId || '').trim();
        if (!key) return null;
        if (deliveryDetailCache.has(key)) return deliveryDetailCache.get(key);

        try {
          const url = `${origin}/dispatcher/api/delivery/${encodeURIComponent(key)}`;
          const res = await fetch(url, { credentials: 'include', headers, cache: 'no-store' });
          if (!res.ok) {
            deliveryDetailCache.set(key, null);
            return null;
          }
          const json = await res.json();
          deliveryDetailCache.set(key, json || null);
          return json || null;
        } catch {
          deliveryDetailCache.set(key, null);
          return null;
        }
      }

      function extractCodesLikePrioExpressScript(obj) {
        if (!obj) return [];
        const set = new Set();

        const addFromVal = v => {
          if (v == null) return;
          String(v)
            .split(/[^\dA-Za-z]+/)
            .map(s => s.trim())
            .filter(Boolean)
            .forEach(code => set.add(code));
        };

        const addFromArr = arr => {
          if (!Array.isArray(arr)) return;
          arr.forEach(addFromVal);
        };

        addFromVal(obj.serviceCode);
        addFromVal(obj.servicecode);
        addFromVal(obj.service_code);
        addFromArr(obj.serviceCodes);

        if (obj.service && typeof obj.service === 'object') {
          addFromVal(obj.service.code);
          addFromVal(obj.service.serviceCode);
          addFromVal(obj.service.id);
          addFromArr(obj.service.serviceCodes);
        }

        if (obj.product && typeof obj.product === 'object') {
          addFromVal(obj.product.serviceCode);
          addFromVal(obj.product.code);
          addFromVal(obj.product.id);
          addFromArr(obj.product.serviceCodes);
        }

        return Array.from(set).sort((a, b) => collator.compare(a, b));
      }

      function parcelKindFromDetailParcel(p) {
        const text = [
          p?.priority,
          p?.prio,
          p?.serviceName,
          p?.serviceLabel,
          p?.productName,
          p?.productLabel,
          p?.service?.name,
          p?.service?.label,
          p?.product?.name,
          p?.product?.label,
          Array.isArray(p?.elements) ? p.elements.join(' ') : p?.elements
        ].filter(Boolean).join(' | ');

        return serviceCodeKind(text);
      }

      function parcelServiceDataFromDetail(detail) {
        const map = new Map();
        const parcels = Array.isArray(detail?.parcels) ? detail.parcels : [];

        for (const p of parcels) {
          let psn = String(
            p?.parcelNumber ||
            p?.parcel_number ||
            p?.shipmentNumber ||
            p?.shipment_number ||
            p?.barcode ||
            ''
          ).replace(/\D+/g, '');

          if (!psn) continue;
          if (psn.length === 13) psn = '0' + psn;

          map.set(psn, {
            codes: extractCodesLikePrioExpressScript(p),
            kind: parcelKindFromDetailParcel(p)
          });
        }

        return map;
      }

      const detailMaps = new Map();
      const idsToLoad = Array.from(new Set(
        baseRows
          .map(r => r.__raw?.id ?? r.__raw?.deliveryId ?? r.__raw?.stopId ?? '')
          .map(v => String(v || '').trim())
          .filter(Boolean)
      ));

      for (const id of idsToLoad) {
        const detail = await fetchDeliveryDetail(id);
        detailMaps.set(id, parcelServiceDataFromDetail(detail));
      }

      const listRows = baseRows.flatMap(r => {
        const rowId = String(r.__raw?.id ?? r.__raw?.deliveryId ?? r.__raw?.stopId ?? '').trim();
        const svcMap = detailMaps.get(rowId) || new Map();
        const psns = r.__parcelList.length ? r.__parcelList : ['—'];

        if (psns.length === 1 && psns[0] === '—') {
          return [{
            systempartner: r.__partner,
            driver: r.__driver || '—',
            tour: r.__tour || '—',
            serviceCode: {
              codes: String(r.__serviceCode || '')
                .split(/[^\dA-Za-z]+/)
                .map(s => s.trim())
                .filter(Boolean),
              kind: ''
            },
            type: r.__type === 'PICKUP' ? 'Abholung' : 'Zustellung',
            address: r.__addr,
            packages: r.__pkgCount,
            parcelsList: [],
            status: r.__status || '—',
            reason: cleanAdditionalCode(r.__additionalCode) || '—',
            time: r.__type === 'PICKUP'
              ? (formatTime(r.__pickupAt) || '—')
              : (formatTime(r.__deliveredAt) || '—')
          }];
        }

        return psns.map(psn => {
          let cleanPsn = String(psn || '').replace(/\D+/g, '');
          if (cleanPsn.length === 13) cleanPsn = '0' + cleanPsn;

          const packetSvc = svcMap.get(cleanPsn) || {
            codes: String(r.__serviceCode || '')
              .split(/[^\dA-Za-z]+/)
              .map(s => s.trim())
              .filter(Boolean),
            kind: ''
          };

          return {
            systempartner: r.__partner,
            driver: r.__driver || '—',
            tour: r.__tour || '—',
            serviceCode: packetSvc,
            type: r.__type === 'PICKUP' ? 'Abholung' : 'Zustellung',
            address: r.__addr,
            packages: 1,
            parcelsList: [psn],
            status: r.__status || '—',
            reason: cleanAdditionalCode(r.__additionalCode) || '—',
            time: r.__type === 'PICKUP'
              ? (formatTime(r.__pickupAt) || '—')
              : (formatTime(r.__deliveredAt) || '—')
          };
        });
      });

      const columns = [
        { key: 'systempartner', label: 'Systempartner' },
        { key: 'driver', label: 'Fahrername' },
        { key: 'tour', label: 'Tour' },
        { key: 'serviceCode', label: 'Servicecode' },
        { key: 'type', label: 'Typ' },
        { key: 'address', label: 'Adresse' },
        { key: 'packages', label: 'Pakete' },
        { key: 'parcelsList', label: 'Paketscheine' },
        { key: 'status', label: 'Status' },
        { key: 'reason', label: 'Grund' },
        { key: 'time', label: 'Zeit' }
      ];

      const totals = {
        systempartner: 'Gesamt',
        driver: '',
        tour: '',
        serviceCode: '',
        type: '',
        address: '',
        packages: listRows.reduce((sum, r) => sum + Number(r.packages || 0), 0),
        parcelsList: '',
        status: '',
        reason: '',
        time: listRows.length
      };

      openModalTable(`Komplette Tour – ${partner} – Tour ${tour}`, listRows, columns, totals);
    } finally {
      setLoading(false);
    }
  }

  function formatModalCell(key, value) {
    if (key === 'parcel' && value && value !== '—') {
      return `<button class="${NS}psn-btn" data-psn="${esc(String(value))}">${esc(String(value))}</button>`;
    }

    if (key === 'parcelsList') {
      const arr = Array.isArray(value) ? value.filter(Boolean) : [];
      if (!arr.length) return '—';
      return `<div class="${NS}psn-wrap">${arr.map(psn => `<button class="${NS}psn-btn" data-psn="${esc(String(psn))}">${esc(String(psn))}</button>`).join('')}</div>`;
    }

    if (key === 'status') {
      const txt = String(value || '—');
      return `<span class="${NS}status ${getStatusBadgeClass(txt)}">${esc(txt)}</span>`;
    }

    if (key === 'serviceCode') {
      return renderServiceCodeBadge(value || '—');
    }

    return esc(value ?? '');
  }

    function renderModalTable() {
    const t = document.getElementById(NS + 'modal-title');
    const b = document.getElementById(NS + 'modal-body');
    if (t) t.textContent = state.modal.title || '';
    if (!b) return;

    const cols = state.modal.columns || [];
    const rows = state.modal.rows || [];
    const totals = state.modal.totals || null;

    b.innerHTML = `
      <div class="${NS}modal-scroll">
        <table class="${NS}tbl" id="${NS}modal-table">
          <thead>
            <tr>${cols.map((c, i) => `<th data-col="${i}">${esc(c.label)}</th>`).join('')}</tr>
          </thead>
          <tbody>
            ${rows.map(r => `<tr>${cols.map(c => `<td class="${c.key === 'driver' ? NS + 'driver-cell' : ''}" title="${c.key === 'driver' ? esc(String(r[c.key] || '')) : ''}">${formatModalCell(c.key, r[c.key])}</td>`).join('')}</tr>`).join('')}
          </tbody>
          ${totals ? `<tfoot><tr>${cols.map(c => `<td class="${c.key === 'driver' ? NS + 'driver-cell' : ''}" title="${c.key === 'driver' ? esc(String(totals[c.key] || '')) : ''}">${formatModalCell(c.key, totals[c.key])}</td>`).join('')}</tr></tfoot>` : ''}
        </table>
      </div>
    `;

    if (state.modal.sortCol >= 0) {
      const th = b.querySelector(`th[data-col="${state.modal.sortCol}"]`);
      if (th) th.classList.add(state.modal.sortDir === 'asc' ? NS + 'sort-asc' : NS + 'sort-desc');
    }
  }

  function sortModalByColumn(colIndex) {
    if (state.modal.type !== 'table') return;
    const cols = state.modal.columns || [];
    const key = cols[colIndex]?.key;
    if (!key) return;

    const asc = state.modal.sortCol === colIndex ? state.modal.sortDir !== 'asc' : true;
    state.modal.sortCol = colIndex;
    state.modal.sortDir = asc ? 'asc' : 'desc';

    state.modal.rows.sort((a, b) => {
      const A = a[key] ?? '';
      const B = b[key] ?? '';

      if (Array.isArray(A) || Array.isArray(B)) {
        const sa = Array.isArray(A) ? A.join(', ') : String(A);
        const sb = Array.isArray(B) ? B.join(', ') : String(B);
        return asc ? collator.compare(sa, sb) : collator.compare(sb, sa);
      }

      const nA = Number(A);
      const nB = Number(B);
      if (!Number.isNaN(nA) && !Number.isNaN(nB) && String(A).trim() !== '' && String(B).trim() !== '') {
        return asc ? (nA - nB) : (nB - nA);
      }

      return asc ? collator.compare(String(A), String(B)) : collator.compare(String(B), String(A));
    });

    renderModalTable();
  }

  function hideModal() {
    const m = document.getElementById(NS + 'modal');
    if (m) m.style.display = 'none';
  }

  function parseSortableCellValue(text) {
    const s = norm(String(text || ''));
    if (!s || s === '—') return { type: 'text', value: '' };

    const normalized = s.replace(/\./g, '').replace(',', '.');
    if (/^-?\d+(?:\.\d+)?$/.test(normalized)) {
      const n = Number(normalized);
      if (!Number.isNaN(n)) return { type: 'number', value: n };
    }

    return { type: 'text', value: s.toLowerCase() };
  }

  function compareSortableValues(a, b, asc) {
    if (a.type === 'number' && b.type === 'number') {
      return asc ? (a.value - b.value) : (b.value - a.value);
    }
    const r = collator.compare(String(a.value || ''), String(b.value || ''));
    return asc ? r : -r;
  }

  function sortDomTableByColumn(table, colIndex, asc) {
    if (!table) return;

    const tbody = table.querySelector('tbody');
    if (!tbody) return;

    const rows = Array.from(tbody.querySelectorAll('tr'));
    rows.sort((ra, rb) => {
      const ta = ra.children[colIndex]?.textContent || '';
      const tb = rb.children[colIndex]?.textContent || '';
      const va = parseSortableCellValue(ta);
      const vb = parseSortableCellValue(tb);
      return compareSortableValues(va, vb, asc);
    });

    rows.forEach(tr => tbody.appendChild(tr));
  }

  function clearTableSortMarkers(table) {
    table.querySelectorAll('thead th').forEach(th => {
      th.classList.remove(NS + 'sort-asc', NS + 'sort-desc');
      delete th.dataset.sortDir;
    });
  }

  function bindSortableTable(table) {
    if (!table || table.dataset.spxSortableBound === '1') return;
    table.dataset.spxSortableBound = '1';

    const ths = Array.from(table.querySelectorAll('thead th'));
    if (!ths.length) return;

    ths.forEach((th, colIndex) => {
      th.style.cursor = 'pointer';

      th.addEventListener('click', (ev) => {
        ev.preventDefault();
        ev.stopPropagation();

        const asc = th.dataset.sortDir !== 'asc';

        clearTableSortMarkers(table);
        th.dataset.sortDir = asc ? 'asc' : 'desc';
        th.classList.add(asc ? NS + 'sort-asc' : NS + 'sort-desc');

        sortDomTableByColumn(table, colIndex, asc);
      });
    });
  }

  function bindAllSortableTables(root = document) {
    const scopes = [
      document.getElementById(NS + 'panel'),
      document.getElementById(NS + 'modal'),
      root
    ].filter(Boolean);

    for (const scope of scopes) {
      const tables = scope.querySelectorAll('table');
      tables.forEach(tbl => bindSortableTable(tbl));
    }
  }

  (function installGlobalTableSorting() {
    if (window.__spx_sorting_installed) return;
    window.__spx_sorting_installed = true;

    const observer = new MutationObserver(() => {
      bindAllSortableTables(document);
    });

    const start = () => {
      bindAllSortableTables(document);
      observer.observe(document.documentElement || document.body, {
        childList: true,
        subtree: true
      });
    };

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', start, { once: true });
    } else {
      start();
    }
  })();

  function openParcelTracking(psn) {
    const id = String(psn || '').replace(/\D+/g, '');
    if (!id) return;
    window.open(`https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${id}`, '_blank', 'noopener');
  }

  function tableToTSV(table) {
    const rows = Array.from(table.querySelectorAll('tr'));
    return rows.map(tr => Array.from(tr.children).map(td => norm(td.textContent || '')).join('\t')).join('\n');
  }

  function cloneTableWithInlineStyles(table) {
    const clone = table.cloneNode(true);
    const origEls = [table, ...table.querySelectorAll('*')];
    const cloneEls = [clone, ...clone.querySelectorAll('*')];

    for (let i = 0; i < origEls.length; i++) {
      const src = origEls[i];
      const dst = cloneEls[i];
      if (!src || !dst) continue;

      const cs = getComputedStyle(src);
      const styleProps = [
        'font-family', 'font-size', 'font-weight', 'color', 'background-color',
        'border-top', 'border-right', 'border-bottom', 'border-left',
        'padding-top', 'padding-right', 'padding-bottom', 'padding-left',
        'text-align', 'vertical-align', 'white-space', 'width', 'min-width',
        'max-width', 'text-decoration', 'border-radius'
      ];

      let styleText = '';
      for (const prop of styleProps) {
        const val = cs.getPropertyValue(prop);
        if (val) styleText += `${prop}:${val};`;
      }
      dst.setAttribute('style', styleText);
    }

    return clone;
  }

  function tableToStyledHtml(table) {
    const clone = cloneTableWithInlineStyles(table);
    const all = [clone, ...clone.querySelectorAll('*')];

    all.forEach(el => {
      const tag = el.tagName;
      if (tag === 'TABLE') {
        el.style.borderCollapse = 'separate';
        el.style.borderSpacing = '0';
        el.style.tableLayout = 'fixed';
        el.style.width = '100%';
        el.style.fontFamily = 'Segoe UI, Arial, sans-serif';
        el.style.fontSize = '12px';
        el.style.color = '#111827';
      }

      if (tag === 'TH' || tag === 'TD') {
        el.style.border = '1px solid #d1d5db';
        el.style.padding = '6px 8px';
        el.style.verticalAlign = 'top';
        el.style.whiteSpace = 'nowrap';
      }

      if (tag === 'TH') {
        el.style.backgroundColor = '#ead0d0';
        el.style.color = '#8a1f1f';
        el.style.fontWeight = '700';
        el.style.textAlign = el.cellIndex === 0 ? 'left' : 'right';
      }

      if (tag === 'TD') {
        const parent = el.parentElement;
        const section = parent?.parentElement?.tagName;
        if (section === 'TFOOT') {
          el.style.backgroundColor = '#d7e6f3';
          el.style.color = '#0b3d6b';
          el.style.fontWeight = '700';
        } else {
          el.style.backgroundColor = '#ffffff';
        }
      }
    });

    return `
      <html>
        <head><meta charset="utf-8"></head>
        <body style="margin:0;padding:12px;background:#ffffff;">
          <div style="font-family:Segoe UI, Arial, sans-serif;font-size:12px;color:#111827;">
            ${clone.outerHTML}
          </div>
        </body>
      </html>
    `;
  }

  async function copyHtmlAndText(html, text) {
    try {
      if (navigator.clipboard && window.ClipboardItem) {
        const item = new ClipboardItem({
          'text/html': new Blob([html], { type: 'text/html' }),
          'text/plain': new Blob([text], { type: 'text/plain' })
        });
        await navigator.clipboard.write([item]);
        return true;
      }
    } catch {}
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      return false;
    }
  }

  async function copyMainTable() {
    const table = document.getElementById(NS + 'main-table');
    if (!table) return;
    const ok = await copyHtmlAndText(tableToStyledHtml(table), tableToTSV(table));
    if (!ok) alert('Kopieren nicht möglich.');
  }

  async function copyModalTable() {
    const table = document.querySelector('#' + NS + 'modal-body table');
    if (!table) return;
    const ok = await copyHtmlAndText(tableToStyledHtml(table), tableToTSV(table));
    if (!ok) alert('Kopieren nicht möglich.');
  }

  async function refreshSummary({ forceReload = false } = {}) {
    if (!lastOkRequest) throw new Error('Bitte zuerst einmal die normale Pickup-/Delivery-Liste laden.');

    const cacheKey = getSummaryCacheKey();

    if (forceReload) {
      clearRuntimeCaches();
      await loadTourPartnerMap(true);
    }

    if (!forceReload) {
      const cached = getCachedSummary(cacheKey);
      if (cached) {
        state.summaryRows = cached.summaryRows;
        state.standText = cached.standText;
        renderStand();
        renderMainTable();
        return;
      }
    }

    const result = await fetchSummaryOnly(({ loadedPages, totalPages, rowsLoaded }) => {
      setLoading(true, `Lade Daten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
    });

    state.summaryRows = result.summaryRows;
    setLoadedStandFromDateRange(
      getSelectedDateFrom() || state.detectedDateFrom || state.detectedDateValue,
      getSelectedDateTo() || state.detectedDateTo || state.detectedDateValue
    );

    setCachedSummary(cacheKey, {
      summaryRows: result.summaryRows,
      standText: state.standText
    });

    setCachedDetailRows(cacheKey, result.normalized);
    renderMainTable();
  }

  async function fullRefresh({ forceReload = false } = {}) {
    if (isBusy) return;
    try {
      isBusy = true;
      setLoading(true, 'Lade Daten …');
      await refreshSummary({ forceReload });
    } catch (e) {
      console.error(e);
      alert(String(e?.message || e));
    } finally {
      setLoading(false);
      isBusy = false;
    }
  }

  function findPickupDeliveryTab() {
    const selectors = ['button', '[role="tab"]', '.tab', '.tabs__tab', '.mat-tab-label', 'a', 'div'];
    for (const sel of selectors) {
      const els = Array.from(document.querySelectorAll(sel));
      const hit = els.find(el => {
        const txt = norm(el.textContent || '').toLowerCase();
        return txt.includes('abholung') && txt.includes('zustellung');
      });
      if (hit) return hit;
    }
    return null;
  }

  function isPickupDeliveryTabActive(el) {
    if (!el) return false;
    const txt = norm(el.textContent || '').toLowerCase();
    if (!txt.includes('abholung') || !txt.includes('zustellung')) return false;

    const cls = String(el.className || '').toLowerCase();
    const ariaSelected = String(el.getAttribute('aria-selected') || '').toLowerCase();
    const ariaCurrent = String(el.getAttribute('aria-current') || '').toLowerCase();
    const cs = getComputedStyle(el);

    if (ariaSelected === 'true' || ariaCurrent === 'page') return true;
    if (cls.includes('active') || cls.includes('selected')) return true;

    return (
      cs.backgroundColor === 'rgb(226, 0, 52)' ||
      cs.backgroundColor === 'rgb(229, 0, 55)' ||
      cs.color === 'rgb(255, 255, 255)'
    );
  }

  function clickElement(el) {
    if (!el) return false;
    try {
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
      el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
      el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      return true;
    } catch {
      try {
        el.click();
        return true;
      } catch {
        return false;
      }
    }
  }

  function autoSwitchToPickupDeliveryTab() {
    let tries = 0;
    const maxTries = 20;

    const run = () => {
      tries++;
      const tab = findPickupDeliveryTab();
      if (tab) {
        if (!isPickupDeliveryTabActive(tab)) clickElement(tab);
        if (isPickupDeliveryTabActive(tab) || tries >= maxTries) return;
      }
      if (tries < maxTries) setTimeout(run, 350);
    };

    run();
  }

  (function hookNetwork() {
    if (!window.__spx_fetch_hooked && window.fetch) {
      const orig = window.fetch;
      window.fetch = async function (input, init = {}) {
        const res = await orig(input, init);
        try {
          const urlStr = typeof input === 'string' ? input : (input && input.url) || '';
          if (urlStr.includes('/dispatcher/api/pickup-delivery') && res.ok) {
            const u = new URL(urlStr, location.origin);
            const q = u.searchParams;
            if (!q.get('parcelNumber') && !q.get('parcel_number')) {
              const headers = {};
              const src = (init && init.headers) || (input && input.headers);
              if (src) {
                if (src.forEach) src.forEach((v, k) => headers[String(k).toLowerCase()] = String(v));
                else if (Array.isArray(src)) src.forEach(([k, v]) => headers[String(k).toLowerCase()] = String(v));
                else Object.entries(src).forEach(([k, v]) => headers[String(k).toLowerCase()] = String(v));
              }
              if (!headers['authorization']) {
                const m = document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
                if (m) headers['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
              }
              lastOkRequest = { url: u, headers };
              const dd = detectDateParamsFromUrl(u);
              state.detectedDateParam = dd.key;
              state.detectedDateValue = dd.from || toIsoDateLocal(new Date());
              state.detectedDateFrom = dd.from || toIsoDateLocal(new Date());
              state.detectedDateTo = dd.to || dd.from || toIsoDateLocal(new Date());
              renderDateOptions();
              const note = document.getElementById(NS + 'note');
              if (note) note.style.display = 'none';
            }
          }
        } catch {}
        return res;
      };
      window.__spx_fetch_hooked = true;
    }

    if (!window.__spx_xhr_hooked && window.XMLHttpRequest) {
      const X = window.XMLHttpRequest;
      const open = X.prototype.open;
      const send = X.prototype.send;
      const setH = X.prototype.setRequestHeader;

      X.prototype.open = function (method, url) {
        this.__spx_url = typeof url === 'string' ? new URL(url, location.origin) : null;
        this.__spx_headers = {};
        return open.apply(this, arguments);
      };
      X.prototype.setRequestHeader = function (k, v) {
        try { this.__spx_headers[String(k).toLowerCase()] = String(v); } catch {}
        return setH.apply(this, arguments);
      };
      X.prototype.send = function () {
        const onload = () => {
          try {
            if (this.__spx_url && this.__spx_url.href.includes('/dispatcher/api/pickup-delivery') && this.status >= 200 && this.status < 300) {
              const q = this.__spx_url.searchParams;
              if (!q.get('parcelNumber') && !q.get('parcel_number')) {
                if (!this.__spx_headers['authorization']) {
                  const m = document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
                  if (m) this.__spx_headers['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
                }
                lastOkRequest = { url: this.__spx_url, headers: this.__spx_headers };
                const dd = detectDateParamsFromUrl(this.__spx_url);
                state.detectedDateParam = dd.key;
                state.detectedDateValue = dd.from || toIsoDateLocal(new Date());
                state.detectedDateFrom = dd.from || toIsoDateLocal(new Date());
                state.detectedDateTo = dd.to || dd.from || toIsoDateLocal(new Date());
                renderDateOptions();
                const note = document.getElementById(NS + 'note');
                if (note) note.style.display = 'none';
              }
            }
          } catch {}
          this.removeEventListener('load', onload);
        };
        this.addEventListener('load', onload);
        return send.apply(this, arguments);
      };
      window.__spx_xhr_hooked = true;
    }
  })();

  function startModuleOnce() {
    autoSwitchToPickupDeliveryTab();
    if (started) {
      togglePanel(true);
      return;
    }
    started = true;
    mountUI();
    togglePanel(true);
    autoSwitchToPickupDeliveryTab();
  }
function recipientNameOf(r) {
  const candidates = [
    r?.name,
    r?.recipientName,
    r?.receiverName,
    r?.consigneeName,
    r?.customerName,
    r?.contactName,
    r?.fullName,
    r?.displayName,
    r?.companyName,
    r?.company,
    r?.organisationName,
    r?.organizationName,
    r?.personName,

    r?.recipient?.name,
    r?.recipient?.fullName,
    r?.recipient?.displayName,
    r?.recipient?.companyName,
    r?.recipient?.company,
    r?.recipient?.organisationName,
    r?.recipient?.organizationName,
    r?.recipient?.contactName,

    r?.address?.name,
    r?.address?.fullName,
    r?.address?.displayName,
    r?.address?.companyName,
    r?.address?.company,
    r?.address?.organisationName,
    r?.address?.organizationName,
    r?.address?.contactName,

    r?.consignee?.name,
    r?.consignee?.fullName,
    r?.consignee?.displayName,
    r?.consignee?.companyName,
    r?.consignee?.company,
    r?.consignee?.contactName,

    r?.customer?.name,
    r?.customer?.fullName,
    r?.customer?.displayName,
    r?.customer?.companyName,
    r?.customer?.company,
    r?.customer?.contactName
  ];

  for (const c of candidates) {
    const s = norm(c);
    if (s) return s;
  }

  return '';
}

addrOf = function (r) {
  const name = recipientNameOf(r);

  const street = norm(
    r?.street ||
    r?.addressLine1 ||
    r?.address ||
    r?.address?.street ||
    r?.recipient?.street ||
    r?.consignee?.street ||
    r?.customer?.street ||
    ''
  );

  const house = norm(
    r?.houseno ||
    r?.houseNo ||
    r?.houseNumber ||
    r?.address?.houseNumber ||
    r?.recipient?.houseNumber ||
    r?.consignee?.houseNumber ||
    r?.customer?.houseNumber ||
    ''
  );

  const postal = norm(
    r?.postalCode ||
    r?.zipCode ||
    r?.zip ||
    r?.postal_code ||
    r?.address?.postalCode ||
    r?.recipient?.postalCode ||
    r?.consignee?.postalCode ||
    r?.customer?.postalCode ||
    ''
  );

  const city = norm(
    r?.city ||
    r?.town ||
    r?.address?.city ||
    r?.recipient?.city ||
    r?.consignee?.city ||
    r?.customer?.city ||
    ''
  );

  return {
    name: name || '—',
    street: [street, house].filter(Boolean).join(' ') || '—',
    city: [postal, city].filter(Boolean).join(' ') || '—',
    text: `${name || '—'}\n${[street, house].filter(Boolean).join(' ') || '—'}\n${[postal, city].filter(Boolean).join(' ') || '—'}`
  };
};

function renderAddressCell(value) {
  if (!value) return '—';

  if (typeof value === 'object' && value !== null) {
    const name = String(value.name || '—').trim() || '—';
    const street = String(value.street || '—').trim() || '—';
    const city = String(value.city || '—').trim() || '—';

    return `
      <div style="white-space:normal; line-height:1.2; text-align:left;">
        <div>${esc(name)}</div>
        <div>${esc(street)}</div>
        <div>${esc(city)}</div>
      </div>
    `;
  }

  const parts = String(value || '')
    .split(/\n+/)
    .map(v => v.trim())
    .filter(Boolean);

  if (!parts.length) return '—';

  return `
    <div style="white-space:normal; line-height:1.2; text-align:left;">
      ${parts.map(line => `<div>${esc(line)}</div>`).join('')}
    </div>
  `;
}

formatModalCell = function (key, value) {
  if (key === 'parcel' && value && value !== '—') {
    return `<button class="${NS}psn-btn" data-psn="${esc(String(value))}">${esc(String(value))}</button>`;
  }

  if (key === 'parcelsList') {
    const arr = Array.isArray(value) ? value.filter(Boolean) : [];
    if (!arr.length) return '—';
    return `<div class="${NS}psn-wrap">${arr.map(psn => `<button class="${NS}psn-btn" data-psn="${esc(String(psn))}">${esc(String(psn))}</button>`).join('')}</div>`;
  }

  if (key === 'status') {
    const txt = String(value || '—');
    return `<span class="${NS}status ${getStatusBadgeClass(txt)}">${esc(txt)}</span>`;
  }

  if (key === 'serviceCode') {
    return renderServiceCodeBadge(value || '—');
  }

  if (key === 'address') {
    return renderAddressCell(value);
  }

  return esc(value ?? '');
};

openHindranceCodeDetails = async function ({ partner, tour, code }) {
  setLoading(true, 'Lade Zustellhindernis-Details …');
  try {
    const rows = await fetchDetailRows(({ loadedPages, totalPages, rowsLoaded }) => {
      setLoading(true, `Lade Detaildaten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
    }, false);

    const wantedCode = cleanAdditionalCode(code);

    const baseRows = rows.filter(r =>
      r.__type === 'DELIVERY' &&
      isDeliveryHindranceStatus(r.__statusNorm || r.__status, r.__raw) &&
      r.__partner === partner &&
      String(r.__tour || '—') === String(tour || '—') &&
      cleanAdditionalCode(r.__additionalCode) === wantedCode
    );

    const listRows = baseRows.flatMap(r => {
      const psns = r.__parcelList.length ? r.__parcelList : ['—'];
      const realCode = cleanAdditionalCode(r.__additionalCode);

      return psns.map(psn => ({
        parcel: psn,
        systempartner: r.__partner,
        driver: r.__driver || '—',
        tour: r.__tour || '—',
        serviceCode: r.__serviceCode || '—',
        additionalCode: realCode || '—',
        address: r.__addr,
        status: r.__status || '—',
        time: formatTime(r.__deliveredAt) || '—'
      }));
    });

    const html = `
      <div style="overflow:auto">
        <table class="${NS}tbl" id="${NS}hindrance-detail-table" style="min-width:1500px">
          <thead>
            <tr>
              <th>Paketscheinnummer</th>
              <th>Systempartner</th>
              <th>Fahrername</th>
              <th>Tour</th>
              <th>Servicecode</th>
              <th>Grund</th>
              <th>Adresse</th>
              <th>Status</th>
              <th>Zeit</th>
            </tr>
          </thead>
          <tbody>
            ${listRows.map(r => `
              <tr>
                <td>${r.parcel !== '—' ? `<button class="${NS}psn-btn" data-psn="${esc(String(r.parcel))}">${esc(String(r.parcel))}</button>` : '—'}</td>
                <td>${esc(r.systempartner)}</td>
                <td class="${NS}driver-cell" title="${esc(r.driver || '—')}">${esc(r.driver || '—')}</td>
                <td>${esc(r.tour)}</td>
                <td>${formatModalCell('serviceCode', r.serviceCode)}</td>
                <td>${esc(r.additionalCode)}</td>
                <td>${formatModalCell('address', r.address)}</td>
                <td>${formatModalCell('status', r.status)}</td>
                <td>${esc(r.time)}</td>
              </tr>
            `).join('')}
          </tbody>
          <tfoot>
            <tr>
              <td>Gesamt</td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td>${listRows.length}</td>
            </tr>
          </tfoot>
        </table>
      </div>
    `;

    openModalHtml(`Zustellhindernisse – ${partner} – Tour ${tour} – Grund ${wantedCode}`, html);
  } finally {
    setLoading(false);
  }
};

    async function openHindranceCodeDetailsAll({ partner = '', tour = '' }) {
  setLoading(true, 'Lade Zustellhindernis-Details …');
  try {
    const rows = await fetchDetailRows(({ loadedPages, totalPages, rowsLoaded }) => {
      setLoading(true, `Lade Detaildaten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
    }, false);

    const baseRows = rows.filter(r => {
      if (r.__type !== 'DELIVERY') return false;
      if (!isDeliveryHindranceStatus(r.__statusNorm || r.__status, r.__raw)) return false;
      if (!hasUsableAdditionalCode(r)) return false;
      if (partner && r.__partner !== partner) return false;
      if (tour && String(r.__tour || '—') !== String(tour || '—')) return false;
      return true;
    });

    const listRows = baseRows.flatMap(r => {
      const psns = r.__parcelList.length ? r.__parcelList : ['—'];
      const realCode = cleanAdditionalCode(r.__additionalCode);

      return psns.map(psn => ({
        parcel: psn,
        systempartner: r.__partner,
        driver: r.__driver || '—',
        tour: r.__tour || '—',
        serviceCode: r.__serviceCode || '—',
        additionalCode: realCode || '—',
        address: r.__addr,
        status: r.__status || '—',
        time: formatTime(r.__deliveredAt) || '—'
      }));
    });

    const html = `
      <div style="overflow:auto">
        <table class="${NS}tbl" id="${NS}hindrance-detail-all-table" style="min-width:1500px">
          <thead>
            <tr>
              <th>Paketscheinnummer</th>
              <th>Systempartner</th>
              <th>Fahrername</th>
              <th>Tour</th>
              <th>Servicecode</th>
              <th>Grund</th>
              <th>Adresse</th>
              <th>Status</th>
              <th>Zeit</th>
            </tr>
          </thead>
          <tbody>
            ${listRows.map(r => `
              <tr>
                <td>${r.parcel !== '—' ? `<button class="${NS}psn-btn" data-psn="${esc(String(r.parcel))}">${esc(String(r.parcel))}</button>` : '—'}</td>
                <td>${esc(r.systempartner)}</td>
                <td class="${NS}driver-cell" title="${esc(r.driver || '—')}">${esc(r.driver || '—')}</td>
                <td>${esc(r.tour)}</td>
                <td>${formatModalCell('serviceCode', r.serviceCode)}</td>
                <td>${esc(r.additionalCode)}</td>
                <td>${formatModalCell('address', r.address)}</td>
                <td>${formatModalCell('status', r.status)}</td>
                <td>${esc(r.time)}</td>
              </tr>
            `).join('')}
          </tbody>
          <tfoot>
            <tr>
              <td>Gesamt</td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td>
                <span class="${NS}numlink"
                      data-kind="hindranceCodeDetailsAll"
                      data-partner="${esc(partner)}"
                      data-tour="${esc(tour)}">${listRows.length}</span>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    `;

    const titleParts = [];
    if (partner) titleParts.push(partner);
    if (tour) titleParts.push(`Tour ${tour}`);
    titleParts.push('Zustellhindernisse – alle');

    openModalHtml(titleParts.join(' – '), html);
  } finally {
    setLoading(false);
  }
}

openHindranceCodeSummary = async function ({ level, partner, tour }) {
  setLoading(true, 'Lade Zustellhindernisse …');
  try {
    const rows = await fetchDetailRows(({ loadedPages, totalPages, rowsLoaded }) => {
      setLoading(true, `Lade Detaildaten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
    }, false);

    let baseRows = rows.filter(r =>
      r.__type === 'DELIVERY' &&
      isDeliveryHindranceStatus(r.__statusNorm || r.__status, r.__raw) &&
      hasUsableAdditionalCode(r)
    );

    if (level === 'partner') {
      baseRows = baseRows.filter(r => r.__partner === partner);
    } else if (level === 'tour') {
      baseRows = baseRows.filter(r => r.__partner === partner && String(r.__tour) === String(tour));
    }

    const groups = new Map();

    for (const r of baseRows) {
      const code = cleanAdditionalCode(r.__additionalCode);
      const key = `${r.__partner}|||${r.__driver || ''}|||${r.__tour}|||${code}`;

      if (!groups.has(key)) {
        groups.set(key, {
          systempartner: r.__partner,
          driver: r.__driver || '—',
          tour: r.__tour || '—',
          additionalCode: code,
          count: 0
        });
      }

      groups.get(key).count += hindranceItemCount(r);
    }

    const listRows = Array.from(groups.values()).sort((a, b) => {
      const c1 = collator.compare(a.systempartner, b.systempartner);
      if (c1 !== 0) return c1;
      const c2 = collator.compare(String(a.tour || ''), String(b.tour || ''));
      if (c2 !== 0) return c2;
      return collator.compare(a.additionalCode, b.additionalCode);
    });

    const titleParts = [];
    if (partner) titleParts.push(partner);
    if (tour) titleParts.push(`Tour ${tour}`);
    titleParts.push('Zustellhindernisse');

    const html = `
      <div style="overflow:auto">
        <table class="${NS}tbl" id="${NS}hindrance-code-table" style="min-width:1100px">
          <thead>
            <tr>
              <th>Systempartner</th>
              <th>Fahrername</th>
              <th>Tour</th>
              <th>Grund</th>
              <th>Anzahl</th>
            </tr>
          </thead>
          <tbody>
            ${listRows.map(r => `
              <tr>
                <td>${esc(r.systempartner)}</td>
                <td class="${NS}driver-cell" title="${esc(r.driver || '—')}">${esc(r.driver || '—')}</td>
                <td>${esc(r.tour)}</td>
                <td>${esc(r.additionalCode)}</td>
                <td>
                  <span class="${NS}numlink"
                        data-kind="hindranceCodeDetails"
                        data-partner="${esc(r.systempartner)}"
                        data-tour="${esc(r.tour)}"
                        data-code="${esc(r.additionalCode)}">${r.count}</span>
                </td>
              </tr>
            `).join('')}
          </tbody>
          <tfoot>
            <tr>
              <td>Gesamt</td>
              <td></td>
              <td></td>
              <td></td>
              <td>
                <span class="${NS}numlink"
                      data-kind="hindranceCodeDetailsAll"
                      data-partner="${esc(partner || '')}"
                      data-tour="${esc(tour || '')}">${listRows.reduce((s, r) => s + r.count, 0)}</span>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    `;

    openModalHtml(titleParts.join(' – '), html);
  } finally {
    setLoading(false);
  }
};

openHindranceCodeDetails = async function ({ partner, tour, code }) {
  setLoading(true, 'Lade Zustellhindernis-Details …');
  try {
    const rows = await fetchDetailRows(({ loadedPages, totalPages, rowsLoaded }) => {
      setLoading(true, `Lade Detaildaten … Seite ${loadedPages} von ${totalPages} · Datensätze ${rowsLoaded}`);
    }, false);

    const wantedCode = cleanAdditionalCode(code);

    const baseRows = rows.filter(r =>
      r.__type === 'DELIVERY' &&
      isDeliveryHindranceStatus(r.__statusNorm || r.__status, r.__raw) &&
      r.__partner === partner &&
      String(r.__tour || '—') === String(tour || '—') &&
      cleanAdditionalCode(r.__additionalCode) === wantedCode
    );

    const listRows = baseRows.flatMap(r => {
      const psns = r.__parcelList.length ? r.__parcelList : ['—'];
      const realCode = cleanAdditionalCode(r.__additionalCode);

      return psns.map(psn => ({
        parcel: psn,
        systempartner: r.__partner,
        driver: r.__driver || '—',
        tour: r.__tour || '—',
        serviceCode: r.__serviceCode || '—',
        additionalCode: realCode || '—',
        address: r.__addr,
        status: r.__status || '—',
        time: formatTime(r.__deliveredAt) || '—'
      }));
    });

    const html = `
      <div style="overflow:auto">
        <table class="${NS}tbl" id="${NS}hindrance-detail-table" style="min-width:1500px">
          <thead>
            <tr>
              <th>Paketscheinnummer</th>
              <th>Systempartner</th>
              <th>Fahrername</th>
              <th>Tour</th>
              <th>Servicecode</th>
              <th>Grund</th>
              <th>Adresse</th>
              <th>Status</th>
              <th>Zeit</th>
            </tr>
          </thead>
          <tbody>
            ${listRows.map(r => `
              <tr>
                <td>${r.parcel !== '—' ? `<button class="${NS}psn-btn" data-psn="${esc(String(r.parcel))}">${esc(String(r.parcel))}</button>` : '—'}</td>
                <td>${esc(r.systempartner)}</td>
                <td class="${NS}driver-cell" title="${esc(r.driver || '—')}">${esc(r.driver || '—')}</td>
                <td>${esc(r.tour)}</td>
                <td>${formatModalCell('serviceCode', r.serviceCode)}</td>
                <td>${esc(r.additionalCode)}</td>
                <td>${formatModalCell('address', r.address)}</td>
                <td>${formatModalCell('status', r.status)}</td>
                <td>${esc(r.time)}</td>
              </tr>
            `).join('')}
          </tbody>
          <tfoot>
            <tr>
              <td>Gesamt</td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td></td>
              <td>
                <span class="${NS}numlink"
                      data-kind="hindranceCodeDetails"
                      data-partner="${esc(partner)}"
                      data-tour="${esc(tour)}"
                      data-code="${esc(wantedCode)}">${listRows.length}</span>
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    `;

    openModalHtml(`Zustellhindernisse – ${partner} – Tour ${tour} – Grund ${wantedCode}`, html);
  } finally {
    setLoading(false);
  }
};

(function installFooterTotalClicks() {
  if (window.__spx_footer_total_clicks_installed) return;
  window.__spx_footer_total_clicks_installed = true;

  document.addEventListener('click', (e) => {
    const el = e.target.closest('[data-kind="hindranceCodeDetailsAll"], [data-kind="hindranceCodeDetails"]');
    if (!el) return;

    if (el.dataset.kind === 'hindranceCodeDetailsAll') {
      e.preventDefault();
      e.stopPropagation();
      openHindranceCodeDetailsAll({
        partner: String(el.dataset.partner || ''),
        tour: String(el.dataset.tour || '')
      }).catch(console.error);
      return;
    }

    if (el.dataset.kind === 'hindranceCodeDetails') {
      e.preventDefault();
      e.stopPropagation();
      openHindranceCodeDetails({
        partner: String(el.dataset.partner || ''),
        tour: String(el.dataset.tour || ''),
        code: String(el.dataset.code || '')
      }).catch(console.error);
    }
  }, true);
})();

})();
