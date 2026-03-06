const express = require('express');
const router = express.Router();
const axios = require('axios');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const { Workbook } = require('exceljs');
const os = require('os');
const mongoose = require('mongoose');
const jwt = require('jsonwebtoken');

const JWT_SECRET = process.env.JWT_MITRA_SECRET || 'pms-mitra-secret-key-2025';

const BLITZ_LOGIN_URL = 'https://driver-api.rideblitz.id/panel/login';
const BLITZ_ORDERS_SEARCH_URL = 'https://adminapis.rideblitz.id/api/v1/orders';
const BLITZ_VALIDATE_BATCH_URL = 'https://bmc.rideblitz.id/v2/validate/batch/orders';
const BLITZ_ADD_BATCH_URL = 'https://bmc.rideblitz.id/v2/add/batch/orders';
const BLITZ_BATCH_DETAILS_URL = 'https://bmc.rideblitz.id/v1/batches/details';
const BLITZ_GENERATE_BATCH_URL = 'https://bmc.rideblitz.id/v1/generate/batch';
const BLITZ_NEARBY_DRIVERS_URL = 'https://driver-api.rideblitz.id/panel/driver';
const BLITZ_ASSIGN_DRIVER_URL = 'https://amc.rideblitz.id/v1/batch/assign/driver';
const BLITZ_DRIVER_LIST_URL = 'https://driver-api.rideblitz.id/v2/panel/driver-list';
const BLITZ_DRIVER_PERFORMANCE_URL = 'https://driver-api.rideblitz.id/v1/panel/driver/performance/batch';
const BLITZ_REMOVE_VALIDATE_URL = 'https://bmc.rideblitz.id/v2/validate/remove/batch/order';
const BLITZ_REMOVE_ORDER_URL = 'https://bmc.rideblitz.id/v2/remove/batch/orders';
const BLITZ_DRIVER_PROFILE_URL = 'https://driver-api.rideblitz.id/panel/driver-profile';
const BLITZ_SAVE_BATCH_URL = 'https://bmc.rideblitz.id/v1/save/batch/orders';

const AUTOMATION_SCRIPT_PATH = path.join(__dirname, '..', 'utils', 'automation.py');
const BLITZ_STATUSES_SKIP_UPLOAD = ['created', 'unbatched', 'batched', 'assigned', 'picked_up', 'in_transit', 'delivered'];

const PARALLEL_BATCH_CONCURRENCY = 20;
const SEARCH_PER_BATCH_TIMEOUT_MS = 15000;
const SEARCH_OVERALL_TIMEOUT_MS = 30000;
const TOKEN_TTL_MS = 55 * 60 * 1000;
const BULK_CACHE_TTL_MS = 5 * 60 * 1000;
const NEGATIVE_CACHE_TTL_MS = 60 * 1000;
const AUTOMATION_TIMEOUT_MS = 240000;

const tokenCache = {};
const bulkOrderCache = new Map();
const inFlightBatchKeys = new Map();
const inFlightSearchRequests = new Map();

let activeUpload = null;
let activeProc = null;
let pendingUploads = [];

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const writeInterruptFile = (file) => {
  try { fs.writeFileSync(file, 'interrupt'); } catch {}
};

const cleanupFile = (file) => {
  if (file && fs.existsSync(file)) {
    try { fs.unlinkSync(file); } catch {}
  }
};

const getJakartaDateString = (date) => {
  const jakartaOffset = 7 * 60 * 60 * 1000;
  const jd = new Date(date.getTime() + jakartaOffset);
  return `${jd.toISOString().split('T')[0]} ${jd.toISOString().split('T')[1].split('.')[0]}`;
};

const getBlitzDateRange = (daysBack = 90) => {
  const now = new Date();
  const past = new Date(now.getTime() - daysBack * 24 * 60 * 60 * 1000);
  return { startDate: getJakartaDateString(past), endDate: getJakartaDateString(now) };
};

const PREFIX_MAP = [
  { from: 'INV-', to: 'V-' },
  { from: 'V-', to: 'INV-' },
];

const getAlternativeOrderId = (id) => {
  for (const { from, to } of PREFIX_MAP) {
    if (id.startsWith(from)) return to + id.slice(from.length);
  }
  return null;
};

const getCached = (id) => {
  const entry = bulkOrderCache.get(id);
  if (!entry) return undefined;
  const ttl = entry.data?.exists ? BULK_CACHE_TTL_MS : NEGATIVE_CACHE_TTL_MS;
  if (Date.now() - entry.ts > ttl) {
    bulkOrderCache.delete(id);
    return undefined;
  }
  return entry.data;
};

const setCached = (id, data) => {
  bulkOrderCache.set(id, { data, ts: Date.now() });
};

const getValidationDataByMerchantOrderId = async (project, merchantOrderId) => {
  const order = await mongoose.connection.db
    .collection(`${project}_merchant_orders`)
    .findOne({ merchant_order_id: merchantOrderId });
  if (!order?.sender_name) return null;
  return mongoose.connection.db
    .collection('adminpanel_validations')
    .findOne({ sender_name: order.sender_name });
};

const getCredentialsByDriverId = async (driverId, project) => {
  if (!driverId || !project) return null;
  try {
    const driverData = await mongoose.connection.db
      .collection(`${project}_delivery`)
      .findOne({ driver_id: driverId.toString() });
    if (!driverData?.user_id) return null;
    const credential = await mongoose.connection.db
      .collection('blitz_logins')
      .findOne({ user_id: driverData.user_id, status: 'active' });
    return credential ? { username: credential.username, password: credential.password } : null;
  } catch {
    return null;
  }
};

const getBlitzCredentials = async (req) => {
  if (req?.body?._blitz_un && req?.body?._blitz_pw)
    return { username: req.body._blitz_un, password: req.body._blitz_pw };
  if (req?.query?._blitz_un && req?.query?._blitz_pw)
    return { username: req.query._blitz_un, password: req.query._blitz_pw };
  if (req?.headers?.authorization?.startsWith('Bearer ')) {
    try {
      const decoded = jwt.verify(req.headers.authorization.substring(7), JWT_SECRET);
      if (decoded.blitz_username && decoded.blitz_password)
        return { username: decoded.blitz_username, password: decoded.blitz_password };
      if (decoded.user_id) {
        const credential = await mongoose.connection.db
          .collection('blitz_logins')
          .findOne({ user_id: decoded.user_id, status: 'active' });
        if (credential) return { username: credential.username, password: credential.password };
      }
    } catch {}
  }
  if (req?.body?.driverId && (req?.body?.project || req?.query?.project)) {
    const cred = await getCredentialsByDriverId(
      req.body.driverId,
      req.body.project || req.query.project
    );
    if (cred) return cred;
  }
  const credential = await mongoose.connection.db
    .collection('blitz_logins')
    .findOne({ status: 'active' });
  if (!credential) throw new Error('No active Blitz credentials found in database');
  return { username: credential.username, password: credential.password };
};

const loginToBlitz = async (username, password) => {
  const response = await axios.post(
    BLITZ_LOGIN_URL,
    { username, password },
    { headers: { 'Content-Type': 'application/json', Accept: 'application/json' }, timeout: 30000 }
  );
  if (response.data.result) return response.data.data.access_token;
  throw new Error('Login failed: ' + (response.data.message || 'Unknown error'));
};

const getAccessToken = async (req) => {
  const credentials = await getBlitzCredentials(req);
  const cacheKey = credentials.username;
  if (tokenCache[cacheKey]?.expiry && Date.now() < tokenCache[cacheKey].expiry)
    return tokenCache[cacheKey].token;
  const token = await loginToBlitz(credentials.username, credentials.password);
  tokenCache[cacheKey] = { token, expiry: Date.now() + TOKEN_TTL_MS };
  return token;
};

const fetchBatchDetails = async (batchId, accessToken) => {
  try {
    const response = await axios.get(`${BLITZ_BATCH_DETAILS_URL}/${batchId}`, {
      headers: { Accept: 'application/json', Authorization: accessToken, bt: '2' },
      timeout: 15000,
    });
    if (response.data.result && response.data.data) {
      const d = response.data.data;
      return {
        driver_name: d.driver?.name || null,
        driver_contact: d.driver?.mobile ? `+62${d.driver.mobile}` : null,
        batch_status: d.batch_status?.name || null,
      };
    }
    return null;
  } catch {
    return null;
  }
};

const getSearchGroupKey = (id) => {
  if (!id) return id;
  const dashIdx = id.indexOf('-');
  if (dashIdx === -1) return id;
  const secondDash = id.indexOf('-', dashIdx + 1);
  if (secondDash === -1) return id;
  return id.slice(0, secondDash);
};

const fetchAllPagesForQuery = async (q, accessToken, startDate, endDate) => {
  const allResults = [];
  let page = 1;
  const limit = 100;

  while (true) {
    try {
      const response = await axios.get(BLITZ_ORDERS_SEARCH_URL, {
        params: {
          sort: 'created_at',
          dir: '-1',
          page,
          limit,
          instant_type: 'non-instant',
          start_date: startDate,
          end_date: endDate,
          q,
          pickup_schedule_type: 'standard,scheduled,immediate',
          pickup_sla_model: 'pickup_slots,operational_hours',
        },
        headers: { Accept: 'application/json', Authorization: accessToken },
        timeout: SEARCH_PER_BATCH_TIMEOUT_MS,
      });

      const results = Array.isArray(response.data?.results)
        ? response.data.results
        : Array.isArray(response.data?.data)
          ? response.data.data
          : [];

      allResults.push(...results);

      const total = response.data?.total || response.data?.meta?.total || 0;
      if (results.length < limit || allResults.length >= total) break;
      page++;
    } catch {
      break;
    }
  }

  return allResults;
};

const searchGroupOfIds = async (groupKey, ids, accessToken) => {
  const { startDate, endDate } = getBlitzDateRange(90);
  const result = {};

  const exactIdSet = new Set(
    ids.flatMap((id) => {
      const alt = getAlternativeOrderId(id);
      return alt ? [id, alt] : [id];
    })
  );

  const queriesToRun = [groupKey];
  const altGroupKey = (() => {
    const alt = getAlternativeOrderId(ids[0]);
    if (!alt) return null;
    return getSearchGroupKey(alt);
  })();
  if (altGroupKey && altGroupKey !== groupKey) queriesToRun.push(altGroupKey);

  const allFound = new Map();

  await Promise.allSettled(
    queriesToRun.map(async (q) => {
      const items = await fetchAllPagesForQuery(q, accessToken, startDate, endDate);
      for (const item of items) {
        const mid = (item.merchant_order_id || '').trim();
        if (exactIdSet.has(mid) && !allFound.has(mid)) {
          allFound.set(mid, item);
        }
      }
    })
  );

  const batchIdsToFetch = [...new Set(
    [...allFound.values()].map((item) => item.batch_id).filter(Boolean)
  )];

  const batchDetailsMap = new Map();
  await Promise.allSettled(
    batchIdsToFetch.map(async (batchId) => {
      const details = await fetchBatchDetails(batchId, accessToken);
      if (details) batchDetailsMap.set(batchId, details);
    })
  );

  for (const id of ids) {
    const alt = getAlternativeOrderId(id);
    const matchedItem = allFound.get(id) || (alt ? allFound.get(alt) : null);

    if (matchedItem) {
      const details = matchedItem.batch_id ? batchDetailsMap.get(matchedItem.batch_id) : null;
      const entry = {
        exists: true,
        order_status: matchedItem.order_status,
        batch_id: matchedItem.batch_id,
        blitz_merchant_order_id: matchedItem.merchant_order_id,
        driver_name: details?.driver_name || matchedItem.driver_name || null,
        driver_contact: details?.driver_contact || matchedItem.driver_contact || null,
        batch_status: details?.batch_status || null,
      };
      result[id] = entry;
      setCached(id, entry);
      if (alt) setCached(alt, entry);
    } else {
      result[id] = { exists: false };
      setCached(id, { exists: false });
    }
  }

  return result;
};

const runParallelBatchSearch = async (ids, accessToken) => {
  const groupMap = new Map();
  for (const id of ids) {
    const key = getSearchGroupKey(id);
    if (!groupMap.has(key)) groupMap.set(key, []);
    groupMap.get(key).push(id);
  }

  const result = {};

  const processGroup = async (groupKey, groupIds) => {
    if (inFlightBatchKeys.has(groupKey)) {
      const res = await inFlightBatchKeys.get(groupKey);
      Object.assign(result, res);
      return;
    }

    const promise = searchGroupOfIds(groupKey, groupIds, accessToken).catch(() => ({}));
    inFlightBatchKeys.set(groupKey, promise);

    let res = {};
    try {
      res = await promise;
    } finally {
      inFlightBatchKeys.delete(groupKey);
    }

    Object.assign(result, res);
  };

  const entries = [...groupMap.entries()];
  const queue = entries.map(([key, groupIds]) => () => processGroup(key, groupIds));
  const workers = [];

  const runWorker = async () => {
    while (queue.length > 0) {
      const task = queue.shift();
      if (task) await task();
    }
  };

  const workerCount = Math.min(PARALLEL_BATCH_CONCURRENCY, entries.length);
  for (let i = 0; i < workerCount; i++) {
    workers.push(runWorker());
  }

  await Promise.allSettled(workers);
  return result;
};

const extractResultsArray = (data) => {
  if (!data) return [];
  if (Array.isArray(data.results)) return data.results;
  if (Array.isArray(data.data)) return data.data;
  if (Array.isArray(data.orders)) return data.orders;
  if (data.data && Array.isArray(data.data.results)) return data.data.results;
  if (data.data && Array.isArray(data.data.orders)) return data.data.orders;
  return [];
};

const extractBlitzErrorMessage = (errorData) =>
  errorData?.error?.message || errorData?.message || null;

const extractOrdersArray = (responseData) => {
  if (!responseData) return [];
  if (Array.isArray(responseData.data)) return responseData.data;
  if (Array.isArray(responseData.blitz_error?.data)) return responseData.blitz_error.data;
  if (Array.isArray(responseData.blitz_error?.blitz_error?.data))
    return responseData.blitz_error.blitz_error.data;
  return [];
};

const extractValidationErrors = (responseData) => {
  return extractOrdersArray(responseData)
    .filter((o) => o.validation?.is_valid === false)
    .map((o) => ({
      merchant_order_id: o.merchant_order_id,
      awb_number: o.awb_number || '',
      validation_message:
        o.validation.reason?.trim() && o.validation.reason !== 'Invalid order'
          ? o.validation.reason
          : o.validation.message || 'Invalid order',
      delivery_attempt_number: o.delivery_attempt || 0,
      package_weight: o.weight ? String(o.weight) + ' kg' : '',
      dropoff_address: o.dropoff?.address || '',
      dropoff_city: o.dropoff?.city || '',
      dropoff_district: o.dropoff?.district || '',
      dropoff_postal_code: o.dropoff?.postal_code || '',
    }));
};

const hasInvalidOrders = (data) =>
  extractOrdersArray(data).some((o) => o.validation?.is_valid === false);

const createExcelFromOrders = async (orders) => {
  const workbook = new Workbook();
  const ws = workbook.addWorksheet('Sheet1');
  const headers = [
    'merchant_order_id*', 'weight*', 'width', 'height', 'length',
    'payment_type*', 'cod_amount', 'sender_name*', 'sender_phone*',
    'pickup_instructions', 'consignee_name*', 'consignee_phone*',
    'destination_district', 'destination_city*', 'destination_province',
    'destination_postalcode*', 'destination_address*', 'dropoff_lat',
    'dropoff_long', 'dropoff_instructions', 'item_value*', 'product_details*',
  ];
  ws.addRow(headers);
  ws.getRow(1).eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10, name: 'Calibri' };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  for (const o of orders) {
    ws.addRow([
      o.merchant_order_id || '', o.weight || 0, o.width || 0, o.height || 0, o.length || 0,
      o.payment_type || 'non_cod', o.cod_amount || 0, o.sender_name || '', o.sender_phone || '',
      o.pickup_instructions || '', o.consignee_name || '', o.consignee_phone || '',
      o.destination_district || '', o.destination_city || '', o.destination_province || '',
      o.destination_postalcode || '', o.destination_address || '', o.dropoff_lat || 0,
      o.dropoff_long || 0, o.dropoff_instructions || '', o.item_value || 0, o.product_details || '',
    ]);
  }
  const tempFile = path.join(
    os.tmpdir(),
    `blitz_upload_${Date.now()}_${Math.random().toString(36).slice(2, 8)}.xlsx`
  );
  await workbook.xlsx.writeFile(tempFile);
  return tempFile;
};

const dedupeOrders = (orders) => {
  const seen = new Set();
  return orders.filter((o) => {
    if (seen.has(o.merchant_order_id)) return false;
    seen.add(o.merchant_order_id);
    return true;
  });
};

const drainUploadQueue = () => {
  if (pendingUploads.length === 0 || activeUpload) return;

  const next = pendingUploads.shift();
  const toMerge = [next];
  const remaining = [];

  for (const p of pendingUploads) {
    const sameConfig =
      p.hubId === next.hubId &&
      p.business === next.business &&
      p.city === next.city &&
      p.serviceType === next.serviceType;
    if (sameConfig) toMerge.push(p);
    else remaining.push(p);
  }
  pendingUploads = remaining;

  if (toMerge.length === 1) {
    startUpload(next);
    return;
  }

  const mergedOrders = dedupeOrders(toMerge.flatMap((u) => u.orders));
  const allResolvers = toMerge.map((u) => ({ resolve: u.resolve, reject: u.reject }));

  startUpload({
    orders: mergedOrders,
    hubId: next.hubId,
    business: next.business,
    city: next.city,
    serviceType: next.serviceType,
    username: next.username,
    password: next.password,
    resolve: (r) => allResolvers.forEach((ar) => ar.resolve(r)),
    reject: (e) => allResolvers.forEach((ar) => ar.reject(e)),
  });
};

const startUpload = async (upload) => {
  const { orders, hubId, business, city, serviceType, username, password, resolve, reject } = upload;

  let excelFile;
  try {
    excelFile = await createExcelFromOrders(orders);
  } catch (e) {
    activeUpload = null;
    activeProc = null;
    reject(e);
    drainUploadQueue();
    return;
  }

  const tempInstanceId = `pending_${Date.now()}`;

  activeUpload = {
    instanceId: tempInstanceId,
    interruptFile: `/tmp/blitz_interrupt_${tempInstanceId}`,
    orders,
    hubId,
    business,
    city,
    serviceType,
    username,
    password,
    passedCheckpoint: false,
    resolve,
    reject,
  };

  if (!fs.existsSync(AUTOMATION_SCRIPT_PATH)) {
    activeUpload = null;
    activeProc = null;
    reject(new Error(`automation.py not found at: ${AUTOMATION_SCRIPT_PATH}`));
    drainUploadQueue();
    return;
  }

  let stdoutBuffer = '';
  let stderrBuffer = '';
  let timeoutHandle;

  const proc = spawn('python3', [AUTOMATION_SCRIPT_PATH], {
    env: {
      ...process.env,
      BLITZ_USERNAME: username,
      BLITZ_PASSWORD: password,
      BLITZ_FILE_PATH: excelFile,
      BLITZ_BUSINESS_HUB: hubId.toString(),
      BLITZ_BUSINESS: business.toString(),
      BLITZ_CITY: city.toString(),
      BLITZ_SERVICE_TYPE: serviceType.toString(),
      BLITZ_AUTO_SUBMIT: 'true',
      BLITZ_GOOGLE_SHEET_URL: '',
      BLITZ_KEEP_FILE: 'false',
    },
  });

  activeProc = proc;

  timeoutHandle = setTimeout(() => {
    try { proc.kill('SIGTERM'); } catch {}
    activeProc = null;
    activeUpload = null;
    reject(new Error(`Automation timeout after ${AUTOMATION_TIMEOUT_MS / 1000}s`));
    drainUploadQueue();
  }, AUTOMATION_TIMEOUT_MS);

  proc.stdout.on('data', (chunk) => {
    const text = chunk.toString();
    stdoutBuffer += text;
    text.split('\n').forEach((line) => {
      line = line.trim();
      if (!line) return;
      console.log(line);

      if (!activeUpload) return;

      const instanceMatch = line.match(/\[(?:DEBUG|CHECKPOINT|INTERRUPTED)\]\[([^\]]+)\]\[([a-f0-9]{8})\]/);
      if (instanceMatch) {
        const realId = instanceMatch[2];
        if (activeUpload.instanceId !== realId) {
          activeUpload.instanceId = realId;
          activeUpload.interruptFile = `/tmp/blitz_interrupt_${realId}`;
        }
      }

      if (line.includes('[CHECKPOINT][SAVE_CLICKED]')) {
        activeUpload.passedCheckpoint = true;
      }
    });
  });

  proc.stderr.on('data', (chunk) => {
    stderrBuffer += chunk.toString();
    console.error(chunk.toString().trim());
  });

  proc.on('close', (code) => {
    if (activeProc !== proc) {
      cleanupFile(excelFile);
      return;
    }

    clearTimeout(timeoutHandle);
    cleanupFile(excelFile);
    activeProc = null;

    const finishedUpload = activeUpload;
    activeUpload = null;

    if (!finishedUpload) {
      drainUploadQueue();
      return;
    }

    if (code === 0) {
      finishedUpload.resolve({ success: true });
    } else if (code === 2) {
      console.log('[upload-queue] Upload interrupted cleanly (exit 2)');
    } else {
      finishedUpload.reject(
        new Error(`Automation failed (exit ${code}): ${stderrBuffer || 'Unknown error'}`)
      );
    }

    drainUploadQueue();
  });

  proc.on('error', (e) => {
    if (activeProc !== proc) return;
    clearTimeout(timeoutHandle);
    cleanupFile(excelFile);
    activeProc = null;
    const finishedUpload = activeUpload;
    activeUpload = null;
    if (finishedUpload) finishedUpload.reject(new Error(`Failed to start automation: ${e.message}`));
    drainUploadQueue();
  });
};

const smartUpload = (orders, hubId, business, city, serviceType, username, password) => {
  return new Promise((resolve, reject) => {
    if (!activeUpload) {
      startUpload({ orders, hubId, business, city, serviceType, username, password, resolve, reject });
      return;
    }

    if (!activeUpload.passedCheckpoint) {
      const mergedOrders = dedupeOrders([...activeUpload.orders, ...orders]);
      const mergedHubId = activeUpload.hubId;
      const mergedBusiness = activeUpload.business;
      const mergedCity = activeUpload.city;
      const mergedServiceType = activeUpload.serviceType;

      const allResolvers = [
        { resolve: activeUpload.resolve, reject: activeUpload.reject },
        { resolve, reject },
        ...pendingUploads.map((p) => ({ resolve: p.resolve, reject: p.reject })),
      ];
      pendingUploads = [];

      writeInterruptFile(activeUpload.interruptFile);

      if (activeProc) {
        try { activeProc.kill('SIGTERM'); } catch {}
        activeProc = null;
      }

      activeUpload = null;

      startUpload({
        orders: mergedOrders,
        hubId: mergedHubId,
        business: mergedBusiness,
        city: mergedCity,
        serviceType: mergedServiceType,
        username,
        password,
        resolve: (r) => allResolvers.forEach((ar) => ar.resolve(r)),
        reject: (e) => allResolvers.forEach((ar) => ar.reject(e)),
      });
    } else {
      pendingUploads.push({ orders, hubId, business, city, serviceType, username, password, resolve, reject });
    }
  });
};

const classifyOrdersByBlitzStatus = async (orders, accessToken) => {
  const ids = orders.map((o) => o.merchant_order_id);
  const searchResults = await runParallelBatchSearch(ids, accessToken);

  const needUpload = [];
  const skipUpload = [];
  orders.forEach((o) => {
    const r = searchResults[o.merchant_order_id];
    if (r?.exists && BLITZ_STATUSES_SKIP_UPLOAD.includes(r.order_status?.toLowerCase()))
      skipUpload.push(o);
    else needUpload.push(o);
  });
  return { needUpload, skipUpload };
};

const waitUntilOrdersAppearInBlitz = async (ids, accessToken, maxRetries = 4, intervalMs = 5000) => {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    ids.forEach((id) => bulkOrderCache.delete(id));
    const searchResults = await runParallelBatchSearch(ids, accessToken);
    const missing = ids.filter((id) => !searchResults[id]?.exists);
    if (missing.length === 0) return { success: true, missing: [] };
    if (attempt < maxRetries) await sleep(intervalMs);
  }
  return { success: false, missing: ids };
};

const executeBatchFlow = async (
  accessToken, merchantOrderIds, hubId, driverId, coordinates, batchOnly, retryCount = 0
) => {
  const MAX_VALIDATE_RETRIES = 2;
  const VALIDATE_RETRY_DELAY_MS = 4000;

  let validateResponse;
  try {
    validateResponse = await axios.post(
      BLITZ_VALIDATE_BATCH_URL,
      { batchId: 0, hub_id: hubId, sequence_type: 1, merchant_order_ids: merchantOrderIds },
      { headers: { Accept: 'application/json', Authorization: accessToken, 'Content-Type': 'application/json', bt: '2' }, timeout: 30000 }
    );
  } catch (e) {
    const d = e.response?.data;
    if (retryCount < MAX_VALIDATE_RETRIES && (!e.response || e.response.status >= 500)) {
      await sleep(VALIDATE_RETRY_DELAY_MS);
      return executeBatchFlow(accessToken, merchantOrderIds, hubId, driverId, coordinates, batchOnly, retryCount + 1);
    }
    throw {
      statusCode: e.response?.status || 500,
      body: { success: false, message: extractBlitzErrorMessage(d) || e.message, blitz_error: d, validation_errors: extractValidationErrors(d) },
    };
  }

  if (!validateResponse.data.result) {
    if (retryCount < MAX_VALIDATE_RETRIES) {
      await sleep(VALIDATE_RETRY_DELAY_MS);
      return executeBatchFlow(accessToken, merchantOrderIds, hubId, driverId, coordinates, batchOnly, retryCount + 1);
    }
    const dataArr = validateResponse.data.data;
    const ve = Array.isArray(dataArr)
      ? dataArr.map((o) => `${o.merchant_order_id}: ${o.validation?.message || 'failed'}`).join(', ')
      : '';
    throw {
      statusCode: 400,
      body: {
        success: false,
        message: (extractBlitzErrorMessage(validateResponse.data) || 'Validation failed') + (ve ? `: ${ve}` : ''),
        blitz_error: validateResponse.data,
        validation_errors: extractValidationErrors(validateResponse.data),
      },
    };
  }

  if (hasInvalidOrders(validateResponse.data))
    throw {
      statusCode: 400,
      body: { success: false, message: 'Please remove invalid orders before creating the batch.', blitz_error: validateResponse.data, validation_errors: extractValidationErrors(validateResponse.data) },
    };

  let saveResponse;
  try {
    saveResponse = await axios.post(
      BLITZ_SAVE_BATCH_URL,
      { batchId: 0, hub_id: hubId, sequence_type: 1, merchant_order_ids: merchantOrderIds },
      { headers: { Accept: 'application/json', Authorization: accessToken, 'Content-Type': 'application/json', bt: '2' }, timeout: 30000 }
    );
  } catch (e) {
    const d = e.response?.data;
    throw {
      statusCode: e.response?.status || 500,
      body: { success: false, message: extractBlitzErrorMessage(d) || e.message, blitz_error: d, validation_errors: extractValidationErrors(d) },
    };
  }

  if (!saveResponse.data.result)
    throw {
      statusCode: 400,
      body: { success: false, message: extractBlitzErrorMessage(saveResponse.data) || 'Save batch failed', blitz_error: saveResponse.data, validation_errors: extractValidationErrors(saveResponse.data) },
    };

  const batchId = saveResponse.data.data.batch_id;

  try {
    await axios.get(`${BLITZ_GENERATE_BATCH_URL}/${batchId}`, {
      headers: { Accept: 'application/json', Authorization: accessToken, bt: '2' },
      timeout: 30000,
    });
  } catch (e) {
    if (e.response?.status !== 424)
      console.warn(`[batch-flow] generate warning batchId=${batchId}: ${e.message}`);
  }

  if (batchOnly) return { batchId, assigned: false };

  await sleep(1500);

  let assignResponse;
  try {
    assignResponse = await axios.post(
      BLITZ_ASSIGN_DRIVER_URL,
      {
        batch_id: parseInt(batchId),
        driver_id: parseInt(driverId),
        lat: parseFloat(coordinates[0]),
        lng: parseFloat(coordinates[1]),
        radius: '20km',
        allow_route_change: false,
        decline_batch_before_accept: false,
        accept_timer: 0,
        cancel_at_first_pickup: false,
        cancel_timer: 0,
      },
      { headers: { Accept: 'application/json', Authorization: accessToken, 'Content-Type': 'application/json', bt: '2' }, timeout: 30000 }
    );
  } catch (e) {
    const d = e.response?.data;
    throw {
      statusCode: e.response?.status || 500,
      body: { success: false, message: extractBlitzErrorMessage(d) || e.message, blitz_error: d },
    };
  }

  if (!assignResponse.data.result)
    throw {
      statusCode: 400,
      body: { success: false, message: extractBlitzErrorMessage(assignResponse.data) || 'Driver assignment failed', blitz_error: assignResponse.data },
    };

  return { batchId, assigned: true, assignmentId: assignResponse.data.data?.assignment_id };
};

router.get('/token', async (req, res) => {
  try {
    const credentials = await getBlitzCredentials(req);
    delete tokenCache[credentials.username];
    const token = await loginToBlitz(credentials.username, credentials.password);
    tokenCache[credentials.username] = { token, expiry: Date.now() + TOKEN_TTL_MS };
    res.json({ success: true, token });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

router.post('/search-orders', async (req, res) => {
  try {
    const { merchantOrderIds } = req.body;
    if (!merchantOrderIds || !Array.isArray(merchantOrderIds) || merchantOrderIds.length === 0)
      return res.status(400).json({ success: false, message: 'merchantOrderIds array is required' });

    const result = {};
    const uncachedIds = [];

    for (const id of merchantOrderIds) {
      const cached = getCached(id);
      if (cached !== undefined) {
        result[id] = cached;
      } else {
        uncachedIds.push(id);
      }
    }

    if (uncachedIds.length > 0) {
      const requestKey = uncachedIds.slice().sort().join('|');

      let searchPromise;
      if (inFlightSearchRequests.has(requestKey)) {
        searchPromise = inFlightSearchRequests.get(requestKey);
      } else {
        const accessToken = await getAccessToken(req);
        const promise = runParallelBatchSearch(uncachedIds, accessToken).finally(() => {
          inFlightSearchRequests.delete(requestKey);
        });
        inFlightSearchRequests.set(requestKey, promise);
        searchPromise = promise;
      }

      const timeoutPromise = new Promise((resolve) =>
        setTimeout(() => resolve(null), SEARCH_OVERALL_TIMEOUT_MS)
      );

      const searchResults = await Promise.race([searchPromise, timeoutPromise]);

      if (searchResults) {
        Object.assign(result, searchResults);
      } else {
        for (const id of uncachedIds) {
          if (!(id in result)) {
            const cached = getCached(id);
            result[id] = cached !== undefined ? cached : { exists: false };
          }
        }
      }
    }

    res.json({
      success: true,
      data: result,
      totalSearched: merchantOrderIds.length,
      totalFound: Object.values(result).filter((v) => v?.exists).length,
    });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to search Blitz orders', error: error.message });
  }
});

router.post('/add-to-existing-batch', async (req, res) => {
  try {
    const { orders, batchId, hubId, business, city, serviceType, driverId, project } = req.body;
    const credentials =
      driverId && project
        ? (await getCredentialsByDriverId(driverId, project)) || (await getBlitzCredentials(req))
        : await getBlitzCredentials(req);
    const accessToken = await loginToBlitz(credentials.username, credentials.password);
    tokenCache[credentials.username] = { token: accessToken, expiry: Date.now() + TOKEN_TTL_MS };

    const merchantOrderIds = orders.map((o) => o.merchant_order_id);
    const { needUpload, skipUpload } = await classifyOrdersByBlitzStatus(orders, accessToken);

    if (needUpload.length > 0) {
      try {
        await smartUpload(needUpload, hubId, business || 12, city || 9, serviceType || 2, credentials.username, credentials.password);
      } catch (uploadError) {
        return res.status(500).json({ success: false, message: `Upload failed: ${uploadError.message}` });
      }
      const recheck = await waitUntilOrdersAppearInBlitz(
        needUpload.map((o) => o.merchant_order_id), accessToken, 4, 5000
      );
      if (!recheck.success)
        return res.status(500).json({
          success: false,
          message: `Upload berhasil, namun ${recheck.missing.length} order belum muncul.`,
          missingOrders: recheck.missing,
        });
    }

    let validateResponse;
    try {
      validateResponse = await axios.post(
        BLITZ_VALIDATE_BATCH_URL,
        { batchId: parseInt(batchId), hub_id: hubId, sequence_type: 1, merchant_order_ids: merchantOrderIds },
        { headers: { Accept: 'application/json', Authorization: accessToken, 'Content-Type': 'application/json', bt: '2' }, timeout: 30000 }
      );
    } catch (e) {
      const d = e.response?.data;
      return res.status(e.response?.status || 500).json({
        success: false, message: extractBlitzErrorMessage(d) || e.message,
        blitz_error: d, validation_errors: extractValidationErrors(d),
      });
    }

    if (!validateResponse.data.result)
      return res.status(400).json({
        success: false, message: extractBlitzErrorMessage(validateResponse.data) || 'Validation failed',
        blitz_error: validateResponse.data, validation_errors: extractValidationErrors(validateResponse.data),
      });

    if (hasInvalidOrders(validateResponse.data))
      return res.status(400).json({
        success: false, message: 'Please remove invalid orders before creating the batch.',
        blitz_error: validateResponse.data, validation_errors: extractValidationErrors(validateResponse.data),
      });

    let addResponse;
    try {
      addResponse = await axios.post(
        BLITZ_ADD_BATCH_URL,
        { sequence_type: 1, batch_id: parseInt(batchId), merchant_order_ids: merchantOrderIds, hub_id: hubId },
        { headers: { Accept: 'application/json', Authorization: accessToken, 'Content-Type': 'application/json', bt: '2' }, timeout: 30000 }
      );
    } catch (e) {
      const d = e.response?.data;
      return res.status(e.response?.status || 500).json({
        success: false, message: `Failed to add to batch: ${extractBlitzErrorMessage(d) || e.message}`,
        blitz_error: d, validation_errors: extractValidationErrors(d),
      });
    }

    if (!addResponse.data.result)
      return res.status(400).json({
        success: false, message: extractBlitzErrorMessage(addResponse.data) || 'Add to batch failed',
        blitz_error: addResponse.data, validation_errors: extractValidationErrors(addResponse.data),
      });

    res.json({
      success: true, batchId,
      uploadedCount: needUpload.length,
      skippedCount: skipUpload.length,
      addedCount: merchantOrderIds.length,
      data: addResponse.data.data,
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

router.get('/batch-details/:batchId', async (req, res) => {
  try {
    const { batchId } = req.params;
    const accessToken = await getAccessToken(req);
    const response = await axios.get(`${BLITZ_BATCH_DETAILS_URL}/${batchId}`, {
      headers: { Accept: 'application/json', Authorization: accessToken, bt: '2' },
      timeout: 30000,
    });
    if (response.data.result && response.data.data) return res.json(response.data);
    res.status(404).json({ result: false, message: 'Batch not found' });
  } catch (error) {
    res.status(500).json({ result: false, message: 'Failed to get batch details', error: error.message });
  }
});

router.get('/active-batch/:driverId', async (req, res) => {
  try {
    const { driverId } = req.params;
    const accessToken = await getAccessToken(req);
    const now = new Date();
    const jakartaOffset = 7 * 60 * 60 * 1000;
    const jNow = new Date(now.getTime() + jakartaOffset);
    const tomorrow = new Date(jNow.getTime() + 24 * 60 * 60 * 1000).toISOString().split('T')[0];
    const sevenDaysAgo = new Date(jNow.getTime() - 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
    const response = await axios.get(`${BLITZ_DRIVER_PERFORMANCE_URL}/${driverId}`, {
      params: { sort: '-1', batchType: '', statusId: '', page: 1, offset: 100, term: '', createdFrom: sevenDaysAgo, createdTo: tomorrow },
      headers: { Accept: 'application/json', Authorization: accessToken },
      timeout: 30000,
    });
    if (!response.data?.result || !response.data?.data)
      return res.json({ success: true, batchId: null });
    const batches = response.data.data.driver_batch_performance_list;
    if (!Array.isArray(batches) || !batches.length)
      return res.json({ success: true, batchId: null });
    const activeBatch = batches.find((b) => b.assignment_status === 1);
    res.json({ success: true, batchId: activeBatch ? activeBatch.id : null });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to get active batch', error: error.message });
  }
});

router.get('/driver-attendance/:driverPhone', async (req, res) => {
  try {
    const { driverPhone } = req.params;
    const accessToken = await getAccessToken(req);
    const response = await axios.get(BLITZ_DRIVER_LIST_URL, {
      params: { sort: '-1', status: '1,2,8,3,4,5,6,7', attendance: '', page: 1, offset: 100, term: driverPhone, app_version_name: '', bank_info_provided: 'undefined', _t: Date.now() },
      headers: { Accept: 'application/json', Authorization: accessToken },
      timeout: 30000,
    });
    if (response.data.result && response.data.data?.driver_list_response?.length > 0)
      return res.json({ success: true, status: response.data.data.driver_list_response[0].drivers?.attendance_status || 'offline' });
    res.json({ success: true, status: 'offline' });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to get driver attendance', error: error.message, status: 'offline' });
  }
});

router.get('/driver-profile/:driverId', async (req, res) => {
  try {
    const { driverId } = req.params;
    const accessToken = await getAccessToken(req);
    const response = await axios.get(`${BLITZ_DRIVER_PROFILE_URL}/${driverId}`, {
      headers: { Accept: 'application/json', Authorization: accessToken },
      timeout: 30000,
    });
    if (response.data.result) return res.json({ success: true, data: response.data.data });
    res.json({ success: false, message: 'Driver profile not found' });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to get driver profile', error: error.message });
  }
});

router.post('/nearby-drivers', async (req, res) => {
  try {
    const { lat, lon } = req.body;
    const accessToken = await getAccessToken(req);
    const response = await axios.post(
      BLITZ_NEARBY_DRIVERS_URL,
      { lat: parseFloat(lat), lon: parseFloat(lon), radius: '20km', hub_ids: [], business_ids: [] },
      { headers: { Accept: 'application/json', Authorization: accessToken, 'Content-Type': 'application/json' }, timeout: 30000 }
    );
    if (response.data.result) return res.json({ success: true, data: response.data.data });
    res.json({ success: false, message: 'Failed to fetch nearby drivers' });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to get nearby drivers', error: error.message });
  }
});

router.post('/validate-batch-orders', async (req, res) => {
  try {
    const { sequenceType, batchId, merchantOrderIds, hubId } = req.body;
    const accessToken = await getAccessToken(req);
    const response = await axios.post(
      BLITZ_VALIDATE_BATCH_URL,
      { sequence_type: sequenceType, batch_id: batchId, merchant_order_ids: merchantOrderIds, hub_id: hubId },
      { headers: { Accept: 'application/json', Authorization: accessToken, 'Content-Type': 'application/json', bt: '2' }, timeout: 30000 }
    );
    res.json(response.data);
  } catch (error) {
    res.status(500).json({ result: false, message: 'Failed to validate batch orders', error: error.message });
  }
});

router.post('/add-batch-orders', async (req, res) => {
  try {
    const { sequenceType, batchId, merchantOrderIds, hubId } = req.body;
    const accessToken = await getAccessToken(req);
    try {
      const response = await axios.post(
        BLITZ_ADD_BATCH_URL,
        { sequence_type: sequenceType, batch_id: batchId, merchant_order_ids: merchantOrderIds, hub_id: hubId },
        { headers: { Accept: 'application/json', Authorization: accessToken, 'Content-Type': 'application/json', bt: '2' }, timeout: 30000 }
      );
      res.json(response.data);
    } catch (axiosError) {
      const d = axiosError.response?.data;
      res.status(axiosError.response?.status || 500).json({
        result: false, message: extractBlitzErrorMessage(d) || axiosError.message,
        blitz_error: d, validation_errors: extractValidationErrors(d),
      });
    }
  } catch (error) {
    res.status(500).json({ result: false, message: 'Failed to add batch orders', error: error.message });
  }
});

router.post('/create-batch-with-driver', async (req, res) => {
  try {
    const { orders, driverId, business, city, serviceType, hubId, coordinates, project, batchOnly } = req.body;
    const isBatchOnly = batchOnly === true || batchOnly === 'true';

    const credentials =
      driverId && project
        ? (await getCredentialsByDriverId(driverId, project)) || (await getBlitzCredentials(req))
        : await getBlitzCredentials(req);

    const accessToken = await loginToBlitz(credentials.username, credentials.password);
    tokenCache[credentials.username] = { token: accessToken, expiry: Date.now() + TOKEN_TTL_MS };

    const { needUpload, skipUpload } = await classifyOrdersByBlitzStatus(orders, accessToken);

    if (needUpload.length > 0) {
      try {
        await smartUpload(needUpload, hubId, business || 12, city || 9, serviceType || 2, credentials.username, credentials.password);
      } catch (uploadError) {
        return res.status(500).json({ success: false, message: 'Failed to upload missing orders', error: uploadError.message });
      }

      const recheck = await waitUntilOrdersAppearInBlitz(
        needUpload.map((o) => o.merchant_order_id), accessToken, 4, 5000
      );
      if (!recheck.success) {
        await sleep(5000);
      } else {
        await sleep(4000);
      }
    }

    const merchantOrderIds = orders.map((o) => o.merchant_order_id);

    try {
      const result = await executeBatchFlow(accessToken, merchantOrderIds, hubId, driverId, coordinates, isBatchOnly);

      if (project && driverId) {
        try {
          const driverDoc = await mongoose.connection.db
            .collection(`${project}_delivery`)
            .findOne({ driver_id: driverId.toString() });

          if (driverDoc) {
            const collection = mongoose.connection.db.collection(`${project}_merchant_orders`);
            const updateOps = orders.map((o) => ({
              updateOne: {
                filter: { merchant_order_id: o.merchant_order_id },
                update: {
                  $set: {
                    account_driver_name: driverDoc.driver_name || null,
                    account_driver_phone: driverDoc.driver_phone || null,
                  },
                },
              },
            }));
            if (updateOps.length > 0) {
              await collection.bulkWrite(updateOps, { ordered: false });
            }
          }
        } catch (e) {
          console.warn('[create-batch] account_driver update failed:', e.message);
        }
      }

      res.json({
        success: true,
        batchId: result.batchId,
        uploadedCount: needUpload.length,
        skippedCount: skipUpload.length,
        batchOnly: isBatchOnly,
        assigned: result.assigned,
        assignmentId: result.assignmentId || null,
      });
    } catch (flowError) {
      if (flowError.statusCode && flowError.body)
        return res.status(flowError.statusCode).json(flowError.body);
      throw flowError;
    }
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

router.post('/remove-order-from-batch', async (req, res) => {
  try {
    const { batchId, merchantOrderId, orderId, project } = req.body;
    const validationData = await getValidationDataByMerchantOrderId(project, merchantOrderId);
    if (!validationData)
      return res.status(400).json({ success: false, message: `Validation data not found for order ${merchantOrderId}.` });
    const hubId = validationData.business_hub;
    const accessToken = await getAccessToken(req);

    const validateResponse = await axios.post(
      `${BLITZ_REMOVE_VALIDATE_URL}/${batchId}`,
      { merchant_order_id: merchantOrderId },
      { headers: { Accept: 'application/json', Authorization: accessToken, 'Content-Type': 'application/json', bt: '2' }, timeout: 30000 }
    );
    if (validateResponse.status !== 200) throw new Error('Validation failed');

    const removeResponse = await axios.post(
      BLITZ_REMOVE_ORDER_URL,
      { sequence_type: 1, batch_id: batchId, merchant_order_ids: [merchantOrderId], hub_id: hubId },
      { headers: { Accept: 'application/json', Authorization: accessToken, 'Content-Type': 'application/json', bt: '2' }, timeout: 30000 }
    );
    if (removeResponse.status !== 200) throw new Error('Remove from Blitz failed');

    bulkOrderCache.delete(merchantOrderId);
    await mongoose.connection.db.collection(`${project}_merchant_orders`).updateOne(
      { _id: new mongoose.Types.ObjectId(orderId) },
      { $set: { assigned_to_driver_id: null, assigned_to_driver_name: null, assigned_to_driver_phone: null, assigned_at: null, assignment_status: 'unassigned', batch_id: null } }
    );
    res.json({ success: true, message: 'Order removed successfully' });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to remove order from batch', error: error.message });
  }
});

router.post('/refresh-token', async (req, res) => {
  try {
    const credentials = await getBlitzCredentials(req);
    delete tokenCache[credentials.username];
    const token = await loginToBlitz(credentials.username, credentials.password);
    tokenCache[credentials.username] = { token, expiry: Date.now() + TOKEN_TTL_MS };
    res.json({ success: true, message: 'Token refreshed successfully', expiresIn: '55 minutes' });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to refresh token', error: error.message });
  }
});

router.post('/clear-cache', async (req, res) => {
  try {
    const credentials = await getBlitzCredentials(req);
    if (tokenCache[credentials.username]) delete tokenCache[credentials.username];
    bulkOrderCache.clear();
    inFlightBatchKeys.clear();
    inFlightSearchRequests.clear();
    res.json({ success: true, message: `Cache cleared for ${credentials.username}` });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to clear cache', error: error.message });
  }
});

module.exports = router;