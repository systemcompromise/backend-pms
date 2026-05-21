const axios = require('axios');
const MitraExtended = require('../models/MitraExtended');

let fetchNikFromLark;
try {
  const larkController = require('../controllers/larkController');
  fetchNikFromLark = larkController.fetchNikFromLark;
  if (!fetchNikFromLark) throw new Error('fetchNikFromLark not exported');
} catch (err) {
  console.error('Warning: larkController import failed:', err.message);
  fetchNikFromLark = async () => [];
}

const RIDEBLITZ_BASE_URL = process.env.RIDEBLITZ_BASE_URL || 'https://driver-api.rideblitz.id';
const RIDEBLITZ_BANK_API_URL = process.env.RIDEBLITZ_BANK_API_URL || 'https://user.rideblitz.id/v1/app/users/bank_detail/drivers';
const RIDEBLITZ_USERNAME = process.env.RIDEBLITZ_USERNAME || 'septa';
const RIDEBLITZ_PASSWORD = process.env.RIDEBLITZ_PASSWORD || 'Blitz#Septa_26';

const PAGINATION_OFFSET = 100;
const MAX_PAGES = 500;
const REQUEST_TIMEOUT = 30000;
const BATCH_SIZE = 200;
const CONCURRENT_REQUESTS = 5;

let cachedToken = null;

const loginToRideblitz = async () => {
  if (!RIDEBLITZ_USERNAME || !RIDEBLITZ_PASSWORD) {
    throw new Error('Rideblitz credentials not configured.');
  }

  const response = await fetch(`${RIDEBLITZ_BASE_URL}/panel/login`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    body: JSON.stringify({
      username: RIDEBLITZ_USERNAME,
      password: RIDEBLITZ_PASSWORD
    })
  });

  if (!response.ok) {
    if (response.status === 401) throw new Error('Rideblitz login failed: Invalid credentials');
    throw new Error(`Rideblitz login failed: HTTP ${response.status}`);
  }

  const data = await response.json();
  if (!data?.result || !data?.data?.access_token) {
    throw new Error('Rideblitz login failed: missing access_token');
  }

  cachedToken = `Bearer ${data.data.access_token}`;
  return cachedToken;
};

const getValidToken = async () => {
  if (!cachedToken) return await loginToRideblitz();
  return cachedToken;
};

const invalidateToken = () => {
  cachedToken = null;
};

const buildDriverListUrl = (page) => {
  const params = new URLSearchParams();
  params.append('sort', '-1');
  [1, 2, 8, 3, 4, 5, 6, 7].forEach(s => params.append('status', s));
  params.append('attendance', '');
  params.append('page', page);
  params.append('offset', PAGINATION_OFFSET);
  params.append('term', '');
  params.append('app_version_name', '');
  params.append('bank_info_provided', 'undefined');
  return `${RIDEBLITZ_BASE_URL}/v2/panel/driver-list?${params.toString()}`;
};

const fetchDriverListPage = async (page, token) => {
  const url = buildDriverListUrl(page);

  const response = await axios.get(url, {
    headers: {
      'Authorization': token,
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    },
    timeout: REQUEST_TIMEOUT
  });

  return response.data?.data?.driver_list_response || null;
};

const fetchDriverListPageWithRetry = async (page, token, cancellationToken) => {
  cancellationToken.throwIfCancelled();

  try {
    const drivers = await fetchDriverListPage(page, token);
    return { drivers, token };
  } catch (error) {
    if (error.response?.status === 401) {
      invalidateToken();
      const freshToken = await loginToRideblitz();
      cancellationToken.throwIfCancelled();
      const drivers = await fetchDriverListPage(page, freshToken);
      return { drivers, token: freshToken };
    }
    throw error;
  }
};

const fetchAllDriversWithPagination = async (progressCallback, cancellationToken) => {
  const allDrivers = [];
  let currentPage = 1;

  cancellationToken.throwIfCancelled();

  progressCallback?.({
    stage: 'auth',
    message: 'Authenticating with Rideblitz...',
    percentage: 2
  });

  let activeToken = await loginToRideblitz();

  progressCallback?.({
    stage: 'rideblitz_fetch',
    message: 'Fetching driver data from Rideblitz... (Page 1)',
    percentage: 5
  });

  while (currentPage <= MAX_PAGES) {
    cancellationToken.throwIfCancelled();

    let result;
    try {
      result = await fetchDriverListPageWithRetry(currentPage, activeToken, cancellationToken);
    } catch (error) {
      if (error.isCancelled || cancellationToken.isCancelled) break;

      if (error.response?.status === 403) {
        throw new Error('Rideblitz API access denied. Check account permissions.');
      }

      if (currentPage === 1) {
        throw new Error(`Failed to fetch drivers from Rideblitz: ${error.message}`);
      }

      break;
    }

    const { drivers, token: returnedToken } = result;
    activeToken = returnedToken;

    if (!drivers || !Array.isArray(drivers) || drivers.length === 0) {
      break;
    }

    allDrivers.push(...drivers);

    progressCallback?.({
      stage: 'rideblitz_fetch',
      message: `Fetching driver data... Page ${currentPage} — ${allDrivers.length} drivers collected`,
      percentage: Math.min(18, 5 + currentPage * 0.5),
      currentPage,
      totalCollected: allDrivers.length
    });

    if (drivers.length < PAGINATION_OFFSET) {
      break;
    }

    currentPage++;
    await new Promise(r => setTimeout(r, 200));
  }

  return allDrivers;
};

const fetchBankDetails = async (userId, token, cancellationToken) => {
  if (cancellationToken.isCancelled || !userId) return null;

  try {
    const response = await axios.get(`${RIDEBLITZ_BANK_API_URL}/${userId}`, {
      headers: {
        'Authorization': token,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      timeout: REQUEST_TIMEOUT
    });

    if (cancellationToken.isCancelled) return null;

    if (response.data?.result && response.data?.data) {
      return {
        bank_name: response.data.data.bank || '',
        bank_account_number: response.data.data.account_number || '',
        bank_account_holder: response.data.data.beneficiary_name || ''
      };
    }
    return null;
  } catch {
    return null;
  }
};

const fetchDriverProfileFromRideblitz = async (driverId, userId, token, cancellationToken) => {
  if (cancellationToken.isCancelled) return null;

  try {
    const response = await axios.get(`${RIDEBLITZ_BASE_URL}/panel/driver-profile/${driverId}`, {
      headers: {
        'Authorization': token,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      timeout: REQUEST_TIMEOUT
    });

    if (cancellationToken.isCancelled) return null;
    if (!response.data?.result || !response.data?.data) return null;

    const data = response.data.data;
    const driverProfile = data.driver_profile || {};
    const documents = driverProfile.documents || [];
    const currentCoords = data.current_cordinates || {};
    const businessHub = data.business_hub || {};

    const ktpDoc = documents.find(d => d.fields?.key === 'ktp');
    const simDoc = documents.find(d => d.fields?.key === 'sim');

    const bankDetails = await fetchBankDetails(userId, token, cancellationToken);

    return {
      driver_id: String(driverId),
      current_lat: currentCoords.lat || null,
      current_lon: currentCoords.lon || null,
      nik: ktpDoc?.fields?.value?.nik || '',
      sim_number: simDoc?.fields?.value?.sim || '',
      sim_expiry: simDoc?.fields?.value?.expiry_date || '',
      bank_name: bankDetails?.bank_name || '',
      bank_account_holder: bankDetails?.bank_account_holder || '',
      bank_account_number: bankDetails?.bank_account_number || '',
      hub_data: businessHub.hub_data || {},
      business_data: businessHub.business_data || {}
    };
  } catch {
    return null;
  }
};

const matchNikWithLarkData = (driverProfile, larkDataArray) => {
  if (!driverProfile.nik || !larkDataArray || larkDataArray.length === 0) return null;
  const cleanNik = String(driverProfile.nik).trim();
  if (!cleanNik || cleanNik === '-') return null;
  const matchedLarkRecord = larkDataArray.find(larkItem => String(larkItem.nik || '').trim() === cleanNik);
  if (!matchedLarkRecord) return null;
  return {
    lark_tanggal_keluar_unit: matchedLarkRecord.tanggal_keluar_unit || '',
    lark_nomor_plat: matchedLarkRecord.plat_nomor || '',
    lark_merk_unit: matchedLarkRecord.merk_unit || '',
    lark_alamat: matchedLarkRecord.alamat || '',
    lark_tanggal_pengembalian_unit: matchedLarkRecord.tanggal_pengembalian_unit || '',
    lark_lama_pemakaian: matchedLarkRecord.lama_pemakaian || '',
    lark_status: matchedLarkRecord.status || '',
    lark_matched_at: new Date()
  };
};

const getStatusDisplay = (status) => {
  const statusMap = {
    registered: 'Registered',
    active: 'Active',
    pending: 'Pending Verification',
    new: 'New',
    inactive: 'Inactive',
    banned: 'Banned'
  };
  return statusMap[status?.toLowerCase()] || status || '-';
};

const formatDateTime = (dateString) => {
  if (!dateString) return '-';
  try {
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return '-';
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    const hours = date.getHours();
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const ampm = hours >= 12 ? 'PM' : 'AM';
    const formattedHours = hours % 12 || 12;
    return `${day}/${month}/${year} ${formattedHours}:${minutes}${ampm}`;
  } catch {
    return '-';
  }
};

const processBatchProfiles = async (batchIds, driverList, larkData, token, cancellationToken) => {
  if (cancellationToken.isCancelled) return [];

  const results = [];

  for (let i = 0; i < batchIds.length; i += CONCURRENT_REQUESTS) {
    if (cancellationToken.isCancelled) break;

    const concurrentBatch = batchIds.slice(i, i + CONCURRENT_REQUESTS);
    const promises = concurrentBatch.map(driverId => {
      const driverInfo = driverList.find(item => String(item.drivers?.id) === String(driverId));
      const userId = driverInfo?.drivers?.user_id;
      return fetchDriverProfileFromRideblitz(driverId, userId, token, cancellationToken);
    });

    const profiles = await Promise.all(promises);
    if (cancellationToken.isCancelled) break;

    for (const profile of profiles) {
      if (!profile) continue;

      const driverInfo = driverList.find(item => String(item.drivers?.id) === String(profile.driver_id));
      const driverData = driverInfo?.drivers || {};
      const accountState = driverData.account_state || {};
      const larkMatch = matchNikWithLarkData(profile, larkData);

      const updateData = {
        driver_id: profile.driver_id,
        name: driverData.name || '',
        phone_number: driverData.phone_number || '',
        city: driverData.city_name || '',
        status: getStatusDisplay(accountState.status),
        attendance: driverData.attendance_status || '',
        otp: driverData.otp || '',
        bank_info_provided: driverData.bank_info_provided || false,
        app_version_name: driverData.app_version_name || '',
        app_version_code: driverData.app_version || '',
        app_android_version: driverData.app_android_version || '',
        android_version: driverData.android_version || '',
        last_active: formatDateTime(driverData.last_active),
        registered_at: formatDateTime(driverInfo?.registered_at),
        hubs: profile.hub_data && Object.keys(profile.hub_data).length > 0
          ? Object.entries(profile.hub_data).map(([id, name]) => `${name} (${id})`).join(', ')
          : '',
        businesses: profile.business_data && Object.keys(profile.business_data).length > 0
          ? Object.entries(profile.business_data).map(([id, name]) => `${name} (${id})`).join(', ')
          : '',
        reason: driverData.reason || '',
        current_lat: profile.current_lat,
        current_lon: profile.current_lon,
        nik: profile.nik,
        sim_number: profile.sim_number,
        sim_expiry: profile.sim_expiry,
        bank_name: profile.bank_name,
        bank_account_holder: profile.bank_account_holder,
        bank_account_number: profile.bank_account_number,
        hub_data: profile.hub_data,
        business_data: profile.business_data,
        lark_tanggal_keluar_unit: '',
        lark_nomor_plat: '',
        lark_merk_unit: '',
        lark_alamat: '',
        lark_tanggal_pengembalian_unit: '',
        lark_lama_pemakaian: '',
        lark_status: '',
        updated_at: new Date()
      };

      if (larkMatch) Object.assign(updateData, larkMatch);
      results.push(updateData);
    }
  }

  return results;
};

const saveBatchToDatabase = async (batchProfiles, cancellationToken) => {
  if (!batchProfiles || batchProfiles.length === 0) return 0;

  const MONGO_CHUNK_SIZE = 500;
  let totalSaved = 0;

  for (let i = 0; i < batchProfiles.length; i += MONGO_CHUNK_SIZE) {
    if (cancellationToken.isCancelled) break;

    const chunk = batchProfiles.slice(i, i + MONGO_CHUNK_SIZE);
    const bulkOps = chunk.map(profile => ({
      updateOne: {
        filter: { driver_id: profile.driver_id },
        update: { $set: profile },
        upsert: true
      }
    }));

    try {
      const result = await MitraExtended.bulkWrite(bulkOps, { ordered: false });
      totalSaved += result.upsertedCount + result.modifiedCount;
    } catch (bulkError) {
      console.error(`Chunk save error (offset ${i}):`, bulkError.message);
    }
  }

  return totalSaved;
};

class CancellationToken {
  constructor() {
    this.cancelled = false;
    this.reason = null;
  }

  cancel(reason = 'Cancelled by user') {
    if (this.cancelled) return;
    this.cancelled = true;
    this.reason = reason;
  }

  throwIfCancelled() {
    if (this.cancelled) {
      const error = new Error(this.reason || 'Operation cancelled');
      error.isCancelled = true;
      throw error;
    }
  }

  get isCancelled() {
    return this.cancelled;
  }
}

const activeCancellationTokens = new Map();

const manualSyncMitraExtended = async (syncId, progressCallback) => {
  const startTime = Date.now();
  const cancellationToken = new CancellationToken();
  activeCancellationTokens.set(syncId, cancellationToken);

  try {
    cancellationToken.throwIfCancelled();
    progressCallback?.({ stage: 'init', message: 'Initializing sync process...', percentage: 0 });

    const driverList = await fetchAllDriversWithPagination(progressCallback, cancellationToken);

    if (cancellationToken.isCancelled) {
      throw Object.assign(new Error('Sync cancelled during driver fetch'), { isCancelled: true });
    }

    if (!driverList || driverList.length === 0) {
      throw new Error('No drivers fetched from Rideblitz. Check credentials and connectivity.');
    }

    progressCallback?.({ stage: 'lark_fetch', message: 'Fetching data from Larksuite...', percentage: 20 });

    let larkData = [];
    try {
      larkData = await fetchNikFromLark();
    } catch (larkError) {
      console.warn('Lark data fetch failed, continuing without Lark matching:', larkError.message);
      larkData = [];
    }

    if (cancellationToken.isCancelled) {
      throw Object.assign(new Error('Sync cancelled during Lark fetch'), { isCancelled: true });
    }

    progressCallback?.({ stage: 'validation', message: 'Validating and transforming data...', percentage: 30 });

    const driverIds = driverList.map(item => item.drivers?.id).filter(id => id);
    cancellationToken.throwIfCancelled();

    await MitraExtended.deleteMany({});

    progressCallback?.({ stage: 'processing', message: 'Processing driver profiles...', percentage: 35 });

    const activeToken = await getValidToken();
    let processedCount = 0;
    let successCount = 0;
    let larkMatchCount = 0;
    const totalDrivers = driverIds.length;

    for (let i = 0; i < totalDrivers; i += BATCH_SIZE) {
      if (cancellationToken.isCancelled) {
        throw Object.assign(
          new Error(`Sync cancelled - ${successCount} records saved`),
          { isCancelled: true, successCount }
        );
      }

      const batchIds = driverIds.slice(i, i + BATCH_SIZE);
      const batchProfiles = await processBatchProfiles(batchIds, driverList, larkData, activeToken, cancellationToken);

      if (cancellationToken.isCancelled) {
        throw Object.assign(
          new Error(`Sync cancelled - ${successCount} records saved`),
          { isCancelled: true, successCount }
        );
      }

      if (batchProfiles.length > 0) {
        const batchLarkMatches = batchProfiles.filter(p => p.lark_matched_at).length;
        const batchSaved = await saveBatchToDatabase(batchProfiles, cancellationToken);
        successCount += batchSaved;
        larkMatchCount += batchLarkMatches;
      }

      processedCount += batchIds.length;
      const progressPercent = 35 + Math.round((processedCount / totalDrivers) * 60);

      if (!cancellationToken.isCancelled) {
        progressCallback?.({
          stage: 'saving',
          message: `Saving data: ${processedCount}/${totalDrivers} records processed...`,
          percentage: progressPercent,
          successCount
        });
      }

      if (i + BATCH_SIZE < totalDrivers && !cancellationToken.isCancelled) {
        await new Promise(r => setTimeout(r, 300));
      }
    }

    if (cancellationToken.isCancelled) {
      throw Object.assign(
        new Error(`Sync cancelled - ${successCount} records saved`),
        { isCancelled: true, successCount }
      );
    }

    progressCallback?.({ stage: 'finalizing', message: 'Finalizing sync process...', percentage: 95 });

    const duration = Date.now() - startTime;
    const summary = {
      totalDrivers: driverIds.length,
      successCount,
      larkMatchCount,
      durationMs: duration,
      durationMinutes: (duration / 1000 / 60).toFixed(2),
      timestamp: new Date().toISOString(),
      aborted: false
    };

    progressCallback?.({
      stage: 'complete',
      message: `Sync completed: ${summary.successCount}/${summary.totalDrivers} records saved`,
      percentage: 100
    });

    return summary;

  } catch (error) {
    if (error.isCancelled || cancellationToken.isCancelled) {
      const savedCount = error.successCount || 0;
      throw Object.assign(
        new Error(`Sync cancelled by user - ${savedCount} records saved`),
        { isCancelled: true, successCount: savedCount }
      );
    }
    throw error;
  } finally {
    activeCancellationTokens.delete(syncId);
  }
};

const cancelSync = (syncId) => {
  const cancellationToken = activeCancellationTokens.get(syncId);
  if (cancellationToken) {
    cancellationToken.cancel('Cancelled by user request');
    return true;
  }
  return false;
};

const cancelAllSyncs = () => {
  let cancelledCount = 0;
  for (const [, cancellationToken] of activeCancellationTokens.entries()) {
    cancellationToken.cancel('All syncs cancelled');
    cancelledCount++;
  }
  activeCancellationTokens.clear();
  return cancelledCount;
};

module.exports = {
  manualSyncMitraExtended,
  cancelSync,
  cancelAllSyncs
};