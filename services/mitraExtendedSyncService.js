const axios = require('axios');
const MitraExtended = require('../models/MitraExtended');

let fetchNikFromLark;
try {
  const larkController = require('../controllers/larkController');
  fetchNikFromLark = larkController.fetchNikFromLark;
  if (!fetchNikFromLark) throw new Error('fetchNikFromLark is not exported from larkController');
} catch (err) {
  console.error('Warning: larkController import failed:', err.message);
  fetchNikFromLark = async () => {
    console.warn('fetchNikFromLark unavailable, returning empty array');
    return [];
  };
}

const getRideblitzConfig = () => ({
  BASE_URL: process.env.RIDEBLITZ_BASE_URL || 'https://driver-api.rideblitz.id',
  AUTH_TOKEN: process.env.RIDEBLITZ_AUTH_TOKEN,
  BANK_API_URL: process.env.RIDEBLITZ_BANK_API_URL || 'https://user.rideblitz.id/v1/app/users/bank_detail/drivers',
  TIMEOUT: 15000,
  BATCH_SIZE: 200,
  CONCURRENT_REQUESTS: 5,
  PAGINATION_OFFSET: 200,
  MAX_PAGES: 500
});

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

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

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

const formatHubBusinessData = (data) => {
  if (!data || typeof data !== 'object' || Object.keys(data).length === 0) return '';
  return Object.entries(data).map(([id, name]) => `${name} (${id})`).join(', ');
};

const fetchAllDriversWithPagination = async (progressCallback, cancellationToken) => {
  const config = getRideblitzConfig();

  if (!config.AUTH_TOKEN) {
    throw new Error('RIDEBLITZ_AUTH_TOKEN environment variable is not set. Please configure it in your .env file.');
  }

  const allDrivers = [];
  let currentPage = 1;
  const OFFSET = config.PAGINATION_OFFSET;
  const MAX_PAGES = config.MAX_PAGES;

  cancellationToken.throwIfCancelled();

  progressCallback?.({
    stage: 'rideblitz_fetch',
    message: 'Fetching driver data from Rideblitz...',
    percentage: 5
  });

  while (currentPage <= MAX_PAGES) {
    if (cancellationToken.isCancelled) break;

    try {
      const source = axios.CancelToken.source();

      const response = await axios.get(`${config.BASE_URL}/v2/panel/driver-list`, {
        params: {
          sort: -1,
          status: [1, 2, 8, 3, 4, 5, 6, 7],
          attendance: '',
          page: currentPage,
          offset: OFFSET,
          term: '',
          app_version_name: '',
          bank_info_provided: 'undefined'
        },
        paramsSerializer: (params) => {
          const parts = [];
          for (const [key, value] of Object.entries(params)) {
            if (Array.isArray(value)) {
              value.forEach(v => parts.push(`${key}=${encodeURIComponent(v)}`));
            } else {
              parts.push(`${key}=${encodeURIComponent(value)}`);
            }
          }
          return parts.join('&');
        },
        headers: {
          Authorization: config.AUTH_TOKEN,
          Accept: 'application/json'
        },
        timeout: config.TIMEOUT,
        cancelToken: source.token
      });

      cancellationToken.throwIfCancelled();

      if (response.data?.data?.driver_list_response) {
        const drivers = response.data.data.driver_list_response;
        if (drivers.length > 0) {
          allDrivers.push(...drivers);
          progressCallback?.({
            stage: 'rideblitz_fetch',
            message: `Fetched ${allDrivers.length} drivers from Rideblitz...`,
            percentage: Math.min(15, 5 + currentPage * 0.3)
          });
          if (drivers.length < OFFSET) break;
          currentPage++;
          await sleep(200);
        } else {
          break;
        }
      } else {
        break;
      }
    } catch (error) {
      if (axios.isCancel(error) || error.isCancelled || cancellationToken.isCancelled) break;
      if (error.response?.status === 401) {
        throw new Error('Rideblitz API authentication failed. Token may be expired. Please update RIDEBLITZ_AUTH_TOKEN in environment variables.');
      }
      console.error(`Error fetching page ${currentPage}:`, error.message);
      break;
    }
  }

  console.log(`Total drivers fetched: ${allDrivers.length}`);
  return allDrivers;
};

const fetchBankDetails = async (userId, cancellationToken) => {
  const config = getRideblitzConfig();
  if (cancellationToken.isCancelled || !userId) return null;

  try {
    const response = await axios.get(`${config.BANK_API_URL}/${userId}`, {
      headers: {
        Authorization: config.AUTH_TOKEN,
        Accept: 'application/json'
      },
      timeout: config.TIMEOUT
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

const fetchDriverProfileFromRideblitz = async (driverId, userId, cancellationToken) => {
  const config = getRideblitzConfig();
  if (cancellationToken.isCancelled) return null;

  try {
    const response = await axios.get(`${config.BASE_URL}/panel/driver-profile/${driverId}`, {
      headers: {
        Authorization: config.AUTH_TOKEN,
        Accept: 'application/json'
      },
      timeout: config.TIMEOUT
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

    const bankDetails = await fetchBankDetails(userId, cancellationToken);

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

const processBatchProfiles = async (batchIds, driverList, larkData, cancellationToken) => {
  if (cancellationToken.isCancelled) return [];

  const results = [];
  const config = getRideblitzConfig();
  const concurrentLimit = config.CONCURRENT_REQUESTS;

  for (let i = 0; i < batchIds.length; i += concurrentLimit) {
    if (cancellationToken.isCancelled) break;

    const concurrentBatch = batchIds.slice(i, i + concurrentLimit);
    const promises = concurrentBatch.map(driverId => {
      const driverInfo = driverList.find(item => String(item.drivers?.id) === String(driverId));
      const userId = driverInfo?.drivers?.user_id;
      return fetchDriverProfileFromRideblitz(driverId, userId, cancellationToken);
    });

    const profiles = await Promise.all(promises);
    if (cancellationToken.isCancelled) break;

    for (const profile of profiles) {
      if (!profile) continue;

      const driverInfo = driverList.find(item => String(item.drivers?.id) === String(profile.driver_id));
      const driverData = driverInfo?.drivers || {};
      const accountState = driverData.account_state || {};

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
        hubs: formatHubBusinessData(profile.hub_data),
        businesses: formatHubBusinessData(profile.business_data),
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

const manualSyncMitraExtended = async (syncId, progressCallback) => {
  const config = getRideblitzConfig();

  if (!config.AUTH_TOKEN) {
    throw new Error('RIDEBLITZ_AUTH_TOKEN is not configured. Please add it to your .env file: RIDEBLITZ_AUTH_TOKEN=your_token_here');
  }

  const startTime = Date.now();
  const cancellationToken = new CancellationToken();
  activeCancellationTokens.set(syncId, cancellationToken);

  try {
    cancellationToken.throwIfCancelled();
    progressCallback?.({ stage: 'init', message: 'Initializing sync process...', percentage: 0 });

    const driverList = await fetchAllDriversWithPagination(progressCallback, cancellationToken);
    if (cancellationToken.isCancelled) throw Object.assign(new Error('Sync cancelled during driver fetch'), { isCancelled: true });

    progressCallback?.({ stage: 'lark_fetch', message: 'Fetching data from Larksuite...', percentage: 20 });

    const larkData = await fetchNikFromLark();
    if (cancellationToken.isCancelled) throw Object.assign(new Error('Sync cancelled during Lark fetch'), { isCancelled: true });

    progressCallback?.({ stage: 'validation', message: 'Validating and transforming data...', percentage: 30 });

    const driverIds = driverList.map(item => item.drivers?.id).filter(id => id);
    cancellationToken.throwIfCancelled();

    await MitraExtended.deleteMany({});

    progressCallback?.({ stage: 'processing', message: 'Processing driver profiles...', percentage: 35 });

    let processedCount = 0;
    let successCount = 0;
    let larkMatchCount = 0;
    const totalDrivers = driverIds.length;
    const BATCH_SIZE = config.BATCH_SIZE;

    for (let i = 0; i < totalDrivers; i += BATCH_SIZE) {
      if (cancellationToken.isCancelled) {
        throw Object.assign(new Error(`Sync cancelled - ${successCount} records saved`), { isCancelled: true, successCount });
      }

      const batchIds = driverIds.slice(i, i + BATCH_SIZE);
      const batchProfiles = await processBatchProfiles(batchIds, driverList, larkData, cancellationToken);

      if (cancellationToken.isCancelled) {
        throw Object.assign(new Error(`Sync cancelled - ${successCount} records saved`), { isCancelled: true, successCount });
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
        await sleep(300);
      }
    }

    if (cancellationToken.isCancelled) {
      throw Object.assign(new Error(`Sync cancelled - ${successCount} records saved`), { isCancelled: true, successCount });
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
    if (error.isCancelled || cancellationToken.isCancelled || error.message.includes('cancel')) {
      const savedCount = error.successCount || 0;
      throw Object.assign(new Error(`Sync cancelled by user - ${savedCount} records saved`), { isCancelled: true, successCount: savedCount });
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