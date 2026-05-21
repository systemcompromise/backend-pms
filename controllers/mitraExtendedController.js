const MitraExtended = require('../models/MitraExtended');
const cacheWarmer = require('../services/cacheWarmer');
const {
  manualSyncMitraExtended,
  cancelSync,
  cancelAllSyncs
} = require('../services/mitraExtendedSyncService');

const getBulkMitraExtendedData = async (req, res) => {
  const startTime = Date.now();

  req.on('close', () => {
    if (!res.writableEnded) {
      console.log('Client disconnected, aborting query...');
    }
  });

  try {
    let cachedData = cacheWarmer.getCachedData('mitraExtended');

    if (!cachedData) {
      cachedData = await MitraExtended.find({})
        .select('-__v')
        .lean()
        .hint({ driver_id: 1 })
        .maxTimeMS(30000)
        .exec();

      if (cachedData && cachedData.length > 0) {
        cacheWarmer.setCachedData('mitraExtended', cachedData);
      }
    }

    if (req.aborted || res.writableEnded) return;

    if (!cachedData || cachedData.length === 0) {
      return res.status(200).json({
        success: true,
        data: [],
        pagination: { currentPage: 1, totalPages: 0, totalRecords: 0, hasNextPage: false },
        meta: { queryTime: Date.now() - startTime, timestamp: new Date().toISOString(), source: 'database' }
      });
    }

    const duration = Date.now() - startTime;

    return res.status(200).json({
      success: true,
      data: cachedData,
      pagination: { currentPage: 1, totalPages: 1, totalRecords: cachedData.length, hasNextPage: false },
      meta: { queryTime: duration, timestamp: new Date().toISOString(), source: 'cache' }
    });

  } catch (error) {
    if (req.aborted || res.writableEnded) return;
    const duration = Date.now() - startTime;
    res.status(500).json({
      success: false,
      message: 'Failed to fetch mitra extended data',
      error: error.message,
      queryTime: duration
    });
  }
};

const manualSyncController = async (req, res) => {
  const syncId = `sync_${Date.now()}`;

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('X-Accel-Buffering', 'no');
  res.flushHeaders();

  let streamClosed = false;
  let lastSavedCount = 0;

  const sendProgress = (data) => {
    if (!streamClosed && !res.writableEnded) {
      try {
        res.write(`data: ${JSON.stringify(data)}\n\n`);
      } catch (error) {
        streamClosed = true;
      }
    }
  };

  const closeStream = () => {
    if (streamClosed || res.writableEnded) return;
    streamClosed = true;
    try {
      res.end();
    } catch (error) {
      console.log(`Error closing stream: ${error.message}`);
    }
  };

  req.on('close', () => {
    cancelSync(syncId);
    streamClosed = true;
  });

  req.on('error', (error) => {
    cancelSync(syncId);
    streamClosed = true;
  });

  if (req.user.role !== 'owner') {
    sendProgress({
      type: 'error',
      stage: 'error',
      message: 'Access denied. Only owner role can perform manual sync.',
      percentage: 0,
      syncId
    });
    return closeStream();
  }

  try {
    sendProgress({
      type: 'progress',
      stage: 'init',
      message: 'Starting sync process...',
      percentage: 0,
      syncId
    });

    const result = await manualSyncMitraExtended(syncId, (progress) => {
      if (!streamClosed) {
        if (progress.successCount) lastSavedCount = progress.successCount;
        sendProgress({ type: 'progress', ...progress, syncId });
      }
    });

    if (streamClosed) return;

    cacheWarmer.clearCache('mitraExtended');
    await cacheWarmer.warmMitraExtendedCache();

    sendProgress({
      type: 'complete',
      stage: 'complete',
      message: 'Sync completed successfully',
      percentage: 100,
      syncId
    });

    setTimeout(() => closeStream(), 500);

  } catch (error) {
    const isCancellation = error.isCancelled || error.message.includes('cancel') || streamClosed;

    if (!streamClosed && !res.writableEnded) {
      cacheWarmer.clearCache('mitraExtended');
      cacheWarmer.warmMitraExtendedCache().catch(() => {});

      if (isCancellation) {
        sendProgress({
          type: 'cancelled',
          stage: 'cancelled',
          message: `Sync cancelled - ${lastSavedCount} records saved before cancellation`,
          percentage: 0,
          syncId,
          savedCount: lastSavedCount
        });
      } else {
        sendProgress({
          type: 'error',
          stage: 'error',
          message: error.message,
          error: error.message,
          syncId
        });
      }

      setTimeout(() => closeStream(), 500);
    } else {
      cacheWarmer.clearCache('mitraExtended');
      cacheWarmer.warmMitraExtendedCache().catch(() => {});
    }
  }
};

const cancelSyncEndpoint = async (req, res) => {
  try {
    if (req.user.role !== 'owner') {
      return res.status(403).json({ success: false, message: 'Access denied. Only owner role can cancel sync.' });
    }

    const { syncId } = req.body;

    if (syncId) {
      const cancelled = cancelSync(syncId);
      cacheWarmer.clearCache('mitraExtended');

      if (cancelled) {
        return res.status(200).json({ success: true, message: `Sync ${syncId} cancelled successfully`, syncId });
      } else {
        return res.status(404).json({ success: false, message: `Sync ${syncId} not found or already completed`, syncId });
      }
    } else {
      const cancelledCount = cancelAllSyncs();
      cacheWarmer.clearCache('mitraExtended');
      return res.status(200).json({ success: true, message: `Cancelled ${cancelledCount} active sync process(es)`, cancelledCount });
    }
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to cancel sync', error: error.message });
  }
};

const getExtendedDataByDriverId = async (req, res) => {
  try {
    const { driver_id } = req.params;
    if (!driver_id) {
      return res.status(400).json({ success: false, message: 'Driver ID is required' });
    }

    const extendedData = await MitraExtended.findOne({ driver_id }).lean().exec();
    if (!extendedData) {
      return res.status(404).json({ success: false, message: 'Extended data not found' });
    }

    res.status(200).json({ success: true, data: extendedData });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to fetch extended data', error: error.message });
  }
};

const createOrUpdateExtendedData = async (req, res) => {
  try {
    const { driver_id } = req.params;
    const updateData = req.body;

    if (!driver_id) {
      return res.status(400).json({ success: false, message: 'Driver ID is required' });
    }

    const extendedData = await MitraExtended.findOneAndUpdate(
      { driver_id },
      { $set: { ...updateData, driver_id, updated_at: new Date() } },
      { new: true, upsert: true, runValidators: true, lean: true }
    );

    cacheWarmer.clearCache('mitraExtended');

    res.status(200).json({ success: true, message: 'Extended data saved successfully', data: extendedData });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to save extended data', error: error.message });
  }
};

const deleteExtendedData = async (req, res) => {
  try {
    const { driver_id } = req.params;
    if (!driver_id) {
      return res.status(400).json({ success: false, message: 'Driver ID is required' });
    }

    const deletedData = await MitraExtended.findOneAndDelete({ driver_id }).lean();
    if (!deletedData) {
      return res.status(404).json({ success: false, message: 'Extended data not found' });
    }

    cacheWarmer.clearCache('mitraExtended');

    res.status(200).json({ success: true, message: 'Extended data deleted successfully', data: deletedData });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Failed to delete extended data', error: error.message });
  }
};

module.exports = {
  getExtendedDataByDriverId,
  createOrUpdateExtendedData,
  deleteExtendedData,
  getBulkMitraExtendedData,
  manualSyncController,
  cancelSyncEndpoint
};