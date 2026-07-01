/* =========================================================================
 * 🟢 MACHINE CAPABILITY SERVICE
 *
 * Helper untuk read & query capability mesin dari sheet M_MC.
 * Hasil di-cache 5 menit untuk performance (CacheService).
 *
 * Schema M_MC:
 *   MC_No | MC_Name | Type | Max_Lebar | Max_OUT_Count | Is_Active | Display_Order
 *
 * Public Functions:
 *   - getMachineCapability(mcNo)            → 1 mesin spesifik
 *   - getMachinesByType(type, activeOnly)   → list mesin per Type (SHR/CTL/SLT)
 *   - getAllMachines(activeOnly)            → list semua mesin
 *   - clearMachineCache()                   → invalidate cache (panggil kalau M_MC di-update)
 * ========================================================================= */

var MC_CACHE_KEY     = 'm_mc_capability_v1';
var MC_CACHE_TTL_SEC = 300;  // 5 menit

/**
 * Internal: Load semua capability dari M_MC sheet, sorted by Display_Order.
 * Pakai cache supaya tidak read sheet tiap call.
 */
function _loadMachineCapabilityFromSheet() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get(MC_CACHE_KEY);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      Logger.log('MC cache parse error, reload: ' + e);
    }
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('M_MC');
  if (!sheet) {
    Logger.log('⚠️ Sheet M_MC tidak ditemukan');
    return [];
  }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const hdr = data[0].map(function(h){ return String(h).trim(); });
  const I = {
    mcNo    : hdr.indexOf('MC_No'),
    mcName  : hdr.indexOf('MC_Name'),
    type    : hdr.indexOf('Type'),
    maxLbr  : hdr.indexOf('Max_Lebar'),
    maxOut  : hdr.indexOf('Max_OUT_Count'),
    active  : hdr.indexOf('Is_Active'),
    order   : hdr.indexOf('Display_Order')
  };

  // Validate required columns
  const missing = [];
  Object.keys(I).forEach(function(k){ if (I[k] === -1) missing.push(k); });
  if (missing.length > 0) {
    Logger.log('⚠️ M_MC kolom hilang: ' + missing.join(', '));
    return [];
  }

  const list = [];
  for (let i = 1; i < data.length; i++) {
    const mcNo = String(data[i][I.mcNo] || '').trim();
    if (!mcNo) continue;
    list.push({
      mc_no          : mcNo,
      mc_name        : String(data[i][I.mcName] || '').trim() || mcNo,
      type           : String(data[i][I.type]   || '').trim().toUpperCase(),
      max_lebar      : Number(data[i][I.maxLbr]) || 0,
      max_out_count  : Number(data[i][I.maxOut]) || 1,
      is_active      : data[i][I.active] === true || String(data[i][I.active]).toUpperCase() === 'TRUE',
      display_order  : Number(data[i][I.order])  || 999
    });
  }

  // Sort by Display_Order
  list.sort(function(a, b){ return a.display_order - b.display_order; });

  // Cache (TTL 5 menit)
  try {
    cache.put(MC_CACHE_KEY, JSON.stringify(list), MC_CACHE_TTL_SEC);
  } catch (e) {
    Logger.log('MC cache put error (mungkin too large): ' + e);
  }

  return list;
}

/**
 * Get capability spesifik untuk 1 mesin.
 * @param {string} mcNo - e.g. 'SHR-01'
 * @return {Object|null} {mc_no, mc_name, type, max_lebar, max_out_count, is_active, display_order} | null
 */
function getMachineCapability(mcNo) {
  if (!mcNo) return null;
  const list   = _loadMachineCapabilityFromSheet();
  const target = String(mcNo).trim().toUpperCase();
  for (let i = 0; i < list.length; i++) {
    if (list[i].mc_no.toUpperCase() === target) return list[i];
  }
  return null;
}

/**
 * Get list mesin per Type. Berguna untuk populate dropdown.
 * @param {string} type - 'CTL' / 'SHR' / 'SLT'
 * @param {boolean} activeOnly - kalau true, filter Is_Active=TRUE saja
 * @return {Array<Object>} list sorted by Display_Order
 */
function getMachinesByType(type, activeOnly) {
  if (!type) return [];
  const list = _loadMachineCapabilityFromSheet();
  const want = String(type).trim().toUpperCase();
  return list.filter(function(m) {
    if (m.type !== want) return false;
    if (activeOnly && !m.is_active) return false;
    return true;
  });
}

/**
 * Get semua mesin. Optional filter by Is_Active.
 * @param {boolean} activeOnly
 * @return {Array<Object>}
 */
function getAllMachines(activeOnly) {
  const list = _loadMachineCapabilityFromSheet();
  if (activeOnly) return list.filter(function(m){ return m.is_active; });
  return list;
}

/**
 * Invalidate cache. Panggil setelah M_MC di-update manual via sheet.
 * Bisa juga di-set sebagai trigger onEdit (advanced).
 */
function clearMachineCache() {
  CacheService.getScriptCache().remove(MC_CACHE_KEY);
  Logger.log('✅ Machine capability cache cleared');
}