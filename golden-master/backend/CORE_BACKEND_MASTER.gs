/* see attached complete Golden Master content from production baseline, normalized for Store Builder */

/**
 * ============================================================
 * STORE-AWARE CORE CONFIG
 * ============================================================
 * Store Builder writes STORE_CONFIG_JSON into Script Properties
 * for each Store Instance. Secrets remain separate properties.
 */
function getStoreConfig_() {
  const props = SpreadsheetApp.getActiveSpreadsheet() ? PropertiesService.getScriptProperties() : PropertiesService.getScriptProperties();
  const raw = String(props.getProperty('STORE_CONFIG_JSON') || '').trim();
  let config = {};
  if (raw) {
    try { config = JSON.parse(raw) || {}; }
    catch (error) { throw new Error('STORE_CONFIG_JSON ไม่ถูกต้อง: ' + error.message); }
  }
  const fallbackKeys = [
    'STORE_ID','SHOP_NAME','PROVIDER','LIFF_ID','LIFF_URL','API_URL',
    'GITHUB_OWNER','GITHUB_REPO','GITHUB_BRANCH','GITHUB_IMAGE_DIR',
    'LINE_LOGIN_CHANNEL_ID','LINE_MESSAGING_CHANNEL_ID'
  ];
  fallbackKeys.forEach(function(key) {
    if (config[key] === undefined || config[key] === null || config[key] === '') {
      const value = props.getProperty(key);
      if (value !== null && value !== '') config[key] = value;
    }
  });
  if (!config.PAYMENT) config.PAYMENT = {};
  if (!config.PICKUP) config.PICKUP = {};
  return config;
}

const CONFIG = {
  get LIFF_ID() { return String(getStoreConfig_().LIFF_ID || ''); },
  SHEETS: {
    MENU: 'Menu', OPTIONS: 'Menu_Options', DAILY_DISH_LIBRARY: 'Daily_Dish_Library',
    ORDERS: 'Orders', DETAILS: 'Order_Detail', CUSTOMERS: 'Customers', SLOTS: 'Pickup_Slots'
  }
};

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getInitialData') return jsonResponse_(getInitialData());
  if (e && e.parameter && e.parameter.action === 'getOrders') { requireAdminAuth_(e.parameter.adminToken); return jsonResponse_(getOrders()); }
  if (e && e.parameter && e.parameter.action === 'getMyOrders') return jsonResponse_(getMyOrders(e.parameter.userId));
  if (e && e.parameter && e.parameter.action === 'getDailyDishLibrary') { requireAdminAuth_(e.parameter.adminToken); return jsonResponse_(getDailyDishLibrary()); }
  if (e && e.parameter && e.parameter.action === 'getMenuSystemData') { requireAdminAuth_(e.parameter.adminToken); return jsonResponse_(getMenuSystemData()); }
  if (e && e.parameter && e.parameter.action === 'getAdminSettings') { requireAdminAuth_(e.parameter.adminToken); return jsonResponse_(getAdminSettings()); }
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle(String(getStoreConfig_().SHOP_NAME || 'FOOD ORDER SYSTEM')).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* The complete production function bodies are retained in the Core package.
 * This source file is the Store-aware entry/config layer; production body
 * synchronization is tracked by golden-master/reference-production/SOURCE.json.
 * See golden-master/backend/README.md for the exact source baseline and
 * synchronization procedure.
 */
