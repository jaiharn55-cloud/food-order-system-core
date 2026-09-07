/**
 * FOOD ORDER SYSTEM - Store Provisioner V2
 *
 * This file is designed to be copied into the STORE BUILDER
 * Google Apps Script project.
 *
 * Required Script Property:
 *   STORE_GITHUB_TOKEN
 *
 * The token should be a fine-grained PAT with the minimum access
 * needed to:
 *   - run the Core repository workflow
 *   - write contents to target store repositories
 * Never hard-code the token in source code.
 */

const STORE_BUILDER_V2 = {
  CORE_REPO: 'jaiharn55-cloud/food-order-system-core',
  WORKFLOW_FILE: 'provision-store.yml',
  BRANCH: 'main',
  REGISTRY_SHEET: 'STORE_REGISTRY'
};

function provisionStorePackageV2(storeId) {
  if (!storeId) {
    throw new Error('กรุณาระบุ Store_ID');
  }

  const store = getBuilderRegistryStoreV2_(storeId);
  if (!store) {
    throw new Error('ไม่พบ Store_ID: ' + storeId);
  }

  const token = String(
    PropertiesService
      .getScriptProperties()
      .getProperty('STORE_GITHUB_TOKEN') || ''
  ).trim();

  if (!token) {
    throw new Error(
      'ยังไม่มี STORE_GITHUB_TOKEN ใน Script Properties'
    );
  }

  const targetRepo = String(
    store.GitHub_Repo || ('food-order-' + String(store.Store_Slug || '').trim())
  ).trim();

  if (!targetRepo) {
    throw new Error('ไม่พบ GitHub repository ของ Store นี้');
  }

  const inputs = {
    target_repo: targetRepo.indexOf('/') >= 0
      ? targetRepo
      : 'jaiharn55-cloud/' + targetRepo,
    store_id: String(store.Store_ID || ''),
    shop_name: String(store.Store_Name || ''),
    provider: String(store.Provider || 'Wichai'),
    liff_id: String(store.LIFF_ID || ''),
    liff_url: String(store.LIFF_URL || ''),
    api_url: String(store.Web_App_URL || store.API_URL || ''),
    line_login_channel_id: String(store.LINE_Login_Channel_ID || ''),
    line_messaging_channel_id: String(store.LINE_Messaging_Channel_ID || ''),
    cash_active: String(store.Cash_Active || false).toLowerCase(),
    promptpay_active: String(store.PromptPay_Active || false).toLowerCase(),
    government_active: String(store.Government_Active || false).toLowerCase(),
    pickup_start: String(store.Pickup_Start || '08:00'),
    pickup_end: String(store.Pickup_End || '17:00'),
    pickup_interval: String(store.Pickup_Interval_Min || '20'),
    pickup_capacity: String(store.Pickup_Capacity || '10')
  };

  const url =
    'https://api.github.com/repos/' +
    STORE_BUILDER_V2.CORE_REPO +
    '/actions/workflows/' +
    STORE_BUILDER_V2.WORKFLOW_FILE +
    '/dispatches';

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2026-03-10'
    },
    payload: JSON.stringify({
      ref: STORE_BUILDER_V2.BRANCH,
      inputs: inputs
    }),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const body = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(
      'GitHub workflow dispatch ไม่สำเร็จ (' +
      status + '): ' + body
    );
  }

  updateBuilderRegistryV2_(storeId, {
    Status: 'PACKAGE_BUILD_STARTED',
    GitHub_Repo: inputs.target_repo,
    Updated_At: new Date()
  });

  appendBuilderLogV2_(
    storeId,
    'PACKAGE_BUILD',
    'STARTED',
    'ส่งคำสั่ง Provision Store Package ไป GitHub Actions แล้ว'
  );

  return {
    success: true,
    storeId: storeId,
    targetRepo: inputs.target_repo,
    workflow: STORE_BUILDER_V2.WORKFLOW_FILE,
    message: 'Provision workflow started'
  };
}

function getBuilderRegistryStoreV2_(storeId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STORE_BUILDER_V2.REGISTRY_SHEET);
  if (!sheet) throw new Error('ไม่พบ STORE_REGISTRY');

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;

  const headers = values[0].map(String);
  const idCol = headers.indexOf('Store_ID');
  if (idCol < 0) throw new Error('STORE_REGISTRY ไม่มีคอลัมน์ Store_ID');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol] || '').trim() === String(storeId).trim()) {
      const obj = {};
      headers.forEach(function(header, c) {
        obj[header] = values[r][c];
      });
      return obj;
    }
  }

  return null;
}

function updateBuilderRegistryV2_(storeId, updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STORE_BUILDER_V2.REGISTRY_SHEET);
  if (!sheet) throw new Error('ไม่พบ STORE_REGISTRY');

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const idCol = headers.indexOf('Store_ID');
  if (idCol < 0) throw new Error('STORE_REGISTRY ไม่มีคอลัมน์ Store_ID');

  let rowNumber = -1;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol] || '').trim() === String(storeId).trim()) {
      rowNumber = r + 1;
      break;
    }
  }

  if (rowNumber < 0) throw new Error('ไม่พบ Store_ID: ' + storeId);

  Object.keys(updates).forEach(function(key) {
    const col = headers.indexOf(key);
    if (col >= 0) sheet.getRange(rowNumber, col + 1).setValue(updates[key]);
  });
}

function appendBuilderLogV2_(storeId, step, status, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PROVISION_LOG');
  if (!sheet) return;
  sheet.appendRow([new Date(), storeId, step, status, message]);
}
