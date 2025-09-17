const HEADERS = [
  'Sipariş',
  'Tarih',
  'Müşteri',
  'E-posta',
  'Telefon',
  'Ülke',
  'Ürün',
  'Toplam',
  'Kargo',
  'Adres',
  'Durum',
];

const SHEET_NAME = 'COD Orders';

/**
 * Serves the COD Orders dashboard as a web application.
 *
 * @return {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.headers = HEADERS;

  const output = template.evaluate().setTitle('COD Orders Dashboard');

  output.setContentSecurityPolicy(
    [
      "default-src 'self' https://www.gstatic.com https://apis.google.com",
      "script-src 'self' 'unsafe-inline' https://www.gstatic.com https://apis.google.com",
      "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
      "img-src 'self' data:",
      "font-src 'self' https://fonts.gstatic.com",
      "connect-src 'self' https://script.google.com",
      "frame-src 'self'",
      "object-src 'none'",
    ].join('; '),
  );

  return output.setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Adds the COD Orders custom menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('COD Orders')
    .addItem('Refresh COD Orders', 'populateOrdersSheet')
    .addItem('Open Filters', 'openFiltersSidebar')
    .addToUi();
}

/**
 * Fetches COD orders and writes them into the spreadsheet.
 *
 * @return {number} Number of orders written to the sheet.
 */
function populateOrdersSheet() {
  const ui = SpreadsheetApp.getUi();

  try {
    const sheet = getOrCreateOrdersSheet_();
    const orders = fetchCodOrders_();

    removeFilterIfExists_(sheet);

    sheet.clearContents();
    sheet.clearFormats();

    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);

    if (orders.length) {
      const values = orders.map((order) => HEADERS.map((header) => order[header] || ''));
      sheet.getRange(2, 1, values.length, HEADERS.length).setValues(values);
      sheet.autoResizeColumns(1, HEADERS.length);
    }

    SpreadsheetApp.getActive().toast(
      `Toplam ${orders.length} sipariş yüklendi.`,
      'COD Orders',
      5,
    );

    return orders.length;
  } catch (error) {
    console.error(error);
    ui.alert(
      'Siparişler alınırken bir hata oluştu. Lütfen alan adı, erişim anahtarı ve API sürümü ayarlarınızı kontrol edin.',
    );
    return 0;
  }
}

/**
 * Opens the sidebar that contains per-column filters.
 */
function openFiltersSidebar() {
  const template = HtmlService.createTemplateFromFile('Filters');
  template.headers = HEADERS;

  const html = template.evaluate().setTitle('COD Orders Filtreleri');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Applies spreadsheet filters based on the provided header -> value map.
 *
 * @param {Object<string, string>} filters
 * @return {number} The number of columns that received a filter.
 */
function applyFilters(filters) {
  const sanitizedFilters = filters || {};
  const sheet = getOrCreateOrdersSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    clearFilters();
    return 0;
  }

  removeFilterIfExists_(sheet);
  const filter = sheet.getRange(1, 1, lastRow, HEADERS.length).createFilter();

  const applied = HEADERS.filter((header) => {
    const value = (sanitizedFilters[header] || '').toString().trim();
    if (!value) {
      return false;
    }

    const columnIndex = HEADERS.indexOf(header) + 1;
    const criteria = SpreadsheetApp.newFilterCriteria().whenTextContains(value).build();
    filter.setColumnFilterCriteria(columnIndex, criteria);
    return true;
  }).length;

  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast(
    applied
      ? `Filtre uygulandı (${applied} sütun).`
      : 'Filtre uygulanmadı; tüm satırlar gösteriliyor.',
    'COD Orders',
    5,
  );

  return applied;
}

/**
 * Removes any active filters from the COD Orders sheet.
 */
function clearFilters() {
  const sheet = getOrCreateOrdersSheet_();
  removeFilterIfExists_(sheet);
  SpreadsheetApp.getActive().toast('Filtreler kaldırıldı.', 'COD Orders', 5);
}

/**
 * Returns COD-tagged orders for rendering in the web UI.
 *
 * @return {{orders: Array<Object<string, string>>, errorMessage: string}}
 */
function getCodOrders() {
  try {
    const orders = fetchCodOrders_();
    return {
      orders,
      errorMessage: orders.length ? '' : 'COD etiketli sipariş bulunamadı.',
    };
  } catch (error) {
    console.error(error);
    return {
      orders: [],
      errorMessage:
        'Siparişler alınırken bir hata oluştu. Lütfen alan adınızı, erişim anahtarınızı ve API sürümünüzü kontrol edin.',
    };
  }
}

/**
 * Fetches COD-tagged orders from Shopify using Admin API credentials stored as script properties.
 *
 * Script properties expected:
 *   - SHOP_DOMAIN: örn. my-store.myshopify.com
 *   - SHOP_ACCESS_TOKEN: Shopify Admin API access token
 *   - (opsiyonel) SHOP_API_VERSION: örn. 2024-01
 *
 * @return {Array<Object<string, string>>}
 */
function fetchCodOrders_() {
  const props = PropertiesService.getScriptProperties();
  const shopDomain = props.getProperty('SHOP_DOMAIN');
  const accessToken = props.getProperty('SHOP_ACCESS_TOKEN');
  const apiVersion = props.getProperty('SHOP_API_VERSION') || '2024-01';

  if (!shopDomain || !accessToken) {
    console.warn('SHOP_DOMAIN veya SHOP_ACCESS_TOKEN tanımlı değil. Örnek veriler döndürülüyor.');
    return getSampleOrders_();
  }

  const endpoint =
    `https://${shopDomain}/admin/api/${apiVersion}/orders.json` +
    '?status=any' +
    '&tagged_with=COD' +
    '&fields=name,created_at,email,phone,total_price,currency,shipping_address,' +
    'customer,line_items,total_shipping_price_set,financial_status,fulfillment_status,tags';

  const response = UrlFetchApp.fetch(endpoint, {
    method: 'get',
    headers: {
      'X-Shopify-Access-Token': accessToken,
      'Content-Type': 'application/json',
    },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      `Shopify API hatası (${response.getResponseCode()}): ${response.getContentText()}`,
    );
  }

  const data = JSON.parse(response.getContentText());
  const codOrders = (data.orders || []).filter((order) => hasCodTag_(order.tags));

  return codOrders.map((order) => {
    const shipping = order.shipping_address || {};
    const customer = order.customer || {};
    const shippingMoney = order.total_shipping_price_set?.shop_money;

    return {
      Sipariş: order.name || '',
      Tarih: formatDate_(order.created_at),
      Müşteri: buildFullName_(customer.first_name, customer.last_name),
      'E-posta': order.email || '',
      Telefon:
        order.phone ||
        customer.phone ||
        customer.default_address?.phone ||
        shipping.phone ||
        '',
      Ülke: shipping.country || '',
      Ürün: (order.line_items || []).map((item) => item.name).join(', '),
      Toplam: formatMoney_(order.total_price, order.currency),
      Kargo: shippingMoney
        ? formatMoney_(shippingMoney.amount, shippingMoney.currency_code)
        : '',
      Adres: formatAddress_(shipping),
      Durum: order.fulfillment_status || order.financial_status || '',
    };
  });
}

/**
 * Returns true if the order tags include COD.
 *
 * @param {string} tags
 * @return {boolean}
 */
function hasCodTag_(tags) {
  if (!tags) {
    return false;
  }
  return tags
    .split(',')
    .map((tag) => tag.trim().toUpperCase())
    .includes('COD');
}

/**
 * Formats ISO date strings into the script time zone.
 *
 * @param {string} isoString
 * @return {string}
 */
function formatDate_(isoString) {
  if (!isoString) {
    return '';
  }
  const timeZone = Session.getScriptTimeZone() || 'Europe/Istanbul';
  return Utilities.formatDate(new Date(isoString), timeZone, 'yyyy-MM-dd HH:mm');
}

/**
 * Combines a first and last name.
 *
 * @param {string} firstName
 * @param {string} lastName
 * @return {string}
 */
function buildFullName_(firstName, lastName) {
  return [firstName, lastName].filter(Boolean).join(' ');
}

/**
 * Formats amount and currency values.
 *
 * @param {string|number} amount
 * @param {string} currency
 * @return {string}
 */
function formatMoney_(amount, currency) {
  if (!amount || !currency) {
    return '';
  }
  return `${amount} ${currency}`;
}

/**
 * Formats a shipping address object into a single string.
 *
 * @param {Object} address
 * @return {string}
 */
function formatAddress_(address) {
  if (!address) {
    return '';
  }

  const parts = [
    address.address1,
    address.address2,
    address.city,
    address.province,
    address.zip,
    address.country,
  ];

  return parts.filter(Boolean).join(', ');
}

/**
 * Provides sample orders when Shopify credentials are not configured.
 *
 * @return {Array<Object<string, string>>}
 */
function getSampleOrders_() {
  return [
    {
      Sipariş: 'ORD-1001',
      Tarih: '2024-01-12 09:30',
      Müşteri: 'Ayşe Yılmaz',
      'E-posta': 'ayse@example.com',
      Telefon: '+90 555 111 2233',
      Ülke: 'Türkiye',
      Ürün: 'Ürün A',
      Toplam: '1299.00 TRY',
      Kargo: '49.00 TRY',
      Adres: 'İstiklal Cd. No:10, Beyoğlu, İstanbul, 34000, Türkiye',
      Durum: 'Beklemede',
    },
    {
      Sipariş: 'ORD-1002',
      Tarih: '2024-01-13 14:05',
      Müşteri: 'Mehmet Demir',
      'E-posta': 'mehmet@example.com',
      Telefon: '+90 555 444 5566',
      Ülke: 'Türkiye',
      Ürün: 'Ürün B',
      Toplam: '899.00 TRY',
      Kargo: '39.00 TRY',
      Adres: 'Bağdat Cd. No:25, Kadıköy, İstanbul, 34728, Türkiye',
      Durum: 'Kargolandı',
    },
  ];
}

/**
 * Returns the COD Orders sheet, creating it when necessary.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateOrdersSheet_() {
  const spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  ensureHeaderRow_(sheet);
  return sheet;
}

/**
 * Removes an existing filter from the given sheet without user notifications.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function removeFilterIfExists_(sheet) {
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
}

/**
 * Ensures the first row of the sheet matches the HEADERS array.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureHeaderRow_(sheet) {
  const range = sheet.getRange(1, 1, 1, HEADERS.length);
  const current = range.getValues()[0];

  const shouldRewrite = HEADERS.some((header, index) => current[index] !== header);
  if (shouldRewrite) {
    range.setValues([HEADERS]);
    range.setFontWeight('bold');
  }
  sheet.setFrozenRows(1);
}
