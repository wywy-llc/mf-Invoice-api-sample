/**
 * 初期設定
 * ・トリガー作成
 * ・シート作成
 */
function initialize() {
  const initTriggers = () => {
    // トリガー作成
    const functionNames = ['onOpen'];
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      const fname = trigger.getHandlerFunction();
      if (functionNames.includes(fname)) {
        ScriptApp.deleteTrigger(trigger);
        switch (fname) {
          case 'onOpen':
            ScriptApp.newTrigger(fname).forSpreadsheet(spreadsheet).onOpen().create();
        }
      }
    }
  }
  const headers = {
    office: [
      'id',
      'name',
      'zip',
      'prefecture',
      'address1',
      'address2',
      'tel',
      'fax',
      'office_type',
      'office_code',
      'registration_code',
      'created_at',
      'updated_at',
    ],
    items: [
      'id',
      'code',
      'name',
      'name_kana',
      'name_suffix',
      'memo',
      'created_at',
      'updated_at',
      'departments',
    ],
    partners: [
      'id',
      'code',
      'name',
      'name_kana',
      'name_suffix',
      'memo',
      'created_at',
      'updated_at',
      'departments',
    ],
    items: [
      'id',
      'name',
      'code',
      'detail',
      'unit',
      'price',
      'quantity',
      'is_deduct_withholding_tax',
      'excise',
      'created_at',
      'updated_at'
    ],
    billings: ['id',
      'pdf_url',
      'operator_id',
      'department_id',
      'member_id',
      'member_name',
      'partner_id',
      'partner_name',
      'office_id',
      'office_name',
      'office_detail',
      'title',
      'memo',
      'payment_condition',
      'billing_date',
      'due_date',
      'sales_date',
      'billing_number',
      'note',
      'document_name',
      'payment_status',
      'email_status',
      'posting_status',
      'created_at',
      'updated_at',
      'is_downloaded',
      'is_locked',
      'deduct_price',
      'tag_names',
      'items',
      'excise_price',
      'excise_price_of_untaxable',
      'excise_price_of_non_taxable',
      'excise_price_of_tax_exemption',
      'excise_price_of_five_percent',
      'excise_price_of_eight_percent',
      'excise_price_of_eight_percent_as_reduced_tax_rate',
      'excise_price_of_ten_percent',
      'subtotal_price',
      'subtotal_of_untaxable_excise',
      'subtotal_of_non_taxable_excise',
      'subtotal_of_tax_exemption_excise',
      'subtotal_of_five_percent_excise',
      'subtotal_of_eight_percent_excise',
      'subtotal_of_eight_percent_as_reduced_tax_rate_excise',
      'subtotal_of_ten_percent_excise',
      'subtotal_with_tax_of_untaxable_excise',
      'subtotal_with_tax_of_non_taxable_excise',
      'subtotal_with_tax_of_five_percent_excise',
      'subtotal_with_tax_of_tax_exemption_excise',
      'subtotal_with_tax_of_eight_percent_excise',
      'subtotal_with_tax_of_eight_percent_as_reduced_tax_rate_excise',
      'subtotal_with_tax_of_ten_percent_excise',
      'total_price',
      'registration_code',
      'use_invoice_template'],
    quotes: [
      'id',
      'pdf_url',
      'operator_id',
      'department_id',
      'member_id',
      'member_name',
      'partner_id',
      'partner_name',
      'partner_detail',
      'office_id',
      'office_name',
      'office_detail',
      'title',
      'memo',
      'quote_date',
      'quote_number',
      'note',
      'expired_date',
      'document_name',
      'order_status',
      'transmit_status',
      'posting_status',
      'created_at',
      'updated_at',
      'is_downloaded',
      'is_locked',
      'deduct_price',
      'tag_names',
      'items',
      'excise_price',
      'excise_price_of_untaxable',
      'excise_price_of_non_taxable',
      'excise_price_of_tax_exemption',
      'excise_price_of_five_percent',
      'excise_price_of_eight_percent',
      'excise_price_of_eight_percent_as_reduced_tax_rate',
      'excise_price_of_ten_percent',
      'subtotal_price',
      'subtotal_of_untaxable_excise',
      'subtotal_of_non_taxable_excise',
      'subtotal_of_tax_exemption_excise',
      'subtotal_of_five_percent_excise',
      'subtotal_of_eight_percent_excise',
      'subtotal_of_eight_percent_as_reduced_tax_rate_excise',
      'subtotal_of_ten_percent_excise',
      'total_price'
    ]
  };
  const initSheets = () => {
    // シート作成
    const client = getMfClient_();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    for (const attr in client) {
      let sheet = spreadsheet.getSheetByName(attr);
      if (sheet) {
        spreadsheet.deleteSheet(sheet);
      }
      sheet = spreadsheet.insertSheet(attr);
      const headerNames = headers[attr];
      if (!headerNames) {
        continue;
      }
      const range = sheet.getRange(1, 1, 1, headerNames.length);
      range.setBackground("#bdbdbd");
      range.setValues([headerNames]);
      if (headerNames.length < sheet.getMaxColumns()) {
        sheet.deleteColumns(headerNames.length + 1, sheet.getMaxColumns() - headerNames.length);
      }
    }
  }
  initTriggers();
  initSheets();
}

/**
 * リダイレクトURIを取得します。
 * @return {string} リダイレクトURI
 */
function getRedirectUri() {
  console.log(MfInvoiceApi.getRedirectUri());
}

/**
 * シンプルトリガー
 * スプレッドシート、をユーザーが開く時に呼び出される関数です。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui
    .createMenu('MF請求書API連携')
    .addItem('認証処理を開始する', 'showMfApiAuthDialog');
  menu.addToUi();
}

/**
 * MF請求書API認証ダイアログを表示します。
 */
function showMfApiAuthDialog() {
  const scriptProps = PropertiesService.getScriptProperties();
  const clientId = scriptProps.getProperty('CLIENT_ID');
  if (!clientId) {
    throw new Error('CLIENT_IDが設定されていません。');
  }
  const clientSecret = scriptProps.getProperty('CLIENT_SECRET');
  if (!clientSecret) {
    throw new Error('CLIENT_SECRETが設定されていません。');
  }
  MfInvoiceApi.showMfApiAuthDialog(clientId, clientSecret);
}

/**
 * MF認証のコールバック関数です。
 * @param request
 */
function mfCallback(request) {
  const scriptProps = PropertiesService.getScriptProperties();
  const clientId = scriptProps.getProperty('CLIENT_ID');
  if (!clientId) {
    throw new Error('CLIENT_IDが設定されていません。');
  }
  const clientSecret = scriptProps.getProperty('CLIENT_SECRET');
  if (!clientSecret) {
    throw new Error('CLIENT_SECRETが設定されていません。');
  }
  return MfInvoiceApi.mfCallback(request, clientId, clientSecret);
}

/**
 * MF請求書APIクライアントを生成します。
 * @returns {MfClient}
 */
function getMfClient_() {
  const scriptProps = PropertiesService.getScriptProperties();
  const clientId = scriptProps.getProperty('CLIENT_ID');
  const clientSecret = scriptProps.getProperty('CLIENT_SECRET');
  return MfInvoiceApi.createClient(clientId, clientSecret);
}

/**
 * 全てのAPIをテスト実行します。
 */
function testAllApi() {
  //  事業者情報の取得
  getMyOffice();

  // 取引先の作成
  createNewPartner();

  // 取引先一覧の取得
  getPartners();

  // 取引先の取得
  getPartner();

  // 取引先の更新
  updatePartner();

  // 品目の作成
  createNewItem();

  // 品目一覧の取得
  getItems();

  // 品目の取得
  getItem();

  // インボイス制度に対応した形式の請求書の作成
  createNewInvoiceTemplateBilling();

  // 請求書一覧の取得
  getBillings();

  // 請求書の取得
  getBilling();

  // 見積書の作成
  createNewQuote();

  // 見積書一覧の取得
  getQuotes();

  // 見積書の取得
  getQuote();
}

//== Office(事業所) ==

/**
 * 事業者情報の取得
 */
function getMyOffice() {
  // API実行： 事業者情報の取得
  const office = getMfClient_().office.getMyOffice();
  console.log(office);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("office");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in office) {
    row.push(office[attr]);
  }
  sheet.appendRow(row);
}

//== Partner ==

/**
 * 取引先の作成
 */
function createNewPartner() {
  // 取引先
  const partner = {
    code: new Date().getTime().toString(),
    name: '取引先名',
    name_kana: '取引先名(カナ)',
    name_suffix: '御中',
    memo: 'メモ',
    departments: [
      {
        zip: '770-0053',
        tel: '03-1234-5678',
        prefecture: '徳島県',
        address1: '徳島市 南島田町２丁目５８ー３',
        address2: 'オレス南島田Ｂ棟',
        person_name: '担当者_氏名',
        person_title: '担当者_役職',
        person_dept: '担当者_部門',
        office_member_name: '自社担当者_氏名',
        email: 'sample@example.com',
        cc_emails: 'sample_cc_01@example.com,sample_cc_02@example.com'
      }
    ]
  }

  // API実行： 取引先の作成
  const createdPartner = getMfClient_().partners.createNew(partner);

  console.log(createdPartner);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("partners");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in createdPartner) {
    row.push(createdPartner[attr]);
  }
  sheet.appendRow(row);
}

/**
 * 取引先一覧の取得
 */
function getPartners() {
  // API実行： 取引先一覧の取得
  const partners = getMfClient_().partners.getPartners();
  console.log(partners.data);
  console.log(partners.data[0].departments);
  console.log('件数： ' + partners.pagination.total_count);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("partners");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  for (const partner of partners.data) {
    const row = [];
    for (const attr in partner) {
      if (attr === 'departments') {
        row.push(JSON.stringify(partner[attr]));
        continue;
      }
      row.push(partner[attr]);
    }
    sheet.appendRow(row);
  }
}

/**
 * 取引先の取得
 */
function getPartner() {
  // 取引先IDの準備
  const partners = getMfClient_().partners.getPartners();
  const partnerId = partners.data[0].id;

  // API実行： 取引先の取得
  const partner = getMfClient_().partners.getPartner(partnerId);
  console.log(partner);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("partners");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in partner) {
    if(attr === 'departments'){
      row.push(JSON.stringify(partner[attr]));
      continue;
    }
    row.push(partner[attr]);
  }
  sheet.appendRow(row);
}

/**
 * 取引先の更新
 */
function updatePartner(){
  // 取引先の準備
  const partners = getMfClient_().partners.getPartners();
  const partner = partners.data[0];

  // 取引先オブジェクト
  const partnerReqBody = {
    code: partner.code + '_更新',
    name: partner.name + '_更新',
    name_kana: partner.name + '_更新', 
    name_suffix: '様',
    memo: partner.memo + '_更新',
  }

  // API実行： 取引先の更新
  const updatedPartner = getMfClient_().partners.updatePartner(partner.id, partnerReqBody);
  console.log(updatedPartner);
}

//== Item(品目) ==

/**
 * 品目の作成
 */
function createNewItem() {
  // 品目
  const newItem = {
    name: '品名',
    code: new Date().getTime().toString(),
    detail: '品目詳細',
    unit: '単位',
    price: 1239,
    quantity: 1,
    is_deduct_withholding_tax: true,
    excise: 'ten_percent',
  }

  // API実行： 品目の作成
  const createdItem = getMfClient_().items.createNew(newItem);
  console.log(createdItem);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("items");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in createdItem) {
    row.push(createdItem[attr]);
  }
  sheet.appendRow(row);
}

/**
 * 品目一覧の取得
 */
function getItems() {
  // API実行： 品目一覧の取得
  const items = getMfClient_().items.getItems();
  console.log(items.data);
  console.log('件数： ' + items.pagination.total_count);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("items");
  for (const item of items.data) {
    const row = [];
    for (const attr in item) {
      row.push(item[attr]);
    }
    sheet.appendRow(row);
  }
}

/**
 * 品目の取得
 */
function getItem() {
  // 品目IDの準備
  const items = getMfClient_().items.getItems();
  const itemId = items.data[0].id;

  // API実行： 品目の取得
  const targetItem = getMfClient_().items.getItem(itemId);
  console.log(targetItem);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("items");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in targetItem) {
    row.push(targetItem[attr]);
  }
  sheet.appendRow(row);
}

//== Billing(請求書) ==

/**
 * インボイス制度に対応した形式の請求書の作成
 */
function createNewInvoiceTemplateBilling() {
  // 部門ID(department.id)の準備
  const partners = getMfClient_().partners.getPartners();
  const partner = partners.data[0];
  const department = partner.departments[0];
  const dateUtil = MfInvoiceApi.getDateUtil(new Date());

  // 商品ID(item.id)の準備
  const items = getMfClient_().items.getItems();
  const item = items.data[0];

  // 先月末
  const endDateLastMonth = dateUtil.getEndDateLastMonth();

  // 本日
  const today = dateUtil.getDateString();

  // 今月末
  const endDateBaseMonth = dateUtil.getEndDateBaseMonth();

  // 請求書
  const billging = {
    department_id: department.id,
    title: '件名',
    memo: 'メモ',
    payment_condition: '振込先',
    billing_date: endDateLastMonth,
    due_date: endDateBaseMonth,
    sales_date: today,
    billing_number: new Date().getTime().toString(),
    note: '備考',
    document_name: '帳票名',
    tag_names: [
      'タグ'
    ],
    items: [{
      item_id: item.id,
      delivery_number: '1234Num567',
      delivery_date: today,
      detail: item.detail,
      unit: item.unit,
      price: item.price,
      quantity: 10,
      is_deduct_withholding_tax: item.isDeductWithholdingTax,
      excise: item.excise
    }]
  }

  // API実行： 請求書の作成
  const createdBillging = getMfClient_().billings.createNew(billging);

  console.log(createdBillging);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("billings");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in createdBillging) {
    if (attr === 'items' || attr === 'tag_names') {
      row.push(JSON.stringify(createdBillging[attr]));
      continue;
    }
    row.push(createdBillging[attr]);
  }
  sheet.appendRow(row);
}

/**
 * 請求書一覧の取得
 */
function getBillings() {

  // 日付操作
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);

  // 先月末
  const from = dateUtil.getEndDateLastMonth();

  // 来月末
  const to = dateUtil.getEndDateNextMonth();

  // 検索キー
  const query = '入金済み';

  // API実行： 請求書一覧の取得
  const billings = getMfClient_().billings.getBillings(from, to, query);
  console.log(billings.data[0]);
  console.log('件数: ' + billings.pagination.total_count);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("billings");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  for (const billing of billings.data) {
    const row = [];
    for (const attr in billing) {
      if (attr === 'items' || attr === 'tag_names') {
        row.push(JSON.stringify(billing[attr]));
        continue;
      }
      row.push(billing[attr]);
    }
    sheet.appendRow(row);
  }
}

/**
 * 請求書の取得
 */
function getBilling() {
  // 請求書IDの取得
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '入金済み';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const billingId = billings.data[0].id

  // API実行： 請求書の取得
  const billing = getMfClient_().billings.getBilling(billingId);
  console.log(billing);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("billings");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in billing) {
    if (attr === 'items' || attr === 'tag_names') {
      row.push(JSON.stringify(billing[attr]));
      continue;
    }
    row.push(billing[attr]);
  }
  sheet.appendRow(row);
}

//== Quote(見積書) ==

/**
 * 見積書の作成
 */
function createNewQuote() {
  // 部門ID(department.id)の取得
  const partners = getMfClient_().partners.getPartners();
  const partner = partners.data[0];
  const department = partner.departments[0];

  // 商品ID(item.id)の取得
  const items = getMfClient_().items.getItems();
  const item = items.data[0];

  // 日付操作
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);

  // 先月末
  const endDateLastMonth = dateUtil.getEndDateLastMonth();

  // 今月末
  const endDateBaseMonth = dateUtil.getEndDateBaseMonth();

  // 見積書の生成
  const quote = {
    department_id: department.id,
    quote_number: new Date().getTime().toString(),
    title: '件名',
    memo: 'メモ',
    quote_date: endDateLastMonth,
    expired_date: endDateBaseMonth,
    note: '備考',
    tag_names: [
      'タグ'
    ],
    items: [
      {
        item_id: item.id,
        detail: item.detail,
        unit: item.unit,
        price: item.price,
        quantity: 10,
        is_deduct_withholding_tax: item.isDeductWithholdingTax,
        excise: item.excise
      }
    ],
    document_name: '帳票名'
  }

  // API実行： 見積書の登録
  const createdQuote = getMfClient_().quotes.createNew(quote);

  console.log(createdQuote);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quotes");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in createdQuote) {
    if (attr === 'items' || attr === 'tag_names') {
      row.push(JSON.stringify(createdQuote[attr]));
      continue;
    }
    row.push(createdQuote[attr]);
  }
  sheet.appendRow(row);
}

/**
 * 見積書一覧の取得
 */
function getQuotes() {
  // 日付操作
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);

  // 先月末
  const from = dateUtil.getEndDateLastMonth();

  // 来月末
  const to = dateUtil.getEndDateNextMonth();

  // 検索キー
  const query = '未設定';

  // API実行： 見積一覧の取得
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);

  console.log(quotes.data[0]);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quotes");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  for (const quote of quotes.data) {
    const row = [];
    for (const attr in quote) {
      if (attr === 'items' || attr === 'tag_names') {
        row.push(JSON.stringify(quote[attr]));
        continue;
      }
      row.push(quote[attr]);
    }
    sheet.appendRow(row);
  }
}

/**
 * 見積書の取得
 */
function getQuote() {
  // 見積書IDの準備
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '未設定';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const quoteId = quotes.data[0].id;

  // API実行： 見積書IDの取得
  const quote = getMfClient_().quotes.getQuote(quoteId);
  console.log(quote);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quotes");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in quote) {
    if (attr === 'items' || attr === 'tag_names') {
      row.push(JSON.stringify(quote[attr]));
      continue;
    }
    row.push(quote[attr]);
  }
  sheet.appendRow(row);
}
