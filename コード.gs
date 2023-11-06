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
  const schemas = {
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
      'use_invoice_template',
      'config'
    ],
    billingItems: [
      'id',
      'name',
      'code',
      'detail',
      'unit',
      'price',
      'quantity',
      'is_deduct_withholding_tax',
      'excise',
      'delivery_date',
      'delivery_number',
      'created_at',
      'updated_at',
      'billing_id'
    ],
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
    ],
    quoteItems: [
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
      'updated_at',
      'quote_id'
    ]
  };
  const initSheets = () => {
    // シート作成
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

    for (const schema of Object.keys(schemas)) {
      let sheet = spreadsheet.getSheetByName(schema);
      if (sheet) {
        // シートの削除（初期化のため）
        spreadsheet.deleteSheet(sheet);
      }
      // シートの挿入
      sheet = spreadsheet.insertSheet(schema);
      const attrs = schemas[schema];
      const range = sheet.getRange(1, 1, 1, attrs.length);
      range.setBackground("#bdbdbd");
      range.setValues([attrs]);

      // 不要な列を削除する
      if (attrs.length < sheet.getMaxColumns()) {
        sheet.deleteColumns(attrs.length + 1, sheet.getMaxColumns() - attrs.length);
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
 * MF認証情報を取得します。
 */
function getMfCredentials_() {
  const scriptProps = PropertiesService.getScriptProperties();
  const clientId = scriptProps.getProperty('CLIENT_ID');
  if (!clientId) {
    throw new Error('CLIENT_IDが設定されていません。');
  }
  const clientSecret = scriptProps.getProperty('CLIENT_SECRET');
  if (!clientSecret) {
    throw new Error('CLIENT_SECRETが設定されていません。');
  }
  const credentials = {
    clientId: clientId,
    clientSecret: clientSecret,
  };
  return credentials;
}

/**
 * MF請求書API認証ダイアログを表示します。
 */
function showMfApiAuthDialog() {
  const credentials = getMfCredentials_();
  MfInvoiceApi.showMfApiAuthDialog(
    credentials.clientId,
    credentials.clientSecret
  );
}

/**
 * MF認証のコールバック関数です。
 * @param request
 */
function mfCallback(request) {
  const credentials = getMfCredentials_();
  return MfInvoiceApi.mfCallback(
    request,
    credentials.clientId,
    credentials.clientSecret
  );
}

/**
 * MF請求書APIクライアントを生成します。
 * @returns {MfClient}
 */
function getMfClient_() {
  const credentials = getMfCredentials_();
  return MfInvoiceApi.createClient(
    credentials.clientId,
    credentials.clientSecret
  );
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

  //　品目の更新
  updateItem();

  // インボイス制度に対応した形式の請求書の作成
  createNewInvoiceTemplateBilling();

  // 請求書一覧の取得
  getBillings();

  // 請求書の入金ステータス変更
  updatePaymentStatus();

  // 請求書の更新
  updateBilling();

  // 請求書の取得
  getBilling();

  // 請求書に品目を追加
  attachBillingItem();

  // 請求書に紐付く品目の一覧取得
  getBillingItems();

  // 請求書に紐づく品目の取得
  getBillingItem();

  // 請求書の郵送依頼
  // 本当に郵送依頼されるので、実行後は必ずキャンセルしてください
  // applyToPostBilling();

  // 請求書の郵送キャンセル
  // cancelPostBilling();

  // 請求書の削除
  deleteBilling();

  // 見積書の作成
  createNewQuote();

  // 

  // 見積書一覧の取得
  getQuotes();

  // 見積書の取得
  getQuote();

  // 見積書の受注ステータス変更
  updateOrderStatus();

  // 見積書の郵送依頼
  // 本当に郵送依頼されるので、実行後は必ずキャンセルしてください
  // applyToPostQuote()

  // 見積書の郵送キャンセル
  // cancelPostQuote()

  // 見積書に紐づく品目一覧の取得
  getQuoteItems();

  // 見積書に紐づく品目の取得
  getQuoteItem()

  // 見積書に紐づく品目を削除
  deleteQuoteItem();

  // 見積書の削除
  deleteQuote();

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
    if (attr === 'departments') {
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
function updatePartner() {
  // 取引先の準備
  const partners = getMfClient_().partners.getPartners();
  const partner = partners.data[0];

  // 取引先オブジェクト
  const partnerReqBody = {
    code: partner.code,
    name: '更新_取引先名',
    name_kana: '更新_カナ',
    name_suffix: '様',
    memo: '更新_メモ',
  }

  // API実行： 取引先の更新
  const updatedPartner = getMfClient_().partners.updatePartner(partner.id, partnerReqBody);
  console.log(updatedPartner);
}

/**
 * 取引先の削除
 */
function deletePartner() {
  // 取引先の準備
  const partners = getMfClient_().partners.getPartners();
  const targetPartnerId = partners.data[0].id;

  // API実行： 取引先の削除
  const result = getMfClient_().partners.deletePartner(targetPartnerId);
  console.log(result);

  if (!result) {
    // 削除に失敗した場合は処理しない。
    return;
  }

  // スプレッドシートから削除
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("partners");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  if (sheet.getLastRow() - 1 === 0) {
    // 削除する行が無いので処理を中止する。
    return;
  }
  sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  ).getValues().forEach((row, index) => {
    const rowPosition = index + 2;
    const billingId = row[0];
    if (targetPartnerId === billingId) {
      sheet.deleteRow(rowPosition);
      // 削除に成功したら処理をやめる
      return;
    }
  });
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
 * 品目の更新
 */
function updateItem() {
  // 品目ID
  const items = getMfClient_().items.getItems();
  const item = items.data[0];

  // 品目
  const itemReqBody = {
    name: '品目_更新',
    code: item.code,
    detail: '品目詳細_更新',
    unit: '単位_更新',
    price: 1240,
    quantity: 1,
    is_deduct_withholding_tax: false,
    excise: 'ten_percent',
  }

  // API実行： 品目の更新
  const updatedItem = getMfClient_().items.updateItem(item.id, itemReqBody);
  console.log(updatedItem);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("items");
  const row = [];
  for (const attr in updatedItem) {
    row.push(updatedItem[attr]);
  }
  sheet.appendRow(row);
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

/**
 * 品目の削除
 */
function deleteItem() {
  // 品目IDの準備
  const items = getMfClient_().items.getItems();
  const targetItemId = items.data[0].id;

  // API実行： 品目の削除
  const result = getMfClient_().items.deleteItem(targetItemId);
  console.log(result);

  if (!result) {
    // 削除に失敗した場合は処理しない。
    return;
  }

  // スプレッドシートから削除
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("items");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  if (sheet.getLastRow() - 1 === 0) {
    // 削除する行が無いので処理を中止する。
    return;
  }
  sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  ).getValues().forEach((row, index) => {
    const rowPosition = index + 2;
    const itemId = row[0];
    if (targetItemId === itemId) {
      sheet.deleteRow(rowPosition);
      // 削除に成功したら処理をやめる
      return;
    }
  });
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
  const query = '';

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
 * 　請求書の更新
 */
function updateBilling() {

  // 請求書の準備
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const billing = billings.data[0]

  // 請求書
  const billgingReqBody = {
    department_id: billing.department_id,
    title: '更新_件名',
    memo: ' 更新_メモ',
    payment_condition: '更新_振込先',
    billing_date: dateUtil.getEndDateLastMonth(),
    due_date: dateUtil.getEndDateBaseMonth(),
    sales_date: dateUtil.getDateString(),
    billing_number: billing.billing_number,
    note: '更新_備考',
    document_name: '更新_帳票名',
    tag_names: [
      '更新_タグ'
    ],
  }

  // API実行： 請求書の更新
  const updatedBilling = getMfClient_().billings.updateBilling(billing.id, billgingReqBody);
  console.log(updatedBilling);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("billings");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in updatedBilling) {
    if (attr === 'items' || attr === 'tag_names') {
      row.push(JSON.stringify(updatedBilling[attr]));
      continue;
    }
    row.push(updatedBilling[attr]);
  }
  sheet.appendRow(row);
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
  const query = '';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const billingId = billings.data[0].id;

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

/**
 * 請求書の入金ステータス変更
 */
function updatePaymentStatus() {
  // 請求書の取得
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const billing = billings.data[0];
  console.log('更新前: ' + billing.payment_status);

  // API実行： 請求書の入金ステータス変更
  const updatedBilling = getMfClient_().billings.updatePaymentStatus(billing.id, MfInvoiceApi.getPaymentStatus('completed'));
  console.log('更新後: ' + updatedBilling.payment_status);

  // スプレッドシートを更新
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("billings");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  if (sheet.getLastRow() - 1 === 0) {
    // 更新する行が無いので処理を中止する。
    return;
  }
  sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  ).getValues().forEach((row, index) => {
    const rowPosition = index + 2;
    const billingId = row[0];
    if (updatedBilling.id === billingId) {
      // 入金ステータスの列番号
      const paymentStatusColumn = 21;

      // 入金ステータスの更新
      sheet.getRange(rowPosition, paymentStatusColumn, 1, 1).setValue(updatedBilling.payment_status);

      // 更新成功したら処理をやめる
      return;
    }
  });
}

/**
 * 請求書に品目を追加
 */
function attachBillingItem() {

  // 商品ID(item.id)の準備
  const items = getMfClient_().items.getItems();
  const item = items.data[0];

  // 請求書の取得
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const targetBilling = billings.data[0];

  // 追加品目
  const itemReqBody = {
    item_id: item.id,
    quantity: 10,
  }

  // API実行： 請求書に品目を追加
  const result = getMfClient_().billings.attachBillingItem(targetBilling.id, itemReqBody);
  console.log(result);
}

/**
 * 請求書の郵送依頼
 */
function applyToPostBilling() {
  // 請求書の取得
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const postBilling = billings.data[0];
  console.log('更新前: ' + postBilling.posting_status);

  // API実行： 請求書の郵送依頼
  const result = getMfClient_().billings.applyToPostBilling(postBilling.id);
  console.log(result);

  // API実行： 請求書の取得
  const billing = getMfClient_().billings.getBilling(postBilling.id);
  console.log('更新後: ' + billing.posting_status);
}

/**
 * 請求書の郵送キャンセル
 */
function cancelPostBilling() {
  // 請求書の取得
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const cancelBilling = billings.data[0];
  console.log('更新前: ' + cancelBilling.posting_status);

  // API実行： 請求書の郵送キャンセル
  const result = getMfClient_().billings.cancelPostBilling(cancelBilling.id);
  console.log(result);

  // API実行： 請求書の取得
  const billing = getMfClient_().billings.getBilling(cancelBilling.id);
  console.log('更新後: ' + billing.posting_status);
}

/**
 * 請求書に紐づく品目一覧の取得
 */
function getBillingItems() {
  // 請求書IDの取得
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const billingId = billings.data[0].id;

  // API実行： 請求書一覧の取得
  const billingItemRes = getMfClient_().billings.getBillingItems(billingId);
  console.log(billingItemRes);
  console.log('件数: ' + billings.pagination.total_count);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("billingItems");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  for (const billingItem of billingItemRes.data) {
    const row = [];
    for (const attr in billingItem) {
      row.push(billingItem[attr]);
    }
    row.push(billingId);
    sheet.appendRow(row);
  }
}

/**
 * 請求書に紐づく品目の取得
 */
function getBillingItem() {
  // 請求書IDの取得
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const billing = billings.data[0];
  const itemId = billing.items[0].id;

  // API実行： 請求書に紐づく品目の取得
  const billingItem = getMfClient_().billings.getBillingItem(billing.id, itemId);
  console.log(billingItem);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("billingItems");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in billingItem) {
    row.push(billingItem[attr]);
  }
  row.push(billing.id);
  sheet.appendRow(row);
}

/**
 * 請求書に紐づく品目の削除
 */
function deleteBillingItem() {
  // 請求書IDの取得
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const billing = billings.data[0];
  const itemId = billing.items[0].id;

  // API実行： 請求書に紐づく品目の削除
  const result = getMfClient_().billings.deleteBillingItem(billing.id, itemId);
  console.log(result);

  // スプレッドシートから削除
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("billingItems");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  if (sheet.getLastRow() - 1 === 0) {
    // 削除する行が無いので処理を中止する。
    return;
  }
  sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  ).getValues().forEach((row, index) => {
    const rowPosition = index + 2;
    const billingItemId = row[0];
    if (itemId === billingItemId) {
      sheet.deleteRow(rowPosition);
      // 削除に成功したら処理をやめる
      return;
    }
  });
}

/**
 * 請求書の削除
 */
function deleteBilling() {
  // 請求書IDの取得
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const billings = getMfClient_().billings.getBillings(from, to, query);
  const targetBillingId = billings.data[0].id;
  console.log('削除対象: ' + targetBillingId)

  // API実行： 請求書の削除
  const result = getMfClient_().billings.deleteBilling(targetBillingId);
  console.log(result);

  if (!result) {
    // 削除に失敗した場合は処理しない。
    return;
  }

  // スプレッドシートから削除
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("billings");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  if (sheet.getLastRow() - 1 === 0) {
    // 削除する行が無いので処理を中止する。
    return;
  }
  sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  ).getValues().forEach((row, index) => {
    const rowPosition = index + 2;
    const billingId = row[0];
    if (targetBillingId === billingId) {
      sheet.deleteRow(rowPosition);
      // 削除に成功したら処理をやめる
      return;
    }
  });
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
  const query = '';

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
 * 見積書に品目を追加
 */
function attachQuoteItem() {
  // 商品ID(item.id)の取得
  const items = getMfClient_().items.getItems();
  const item = items.data[0];

  // 品目
  const quoteItemReqBody =
  {
    item_id: item.id,
    quantity: 10,
  }

  // 見積書IDの準備
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const quoteId = quotes.data[0].id;

  // API実行： 見積書に品目を追加
  const result = getMfClient_().quotes.attachQuoteItem(quoteId, quoteItemReqBody);
  Logger.log(result);

  if(!result){
    throw new Error('見積書への品目追加に失敗しました。');
  }

}

/**
 * 　見積書の更新
 */
function updateQuote() {
  // 見積書の準備
  // 日付操作
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const quote = quotes.data[0];

  // 見積書
  const quoteReqBody = {
    department_id: quote.department_id,
    quote_number: quote.quote_number,
    title: '更新_件名',
    memo: '更新_メモ',
    quote_date: dateUtil.getEndDateLastMonth(),
    expired_date: dateUtil.getDateString(),
    note: '更新_備考',
    tag_names: [
      '更新_タグ'
    ],
  }

  console.log(quoteReqBody);

  // API実行： 見積書の更新
  const updatedQuote = getMfClient_().quotes.updateQuote(quote.id, quoteReqBody);
  console.log(updatedQuote);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quotes");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in updatedQuote) {
    if (attr === 'items' || attr === 'updatedQuote') {
      row.push(JSON.stringify(updatedQuote[attr]));
      continue;
    }
    row.push(updatedQuote[attr]);
  }
  sheet.appendRow(row);
}

/**
 * 見積書の受注ステータス変更
 */
function updateOrderStatus() {
  // 見積書の準備
  // 日付操作
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const quote = quotes.data[0];
  console.log('更新前: ' + quote.order_status);

  // API実行： 見積書の受注ステータス変更
  const updatedQuote = getMfClient_().quotes.updateOrderStatus(quote.id, MfInvoiceApi.getOrderStatus('received'));
  console.log('更新後: ' + updatedQuote.order_status);

  // スプレッドシートを更新
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quotes");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  if (sheet.getLastRow() - 1 === 0) {
    // 更新する行が無いので処理を中止する。
    return;
  }
  sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  ).getValues().forEach((row, index) => {
    const rowPosition = index + 2;
    const quoteId = row[0];
    if (updatedQuote.id === quoteId) {
      // 受注ステータスの列番号
      const orderStatusColumn = 20;

      // 受注ステータスの更新
      sheet.getRange(rowPosition, orderStatusColumn, 1, 1).setValue(updatedQuote.order_status);

      // 更新成功したら処理をやめる
      return;
    }
  });
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
  const query = '';
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

/**
 * 見積書の郵送依頼
 */
function applyToPostQuote() {
  // 見積書の準備
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const postQuote = quotes.data[0];
  console.log('更新前: ' + postQuote.posting_status);

  // API実行： 見積書の郵送依頼
  const result = getMfClient_().quotes.applyToPostQuote(postQuote.id);
  console.log(result);

  // API実行： 見積書の取得
  const updatedQuote = getMfClient_().quotes.getQuote(postQuote.id);
  console.log('更新後: ' + updatedQuote.posting_status);
}

/**
 * 見積書の郵送キャンセル
 */
function cancelPostQuote() {
  // 見積書の準備
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const postQuote = quotes.data[0];
  console.log('更新前: ' + postQuote.posting_status);

  // API実行： 見積書の郵送依頼
  const result = getMfClient_().quotes.cancelPostQuote(postQuote.id);
  console.log(result);

  // API実行： 見積書の取得
  const updatedQuote = getMfClient_().quotes.getQuote(postQuote.id);
  console.log('更新後: ' + updatedQuote.posting_status);
}

/**
 * 見積書に紐づく品目一覧の取得
 */
function getQuoteItems() {
  // 見積書IDの準備
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const quoteId = quotes.data[0].id;

  // API実行： 見積書IDの取得
  const quoteItemRes = getMfClient_().quotes.getQuoteItems(quoteId);
  console.log(quoteItemRes);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quoteItems");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  for (const quoteItem of quoteItemRes.data) {
    const row = [];
    for (const attr in quoteItem) {
      row.push(quoteItem[attr]);
    }
    row.push(quoteId);
    sheet.appendRow(row);
  }
}

/**
 * 見積書に紐づく品目の取得
 */
function getQuoteItem() {
  // 見積書IDの準備
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const quote = quotes.data[0];
  const itemId = quote.items[0].id;

  // API実行： 見積書に紐づく品目の取得
  const quoteItem = getMfClient_().quotes.getQuoteItem(quote.id, itemId);
  console.log(quoteItem);

  // スプレッドシートに追加
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quoteItems");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  const row = [];
  for (const attr in quoteItem) {
    row.push(quoteItem[attr]);
  }
  row.push(quote.id);
  sheet.appendRow(row);
}

/**
 * 見積書に紐づく品目を削除
 */
function deleteQuoteItem() {
  // 見積書IDの準備
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const quote = quotes.data[0];
  const itemId = quote.items[0].id;

  // API実行： 見積書に紐づく品目を削除
  const quoteItem = getMfClient_().quotes.deleteQuoteItem(quote.id, itemId);
  console.log(quoteItem);

  // スプレッドシートから削除
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quoteItems");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  if (sheet.getLastRow() - 1 === 0) {
    // 削除する行が無いので処理を中止する。
    return;
  }
  sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  ).getValues().forEach((row, index) => {
    const rowPosition = index + 2;
    const quoteItemId = row[0];
    if (itemId === quoteItemId) {
      sheet.deleteRow(rowPosition);
      // 削除に成功したら処理をやめる
      return;
    }
  });
}

/**
 * 見積書を請求書に変換
 */
function convertQuoteToBilling() {
  // 見積書の準備
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const quote = quotes.data[0];

  // API実行： 見積書を請求書に変換
  const billing = getMfClient_().quotes.convertQuoteToBilling(quote.id);
  console.log(billing);
}

/**
 * 見積書の削除
 */
function deleteQuote() {
  // 見積書IDの準備
  const baseDate = new Date();
  const dateUtil = MfInvoiceApi.getDateUtil(baseDate);
  const from = dateUtil.getEndDateLastMonth();
  const to = dateUtil.getEndDateNextMonth();
  const query = '';
  const quotes = getMfClient_().quotes.getQuotes(from, to, query);
  const deleteQuoteId = quotes.data[0].id;

  // API実行： 見積書の削除
  const result = getMfClient_().quotes.deleteQuote(deleteQuoteId);
  console.log(result);

  // スプレッドシートから削除
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("quotes");
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  if (sheet.getLastRow() - 1 === 0) {
    // 削除する行が無いので処理を中止する。
    return;
  }
  sheet.getRange(
    2,
    1,
    sheet.getLastRow() - 1,
    sheet.getLastColumn()
  ).getValues().forEach((row, index) => {
    const rowPosition = index + 2;
    const quoteId = row[0];
    if (deleteQuoteId === quoteId) {
      sheet.deleteRow(rowPosition);
      // 削除に成功したら処理をやめる
      return;
    }
  });
}
