/**
 * 初期設定
 * ・トリガー設定
 */
function initialize() {
  // トリガー設定
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
  MfInvoiceClient.showMfApiAuthDialog();
}

/**
 * MF認証のコールバック関数です。
 * @param request
 */
function mfCallback(request) {
  return MfInvoiceClient.mfCallback(request);
}

/**
 * MF請求書APIクライアントを生成します。
 * @returns {MfInvoiceClient}
 */
function getMfClient_() {
  MfInvoiceClient.createClient();
}

function testbillingsList() {
  const baseDate = new Date();
  const dateUtil = MfInvoiceClient.getDateUtil(baseDate);
  const to = dateUtil.getEndDateNextMonth();
  const from = dateUtil.getEndDateLastMonth();
  const query = '入金済み';
  console.log(getMfClient_().billings.getBillings(from, to, query));
}

function testQuotesList() {
  const baseDate = new Date();
  const dateUtil = MfInvoiceClient.getDateUtil(baseDate);
  const to = dateUtil.getEndDateNextMonth();
  const from = dateUtil.getEndDateLastMonth();
  const query = '';
  console.log(getMfClient_().quotes.getQuotes(from, to, query));
}

function testPartnersList() {
  console.log(getMfClient_().partners.getPartners());
}

function testItemList() {
  console.log(getMfClient_().items.getItems());
}

function testOffice() {
  console.log(getMfClient_().office.getMyOffice());
}
