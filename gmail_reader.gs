/***************************************
 * CONFIG
 ***************************************/
const CONFIG = {
  QUERY: "after:2024/12/31 before:2026/01/01", // Gmail search query (Year 2025)
  SHEET_NAME: "GmailData",
  SHEET_ID: "17PtcY1OmToTB6NxyD7alAuuEKGQRnUOmuW0LN3lvsQE",
  RAW_LIMIT: 5000, // giới hạn content
  BATCH_SIZE: 100, // Số lượng mail lấy mỗi lần
};

let count_total_read = 0;
/***************************************
 * MAIN
 ***************************************/
function runReadGmail() {
  const startTime = new Date().getTime();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000; // Giới hạn 4.5 phút để an toàn
  const props = PropertiesService.getScriptProperties();

  // Lấy vị trí đã lưu từ lần chạy trước (nếu có)
  let start = parseInt(props.getProperty('LAST_START_INDEX')) || 0;
  let threads;

  console.log(`---> Bắt đầu chạy. Vị trí bắt đầu: ${start}`);

  do {
    // Kiểm tra thời gian, nếu sắp hết thì dừng và lưu vị trí
    if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
      props.setProperty('LAST_START_INDEX', start.toString());
      console.log(`⚠️ SẮP HẾT THỜI GIAN! Đã lưu vị trí ${start}. Vui lòng CHẠY LẠI script để tiếp tục.`);
      return;
    }

    console.log(`---> Đang tải batch từ index: ${start}`);
    threads = GmailApp.search(CONFIG.QUERY, start, CONFIG.BATCH_SIZE);

    if (threads.length === 0) break;

    threads.forEach((thread) => {
      const messages = thread.getMessages();
      messages.forEach((message) => {
        handleMessage(message);
        count_total_read++;
      });
    });

    start += CONFIG.BATCH_SIZE;
  } while (threads.length === CONFIG.BATCH_SIZE);

  // Nếu chạy xong hết thì xóa vị trí đã lưu để lần sau chạy lại từ đầu
  props.deleteProperty('LAST_START_INDEX');
  console.log(`✅ HOÀN TẤT TOÀN BỘ! Tổng số tin nhắn đã xử lý trong phiên này: ${count_total_read}`);
}

/***************************************
 * MESSAGE HANDLER
 ***************************************/
function handleMessage(message) {
  const messageId = message.getId();
  console.log(`${count_total_read} Processing message from ${messageId}`);

  if (isProcessed(messageId)) {
    return;
  }

  const body = message.getPlainBody();
  const emailFrom = message.getFrom();
  if (!isSendToVietcombank(emailFrom)) {
    console.log(`Skipping message from ${emailFrom}, not sent to Vietcombank.`);
    return;
  }
  const extracted = extractDataBankTransfer(body);

  let beneficiaryName = extracted.beneficiaryName;
  if (!beneficiaryName) {
    beneficiaryName = extracted.beneficiaryName_sub;
  }
  let beneficiaryBank = extracted.beneficiaryBank;
  if (!beneficiaryBank) {
    beneficiaryBank = extracted.creditAccount;
  }
  let paymentDetail = extracted.paymentDetail;
  if (!paymentDetail) {
    paymentDetail = extracted.paymentDetail_sub;
  }

  const row = [
    messageId,
    message.getSubject(),
    message.getDate(),
    beneficiaryName,
    beneficiaryBank,
    extracted.amount,
    extracted.currency,
    paymentDetail,
  ];

  saveRow(row);

  // Optional: gắn label sau khi xử lý
  // message.getThread().addLabel(getProcessedLabel());
}

/***************************************
 * DATA EXTRACT
 ***************************************/
function extractDataBankTransfer(body) {
  return {
    type: "BANK_TRANSFER_VIETCOMBANK",

    remitterName: match(body, /\*Remitter’s name\*\s*([A-Z\s]+)/i),

    beneficiaryName: match(body, /\*Beneficiary Name\*\s*([A-Z0-9_\s]+)/i),
    beneficiaryName_sub: match(
      body,
      /\*Beneficiary Name\s*\*\s*([A-Z0-9\s]+)/i,
    ),
    beneficiaryBank: match(body, /\*Beneficiary Bank Name\*\s*(.+)/i),
    beneficiaryBank_sub: match(body, /\*Beneficiary Bank Name\*\s*(.+)/i),
    creditAccount: match(body, /\*Credit Account\*\s*(\d+)/i),
    amount: parseMoney(match(body, /\*Amount\*\s*([\d,]+)\s*VND/i)),
    currency: "VND",
    chargeCode: match(body, /\*Charge Code\*\s*(.+)/i),
    paymentDetail: match(body, /\*Details of Payment\*\s*(.+)/i),
    paymentDetail_sub: match(
      body,
      /\*Details of Payment\s*\*\s*([\s\S]+?)\n\n/i,
    ),

    rawContent: body.substring(0, CONFIG.RAW_LIMIT),
  };
}

/***********************
 * HELPER
 ***********************/
function match(text, regex) {
  const m = text.match(regex);
  return m ? m[1].trim() : "";
}

function parseMoney(value) {
  if (!value) return 0;
  return Number(value.replace(/,/g, ""));
}

const isSendToVietcombank = (emailFrom) => {
  return emailFrom.includes("VCBDigibank");
};

function matchFirst(text, regex) {
  const m = text.match(regex);
  return m ? m[1] : null;
}

/***************************************
 * GOOGLE SHEET
 ***************************************/
function saveRow(row) {
  getSheet().appendRow(row);
}

function isProcessed(messageId) {
  const sheet = getSheet();
  if (sheet.getLastRow() < 2) return false;

  const ids = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getValues()
    .flat();

  return ids.includes(messageId);
}

function getSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow([
      "MessageID",
      "From",
      "Subject",
      "Date",
      "OrderID",
      "Price",
      "RawContent",
    ]);
  }

  return sheet;
}

/*********************************
 * LABEL (OPTIONAL)
 ***************************************/
function getProcessedLabel() {
  const labelName = "PROCESSED_BY_SCRIPT";
  let label = GmailApp.getUserLabelByName(labelName);

  if (!label) {
    label = GmailApp.createLabel(labelName);
  }

  return label;
}
