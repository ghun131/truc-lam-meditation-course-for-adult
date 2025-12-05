function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Khoá tu")
    .addItem("Tạo danh sách gửi mail", "initDanhSachGuiMailSheet")
    .addItem("Đồng bộ form từ sheet lưu trữ", "syncFormFromSavedData")
    .addItem("Sync danh sách gửi mail", "syncDanhSachGuiMailSheet")
    .addItem("Lọc trùng thiền sinh", "filterDuplicate")
    .addToUi();

  ui.createMenu("Chạy chủ động")
    .addItem("Gửi mail xác nhận toàn bộ", "execSendMail")
    .addItem(
      "Gửi mail nhắc chuyển tiền xe toàn bộ",
      "testSendBusFeePaymentReminder"
    )
    .addToUi();
}

// ------------ CREATE MENU FUNCTIONS ------------
function initDanhSachGuiMailSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Danh sách gửi mail");
  const savingSheet = ss.getSheetByName("Lưu trữ");

  const sourceSheet = ss.getSheetByName("Câu trả lời biểu mẫu 1");
  if (!sourceSheet) {
    throw new Error(
      "Không tìm thấy sheet 'Câu trả lời biểu mẫu 1' để sao chép dữ liệu!"
    );
  }

  if (Boolean(sheet) && Boolean(savingSheet)) {
    return;
  }

  sheet = ss.insertSheet("Danh sách gửi mail");
  initLuuTruSheet(ss);
  console.log("Đã tạo danh sách gửi mail và Lưu trữ!");

  cloneSheetData(sourceSheet, sheet);

  const columns = [
    "Đã chuyển khoản",
    "Đã gửi mail đăng ký thành công",
    "Đã gửi mail nhắc chuyển tiền xe",
    "Thông báo",
    "Note",
  ];

  const lastColumn = sheet.getLastColumn();
  let currentHeaders = [];
  if (lastColumn > 0) {
    currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  }

  const startColumn = lastColumn + 1;

  const headersToAdd = [];
  for (let i = 0; i < columns.length; i++) {
    const column = columns[i];
    const exists = currentHeaders.some(
      (header) =>
        header && header.toString().toLowerCase() === column.toLowerCase()
    );

    if (!exists) {
      headersToAdd.push(column);
    }
  }

  if (headersToAdd.length > 0) {
    const headerRange = sheet.getRange(1, startColumn, 1, headersToAdd.length);
    headerRange.setValues([headersToAdd]);

    headerRange.setFontWeight("bold");
    headerRange.setFontColor("white");
    headerRange.setBackground("#5b3f86");
    headerRange.setBorder(true, true, true, true, true, true);

    for (let i = 0; i < headersToAdd.length; i++) {
      sheet.autoResizeColumn(startColumn + i);
    }

    console.log(
      `Added ${headersToAdd.length} new columns: ${headersToAdd.join(", ")}`
    );
  } else {
    console.log("All required columns already exist in the sheet.");
  }

  return sheet;
}

function syncDanhSachGuiMailSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Danh sách gửi mail");
  const sourceSheet = ss.getSheetByName("Câu trả lời biểu mẫu 1");
  cloneSheetData(sourceSheet, sheet);
}

function syncFormFromSavedData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Lưu trữ");
  const savedData = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();
  const savedDataMap = getSavedDataCode(savedData);
  const courseNameObj = savedDataMap.get("courseName");
  const formId = savedDataMap.get("formId").value;

  const form = FormApp.openById(formId);
  let desc = form.getDescription();
  if (desc.includes(courseNameObj.key)) {
    desc = replaceTextInDescription(desc, savedDataMap);
    form.setDescription(desc);

    // console.log(`Updated form description: ${desc}`);
  }

  const blocks = form.getItems();

  for (let i = 0; i < blocks.length; ++i) {
    const block = blocks[i];

    if ([3, 6].includes(i)) {
      let curBlock = null;

      if (block.getType() === FormApp.ItemType.IMAGE) {
        curBlock = block.asImageItem();
        const blockTitle = replaceTextInDescription(curBlock.getTitle(), savedDataMap);
        curBlock.setTitle(blockTitle);
        console.log(`Updated block ${curBlock.getTitle()} with ${blockTitle}`);
      }
      if (block.getType() === FormApp.ItemType.PAGE_BREAK) {
        curBlock = block.asPageBreakItem();
        const blockDesc = replaceTextInDescription(curBlock.getHelpText(), savedDataMap);
        curBlock.setHelpText(blockDesc);
        console.log(`Updated block ${curBlock.getTitle()} with ${blockDesc}`);
      }

    }
  }
}

function replaceTextInDescription(desc, savedDataMap) {
  savedDataMap.forEach((sValue) => {
    if (desc.includes(sValue.key)) {
      const currValue = sValue.value instanceof Date ?
        `${sValue.value.getDate()}/${sValue.value.getMonth() + 1}/${sValue.value.getFullYear()}` :
        sValue.value;

      desc = desc.replaceAll(sValue.key, currValue);
    }
  })

  return desc;
}

function filterDuplicate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Danh sách gửi mail");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  const hIndice = getHeadersIndices(data[0]);

  const emailIdx = hIndice.get("email");
  const nameIdx = hIndice.get("studentIdx");
  const dobIdx = hIndice.get("dateOfBirth");
  const reportIdx = hIndice.get("report");
  const markedIdx = hIndice.get("sttMarkedIdx");
  const confirmMailSentIdx = hIndice.get("confirmMailSent");
  const docCreatedIdx = hIndice.get("docCreateIdx");
  const remindingEmailIdx = hIndice.get("remindingMailIdx");

  let cache = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = row[emailIdx];
    const name = row[nameIdx];
    const dob = row[dobIdx];
    const studentObj = { idx: i, email, name, dob };
    if (Array.isArray(cache[email])) {
      cache[email].every(
        (item) => `${item.name}${item.dob}` !== `${name}{dob}`
      ) && cache[email].push(studentObj);
    } else {
      cache[email] = [studentObj];
    }
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = row[emailIdx];
    const name = row[nameIdx];
    const dob = row[dobIdx];
    const prevId = (name + dob.toString()).toLowerCase();

    if (cache[email] && cache[email].length > 0) {
      for (const item of cache[email]) {
        const currId = (item.name + item.dob.toString()).toLowerCase();

        if (currId === prevId && i < item.idx) {
          setRowBackgroundColor(sheet, "#F28C28", i);
          sheet
            .getRange(i + 1, reportIdx + 1)
            .setValue(`Trùng với ${item.name}`);
          markedIdx !== undefined &&
            sheet.getRange(i + 1, markedIdx + 1).setValue("x");
          confirmMailSentIdx !== undefined &&
            sheet.getRange(i + 1, confirmMailSentIdx + 1).setValue("x");
          docCreatedIdx !== undefined &&
            sheet.getRange(i + 1, docCreatedIdx + 1).setValue("x");
          remindingEmailIdx !== undefined &&
            sheet.getRange(i + 1, remindingEmailIdx + 1).setValue("x");

          console.log(`Dòng ${i + 1} trùng bạn ${item.name}, email: ${email}`);
        }
      }
    }
  }
}

// ------------ EMAIL TEMPLATE FUNCTIONS ------------

function createSuccessVerificationByBusMail(input) {
  const {
    courseName,
    startDate,
    endDate,
    targetAudience,
    numberOfStudents,
    busReadyTime,
    busStartTime,
    busLocation,
    busMapLink,
    zaloGroupLink,
    contactName,
    contactPhone,
    contactName2,
    contactPhone2,
    cancelDate,
    imageLink,
  } = input;
  return {
    subject: `[Khóa tu ${courseName}] Xác nhận đăng ký thành công - thiền sinh đi ô tô với Đoàn`,
    content: `
    <!DOCTYPE html>
    <html lang="vi">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Xác nhận đăng ký khóa tu ${courseName}</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                line-height: 1.6;
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
                color: #333;
            }
            h1, h2, h3, h4 {
                color: #333;
            }
            .greeting {
                color: #0000CD;
                font-weight: bold;
            }
            .section-title {
                font-weight: bold;
                margin-top: 20px;
                margin-bottom: 10px;
            }
            a {
                color: #0066cc;
                text-decoration: none;
            }
            a:hover {
                text-decoration: underline;
            }
            .signature {
                margin-top: 30px;
                font-style: italic;
            }
            .highlight {
                font-style: italic;
            }
        </style>
    </head>
    <body>
        <img src="${imageLink}" alt="Cẩm nang Thiền sinh"  style="width: 100%; height: auto">
        <p class="greeting">Thân chào bạn,</p>

        <p>Đoàn Thanh Thiếu Niên Phật Tử Trúc Lâm Tây Thiên xác nhận bạn đã đăng ký thành công tham gia <b>Khóa tu ${courseName}</b> tại Thiền viện Trúc Lâm Tây Thiên.</p>
        
        <div class="section-title">1. THÔNG TIN KHÓA TU</div>
        <p>- Thời gian: <b>${startDate} - ${endDate}</b></p>
        <p>- Địa điểm: Thiền viện Trúc Lâm Tây Thiên, Tam Đảo, Vĩnh Phúc</p>
        <p>- Đối tượng: ${targetAudience}</p>
        <p>- Số lượng: ${numberOfStudents} thiền sinh</p>
        <p>- Yêu cầu: Cam kết tham gia đủ ${calculateNumberOfDays(
      startDate,
      endDate
    )} ngày, tuân thủ nội quy khóa tu của Thiền Viện. Không sử dụng thiết bị điện tử cá nhân.</p>

        <div class="section-title">2. THÔNG TIN DI CHUYỂN</div>
        <p>- Thời gian tập trung: <b>${busReadyTime} ngày ${startDate}</b></p>
        <p>- Địa điểm: ${busLocation}. <a href="${busMapLink}">Định vị Google Maps</a></p>
        <p>- Thời gian xe xuất phát lên Thiền viện: <b>${busStartTime} cùng ngày</b></p>
        <p>- Thời gian kết thúc khóa tu, di chuyển về Hà Nội: <b>${endDate}</b>.</p>
        <p>- Thiền sinh cân nhắc trước khi đăng ký, Đoàn sẽ chỉ có thể hỗ trợ hoàn trả lệ phí đối với các trường hợp huỷ trước ngày <b>${cancelDate}</b>.</p>
        <p>- Thiền sinh hoan hỉ di chuyển tới địa điểm tập trung sớm hơn để tránh rơi vào tình trạng ùn tắc. Đoàn sẽ xuất phát theo đúng lịch trình và không chờ những trường hợp tới muộn.</p>
        <div class="section-title">3. TÀI LIỆU THAM KHẢO TRƯỚC KHÓA TU</div>
        <p>- <a href="http://www.thuongchieu.net/index.php/toathien">Phương pháp toạ thiền theo đường lối Thiền tông Việt Nam</a> - H.T Thích Thanh Từ</p>
        <p>- Nhóm Zalo Thiền sinh <a href="${zaloGroupLink}">Link</a></p>

        <div class="section-title">4. LƯU Ý CHUNG</div>
        <p>- Chuẩn bị ít nhất 2 bộ áo lam, 1 áo tràng.</p>
        <p>- Thiền sinh khi đã đăng ký khóa tu mà có việc đột xuất không tham gia được xin hoan hỷ báo lại sớm để BTC có thể kịp thời sắp xếp.</p>
        <p>- Tịnh tài cúng dường là <b>TÙY HỶ</b> để tạo phước đức cho bản thân và trợ duyên cho Thiền Viện chi phí tổ chức khoá tu.</p>
        
        <p>Mọi thông tin vui lòng liên hệ:</p>
        <p>1. ${contactName}: <b>${contactPhone}</b></p>
        <p>2. ${contactName2}: <b>${contactPhone2}</b></p>
        <p>Hẹn gặp lại bạn tại Khóa tu ${courseName} và chúc bạn một ngày an vui!</p>
        
        <p class="signature">Thân ái,</p>
        <p class="signature">TM. BAN TỔ CHỨC KHÓA TU ${courseName}</p>
      </body>
    </html>
  `,
  };
}

function createSuccessVerificationOwnVehicleMail(input) {
  const {
    courseName,
    startDate,
    endDate,
    targetAudience,
    numberOfStudents,
    arrivalTime,
    zaloGroupLink,
    contactName,
    contactPhone,
    contactName2,
    contactPhone2,
    imageLink,
  } = input;
  return {
    subject: `[Khóa tu ${courseName}] Xác nhận đăng ký thành công với Thiền sinh di chuyển tự túc`,
    content: `
    <!DOCTYPE html>
    <html lang="vi">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>[KT ${courseName}] Xác nhận đăng ký thành công với Thiền sinh di chuyển tự túc</title>
      <style>
          body {
              font-family: Arial, sans-serif;
              line-height: 1.6;
              max-width: 800px;
              margin: 0 auto;
              padding: 20px;
              color: #333;
          }
          h1, h2, h3, h4 {
              color: #333;
          }
          .greeting {
              color: #0000CD;
              font-weight: bold;
          }
          .section-title {
              font-weight: bold;
              margin-top: 20px;
              margin-bottom: 10px;
          }
          a {
              color: #0066cc;
              text-decoration: none;
          }
          a:hover {
              text-decoration: underline;
          }
          .signature {
              margin-top: 30px;
              font-style: italic;
          }
          .highlight {
              font-style: italic;
          }
          .email-header {
              border-bottom: 1px solid #eee;
              padding-bottom: 10px;
              margin-bottom: 20px;
          }
          .sender-info {
              display: flex;
              align-items: center;
              margin-bottom: 10px;
          }
          .sender-avatar {
              width: 48px;
              height: 48px;
              border-radius: 50%;
              margin-right: 10px;
              background-color: #f0f0f0;
          }
          .sender-name {
              font-weight: bold;
          }
          .sender-email {
              color: #666;
              font-size: 0.9em;
          }
          .email-timestamp {
              color: #666;
              font-size: 0.9em;
              text-align: right;
          }
          .translate-button {
              background-color: #f8f9fa;
              border: 1px solid #dadce0;
              border-radius: 4px;
              padding: 8px 12px;
              margin: 10px 0;
              display: inline-flex;
              align-items: center;
          }
          .translate-button:hover {
              background-color: #f1f3f4;
          }
          .translate-button .close {
              margin-left: 10px;
              color: #5f6368;
          }
      </style>
  </head>
  <body>
     <img src="${imageLink}" alt="Cẩm nang Thiền sinh"  style="width: 100%; height: auto">
      <p class="greeting">Thân chào bạn,</p>

      <p>Đoàn Thanh Thiếu Niên Phật Tử Trúc Lâm Tây Thiên xác nhận bạn đã đăng ký thành công tham gia <b>Khóa tu ${courseName}</b> tại Thiền viện Trúc Lâm Tây Thiên.</p>
      
      <div class="section-title">1. THÔNG TIN KHÓA TU</div>
      <p>- Thời gian: <b>${startDate} - ${endDate}</b></p>
      <p>- Địa điểm: Thiền viện Trúc Lâm Tây Thiên, Tam Đảo, Vĩnh Phúc</p>
      <p>- Đối tượng: ${targetAudience}</p>
      <p>- Số lượng: ${numberOfStudents} thiền sinh</p>
      <p>- Thời gian tập trung: Thiền sinh hoan hỉ có mặt tại giảng đường Thiền Viện <b>trước ${arrivalTime}</b> để hoàn tất đăng ký và làm thủ tục nhập khóa.</p>
      <p>- Yêu cầu: Cam kết tham gia đủ ${calculateNumberOfDays(
      startDate,
      endDate
    )} ngày, tuân thủ nội quy khóa tu của Thiền Viện. Không sử dụng thiết bị điện tử cá nhân.</p>

      <div class="section-title">2. TÀI LIỆU THAM KHẢO TRƯỚC KHÓA TU</div>
      <p>- <a href="http://www.thuongchieu.net/index.php/toathien">Phương pháp toạ thiền theo đường lối Thiền tông Việt Nam</a> - H.T Thích Thanh Từ</p>
      <p>- Nhóm Zalo Thiền sinh <a href="${zaloGroupLink}">Link</a></p>

      
      <div class="section-title">3. LƯU Ý CHUNG</div>
      <p>- Chuẩn bị ít nhất 2 bộ áo lam, 1 áo tràng.</p>
      <p>- Khi có việc đột xuất không tham gia được khóa tu, mong bạn báo lại sớm để Ban tổ chức có thể kịp thời sắp xếp.</p>
      <p>- Tịnh tài cúng dường là <b>TÙY HỶ</b> để tạo phước đức cho bản thân và trợ duyên cho Thiền Viện chi phí tổ chức khoá tu.</p>
      
      <p>Mọi thông tin vui lòng liên hệ:</p>
      <p>1. ${contactName}: <b>${contactPhone}</b></p>
      <p>2. ${contactName2}: <b>${contactPhone2}</b></p>
      <p>Hẹn gặp lại bạn tại Khóa tu ${courseName} và chúc bạn một ngày an vui!</p>
      
      
      <p class="signature">Thân ái,</p>
      <p class="signature">TM. BAN TỔ CHỨC KHÓA TU ${courseName}</p>
      
  </body>
  </html>
`,
  };
}

function createPaymentReminderMail(input) {
  const {
    courseName,
    cancelDate,
    busFee,
    bankName,
    bankAccountNumber,
    bankAccountName,
  } = input;
  return {
    subject: `[Khóa tu ${courseName}] Thư nhắc v/v chưa đăng ký thành công Khoá tu`,
    content: `
  <!DOCTYPE html>
  <html lang="vi">
  <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>[Khóa tu ${courseName}] Yêu cầu thanh toán lệ phí đi xe ô tô</title>
      <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            color: #333;
        }
        h1, h2, h3, h4 {
            color: #333;
        }
        .greeting {
            color: #0000CD;
            font-weight: bold;
            font-style: italic;
        }
        .section-title {
            font-weight: bold;
            margin-top: 20px;
            margin-bottom: 10px;
        }
        a {
            color: #0066cc;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
        .signature {
            margin-top: 30px;
            font-style: italic;
        }
        .highlight {
            font-weight: bold;
            text-decoration: underline;
        }
        .important {
            font-weight: bold;
        }
        .indent {
            margin-left: 20px;
        }
        .emoji {
            display: inline-block;
            margin-left: 5px;
        }
      </style>
  </head>
  <body>
      <p class="greeting">Thân chào bạn,</p>

      <p>Đoàn Thanh Thiếu Niên Phật Tử Trúc Lâm Tây Thiên - Trần Nhân Tông đã nhận được thông tin của bạn đăng ký đi xe ô tô với Đoàn tham gia Khoá tu ${courseName} tại Thiền viện Trúc Lâm Tây Thiên. Tuy nhiên, Đoàn vẫn chưa nhận được thông tin chuyển khoản lệ phí đi xe ô tô của bạn.</p>

      <p>Bạn vui lòng hoàn thành chuyển khoản lệ phí đi xe ô tô với Đoàn để được xác nhận đăng kí thành công tham gia làm thiền sinh khóa tu. Cụ thể:</p>
      
      <div class="section-title">1. Xác nhận đăng ký đối với Thiền sinh đi xe ô tô Đoàn tổ chức:</div>
      <p>- Chuyển khoản lệ phí đi xe ô tô: ${busFee}/người/2 chiều.</p>
      <p>- Thông tin chuyển khoản lệ phí xe ô tô:</p>
      <p class="indent">+ Ngân hàng ${bankName} - Chủ TK: ${bankAccountName}- Số TK: ${bankAccountNumber}</p>
      <p class="indent">+ Nội dung chuyển khoản <span class="important">BẮT BUỘC CẦN</span> ghi rõ: Họ tên người Đăng kí – SĐT – ${courseName}</p>
      
      <p class="indent">+ Chụp ảnh màn hình đã chuyển khoản để đính kèm và trả lời vào email (ghi rõ Họ tên - SĐT đã đăng kí). <span class="emoji">☘️</span></p>
      
      <p class="section-title">2. Sau khi hoàn thành chuyển khoản đăng kí đi xe ô tô với Đoàn: <span class="highlight">Đoàn sẽ gửi email xác nhận bạn đã đăng kí thành công.</span></p>
      
      <p>Trường hợp bạn không hoàn thành trước ngày ${cancelDate}, Ban tổ chức xác nhận bạn huỷ đăng ký tham gia Khoá tu ${courseName}</p>
      
      <p class="signature">Chúc bạn ngày an vui,</p>
      <p class="signature">Đoàn Thanh Thiếu Niên Phật Tử Trúc Lâm Tây Thiên.</p>
  </body>
</html> 
  `,
  };
}

function execSendMail() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Danh sách gửi mail");

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow === 0) {
    console.log("No data found in the sheet.");
    return;
  }

  const allData = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

  for (let row = 0; row < allData.length; row++) {
    if (row === 0) continue;

    const rowData = allData[row];
    const email = rowData[10]; // Column K
    const vehicle = rowData[1]; // Column B
    const byBus = vehicle === "Đi ô tô cùng Đoàn";
    const paidBusFee = rowData[12] ?
      rowData[12].toLowerCase() :
      ''; // Column M
    const personalVehicle = vehicle === "Tự túc phương tiện";
    const confirmMailSent = rowData[13]; // Column N
    const sentReminderMail = rowData[14]; // Column O

    if (
      email &&
      !confirmMailSent &&
      (personalVehicle || (byBus && paidBusFee.includes("x")))
    ) {
      console.log(`Send successful registration email to: ${email}`);
      sendRegisterSuccessful({ sheet, row, email, byBus });
    }

    const savedSheet = spreadsheet.getSheetByName("Lưu trữ");
    const savedData = savedSheet
      .getRange(1, 1, savedSheet.getLastRow(), savedSheet.getLastColumn())
      .getValues();
    const savedDataMap = getSavedData(savedData);
    const paymentDeadline = savedDataMap.get("deadlinePayment");

    if (
      byBus &&
      !paidBusFee.includes("x") &&
      email &&
      !confirmMailSent &&
      !sentReminderMail &&
      isLatePayment(paymentDeadline)
    ) {
      console.log(`Sent payment reminder to ${email}`);
      sendBusFeePaymentReminder(sheet, row, email);
    }

    // Clean up errored row has been fixed
    const hasError =
      rowData[16] === "Lỗi mail phí xe!" ||
      rowData[16] === "Lỗi mail xác nhận!"; // Column R
    if ((confirmMailSent || sentReminderMail) && hasError) {
      setRowBackgroundColor(sheet, "white", row);
      sheet.getRange(row + 1, 16).setValue("");
    }
  }
}

function sendRegisterSuccessful({ sheet, row, email, byBus }) {
  const savedDataMap = getSavedData();
  const commonData = {
    courseName: savedDataMap.get("courseName"),
    startDate: formatDate(savedDataMap.get("startDate")),
    endDate: formatDate(savedDataMap.get("endDate")),
    targetAudience: savedDataMap.get("targetAudience"),
    numberOfStudents: savedDataMap.get("numberOfStudents"),
    zaloGroupLink: savedDataMap.get("zaloGroupLink"),
    contactName: savedDataMap.get("contactName"),
    contactPhone: savedDataMap.get("contactPhone"),
    contactName2: savedDataMap.get("contactName2"),
    contactPhone2: savedDataMap.get("contactPhone2"),
    cancelDate: formatDate(savedDataMap.get("cancelDate")),
    imageLink: savedDataMap.get("imageLink"),
  };
  const successVerificationByBusMail = createSuccessVerificationByBusMail({
    ...commonData,
    busReadyTime: savedDataMap.get("busReadyTime"),
    busStartTime: savedDataMap.get("busStartTime"),
    busLocation: savedDataMap.get("busLocation"),
    busMapLink: savedDataMap.get("busMapLink"),
  });
  try {
    if (byBus) {
      GmailApp.sendEmail(email, successVerificationByBusMail.subject, "", {
        htmlBody: successVerificationByBusMail.content,
      });
      console.log(`Sent to ${email} by bus`);
    } else {
      const successVerificationOwnVehicleMail =
        createSuccessVerificationOwnVehicleMail({
          ...commonData,
          arrivalTime: savedDataMap.get("arrivalTime"),
        });
      GmailApp.sendEmail(email, successVerificationOwnVehicleMail.subject, "", {
        htmlBody: successVerificationOwnVehicleMail.content,
      });
      console.log(`Sent to ${email} own vehicle`);
    }
    sheet.getRange(row + 1, 14).setValue("x"); // Write to column P
    setRowBackgroundColor(sheet, "white", row);
    sheet.getRange(row + 1, 16).setValue("");
  } catch (error) {
    console.log(`sendingRegisterSuccessful: ${error}`);
    setRowBackgroundColor(sheet, "#ffdddd", row);
    sheet.getRange(row + 1, 16).setValue("Lỗi mail xác nhận!");
  }
}

function testSendBusFeePaymentReminder() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Danh sách gửi mail");

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow === 0) {
    console.log("No data found in the sheet.");
    return;
  }

  const allData = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

  for (let row = 0; row < allData.length; row++) {
    if (row === 0) continue;

    const rowData = allData[row];
    const vehicle = rowData[1]; // Column B
    const byBus = vehicle === "Đi ô tô cùng Đoàn";
    const sentReminderMail = rowData[14]; // Column O
    const email = rowData[10]; // Column K

    if (sentReminderMail.toLowerCase() === "x" || !email || !byBus) {
      console.log(`testSendBusFeePaymentReminder: already sent or invalid`);
      continue;
    }
    sendBusFeePaymentReminder(sheet, row, email);
    console.log(`testSendBusFeePaymentReminder: sent`);
  }
}

function sendBusFeePaymentReminder(sheet, row, email) {
  const savedDataMap = getSavedData();
  const paymentReminderMail = createPaymentReminderMail({
    courseName: savedDataMap.get("courseName"),
    cancelDate: formatDate(savedDataMap.get("cancelDate")),
    busFee: savedDataMap.get("busFee"),
    bankName: savedDataMap.get("bankName"),
    bankAccountNumber: savedDataMap.get("bankAccountNumber"),
    bankAccountName: savedDataMap.get("bankAccountName"),
  });
  try {
    GmailApp.sendEmail(email, paymentReminderMail.subject, "", {
      htmlBody: paymentReminderMail.content,
    });
    sheet.getRange(row + 1, 15).setValue("x");
  } catch (error) {
    console.log(`sendBusFeePaymentReminder: ${error}`);
    setRowBackgroundColor(sheet, "#ffdddd", row);
    sheet.getRange(row + 1, 18).setValue("Lỗi mail phí xe!");
  }
}

// ------------ UTILITY FUNCTIONS ------------

function setRowBackgroundColor(sheet, color, row) {
  const rowRange = sheet.getRange(row + 1, 1, 1, sheet.getLastColumn());
  rowRange.setBackground(color);
}

function isLatePayment(date) {
  if (!date) {
    console.log("Can not process empty date");
    return false;
  }

  const now = new Date();

  return (
    now.getDate() === date.getDate() &&
    now.getMonth() === date.getMonth() &&
    now.getFullYear() === date.getFullYear()
  );
}

function cloneSheetData(sourceSheet, targetSheet) {
  // Clone all data from source sheet
  const lastSelectedColumn = 12; // From column A to L, 12 questions
  const sourceRange = sourceSheet.getRange(1, 1, sourceSheet.getLastRow(), lastSelectedColumn);
  if (sourceRange.getNumRows() > 0) {
    const sourceData = sourceRange.getValues();
    const targetRange = targetSheet.getRange(
      1,
      1,
      sourceData.length,
      lastSelectedColumn,
    );
    targetRange.setValues(sourceData);

    // Copy formatting from first row (headers)
    const sourceHeaderRange = sourceSheet.getRange(
      1,
      1,
      1,
      lastSelectedColumn
    );
    const targetHeaderRange = targetSheet.getRange(
      1,
      1,
      1,
      lastSelectedColumn
    );
    sourceHeaderRange.copyTo(targetHeaderRange);

    console.log(
      `Đã sao chép ${sourceData.length
      } hàng từ "${sourceSheet.getName()}" sang "${targetSheet.getName()}"`
    );
    return sourceData.length;
  }
  return 0;
}
function initLuuTruSheet(ss) {
  const sheet = ss.insertSheet("Lưu trữ");

  // Define the data to populate (label, value)
  const data = [
    ["Tên khoá tu", "Tuệ Giác VI", "TEN_KT"], // Row 0
    ["Ngày bắt đầu", new Date(2025, 7, 12), "NGAY_BAT_DAU"], // Row 1
    ["Ngày kết thúc", new Date(2025, 7, 16), "NGAY_KET_THUC"], // Row 2
    ["Đối tượng", "Nam, Nữ, sinh năm 1990 - năm 2008", "DOI_TUONG"], // Row 3
    ["Số lượng thiền sinh", 300, "SO_LUONG_TS"], // Row 4
    [
      "Địa điểm tập trung đi xe đoàn",
      "cổng Đông công viên Hoà Bình, đường Đỗ Nhuận, Bắc Từ Liêm, Hà Nội (Đối diện bệnh viện Mặt Trời - SunGroup)",
      "DIA_DIEM_TAP_TRUNG"
    ], // Row 5
    [
      "Link địa điểm tập trung đi xe đoàn",
      "https://maps.app.goo.gl/UprGfvKKzuKrwoQr7",
      "LINK_DIA_DIEM_TAP_TRUNG"
    ], // Row 6
    ["Thời gian tập trung", "6h00", "THOI_GIAN_TAP_TRUNG"], // Row 7
    ["Thời gian xe xuất phát", "7h00", "THOI_GIAN_XE_XUAT_PHAT"], // Row 8
    ["Thời gian có mặt tại thiền viện", "9h00", "THOI_GIAN_CO_MAT"], // Row 9
    ["Hạn chót ngày huỷ đăng ký cho thiền sinh", new Date(2025, 7, 25), "HAN_CHOT"], // Row 10
    ["Link nhóm Zalo", "https://www.google.com", "LINK_NHOM_ZALO"], // Row 11
    ["Tên đường dây nóng 1", "Phật tử Diệu Từ", "TEN_DUONG_DAY_NONG_1"], // Row 12
    ["Số điện thoại", "0988 237 713", "SO_DIEN_THOAI_1"], // Row 13
    ["Tên đường dây nóng 2", "Phật tử Chân Mỹ Nghiêm", "TEN_DUONG_DAY_NONG_2"], // Row 14
    ["Số điện thoại", "0848 349 129", "SO_DIEN_THOAI_2"], // Row 15
    ["Lệ phí đi xe đoàn (1 người/2 chiều)", "180,000 VND", "LEPHI_XE_DOAN"], // Row 16
    ["Ngân hàng người chịu trách nhiệm nhận tiền", "VIETINBANK", "TKNH"], // Row 17
    ["Tên chủ tài khoản ", "Mẫn Thị Thảo", "CHU_TK_NGAN_HANG"], // Row 18
    ["Số tài khoản", "123456789", "SO_TK_NH"], // Row 19
    ["Ngày nhắc thanh toán tiền", new Date(2025, 7, 20), "NGAY_NHAC_THANH_TOAN"], // Row 20
    [
      "Link ảnh trên mail",
      "https://ghun131.github.io/meditation-course-images/ktmh_khoa_5_2025.jpg",
      "LINK_ANH_MAIL"
    ], // Row 21
    ["Form Id", "", "FORM_ID"], // Row 22
    ["Giờ kết thúc khoá tu", "17h", "GIO_KET_THUC"], // Row 23
    ["Số ngày tham gia tu tập", 5, "SO_NGAY_TU"], // Row 24
    ["Ngày thông báo", new Date(2025, 7, 10), "NGAY_THONG_BAO"], // Row 25
  ];

  // Set the data
  const range = sheet.getRange(1, 1, data.length, 3);
  range.setValues(data);

  // Format the header row if you want
  const headerRange = sheet.getRange(1, 1, data.length, 1);
  headerRange.setFontWeight("bold");

  // Auto-resize columns
  sheet.autoResizeColumns(1, 3);

  console.log("Lưu trữ sheet initialized with configuration data");

  return sheet;
}

function calculateNumberOfDays(startDate, endDate) {
  const startDateArr = startDate.split("/");
  const endDateArr = endDate.split("/");
  const start = new Date(startDateArr[2], startDateArr[1], startDateArr[0]);
  const end = new Date(endDateArr[2], endDateArr[1], endDateArr[0]);
  const diffTime = Math.abs(end - start);
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  return diffDays;
}

function getHeadersIndices(headerData) {
  const result = new Map();

  for (let i = 0; i < headerData.length; i++) {
    const header = headerData[i].toLowerCase();

    if (
      ["phương tiện", "cách thức", "hình thức"].some((val) =>
        header.includes(val)
      ) &&
      (header.includes("di chuyển") || header.includes("đi lại"))
    ) {
      result.set("vehicle", i);
    }

    if (header.includes("email")) {
      const currValue = result.get("email") || [];
      currValue.length > 0
        ? result.set("email", [i, ...currValue])
        : result.set("email", [i]);
    }

    if (header === "số thứ tự") {
      result.set("stt", i);
    }

    if (header === "đã đánh số thứ tự") {
      result.set("sttMarkedIdx", i);
    }

    if (header === "đã gửi mail nhắc chuyển tiền xe") {
      result.set("remindingMailIdx", i);
    }

    if (header === "đã tạo đơn đăng ký") {
      result.set("docCreateIdx", i);
    }

    if (header === "họ và tên của bạn là?") {
      result.set("studentIdx", i);
    }

    if (header.includes("đăng ký thành công")) {
      result.set("confirmMailSent", i);
    }

    if (header === "thông báo") {
      result.set("report", i);
    }

    if (
      header.includes("chuyển khoản") ||
      header.includes("thanh toán") ||
      header.includes("trả tiền")
    ) {
      result.set("payment", i);
    }

    if (header === "generated document link") {
      result.set("genDocFile", i);
    }

    if (header === "ghi chú") {
      result.set("note", i);
    }

    if (header === "ngày/tháng/năm sinh của bạn") {
      result.set("dateOfBirth", i);
    }
  }

  return result;
}

function getSavedData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lưu trữ");
  const savedData = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();

  const result = new Map([
    ["courseName", savedData[0][1]], // Tên khoá tu
    ["startDate", savedData[1][1]], // Ngày bắt đầu
    ["endDate", savedData[2][1]], // Ngày kết thúc
    ["targetAudience", savedData[3][1]], // Đối tượng
    ["numberOfStudents", savedData[4][1]], // Số lượng thiền sinh
    ["busLocation", savedData[5][1]], // Địa điểm tập trung đi xe đoàn
    ["busMapLink", savedData[6][1]], // Link địa điểm tập trung đi xe đoàn
    ["busReadyTime", savedData[7][1]], // Thời gian tập trung xe đoàn
    ["busStartTime", savedData[8][1]], // Thời gian xe đoàn xuất phát
    ["arrivalTime", savedData[9][1]], // Thời gian có mặt tại thiền viện
    ["cancelDate", savedData[10][1]], // Hạn chót ngày huỷ đăng ký cho thiền sinh
    ["zaloGroupLink", savedData[11][1]], // Link nhóm Zalo
    ["contactName", savedData[12][1]], // Tên người liên hệ 1
    ["contactPhone", savedData[13][1]], // Số điện thoại 1
    ["contactName2", savedData[14][1]], // Tên người liên hệ 2
    ["contactPhone2", savedData[15][1]], // Số điện thoại 2
    ["busFee", savedData[16][1]], // Lệ phí xe đoàn
    ["bankName", savedData[17][1]], // Tên ngân hàng
    ["bankAccountName", savedData[18][1]], // Tên tài khoản ngân hàng
    ["bankAccountNumber", savedData[19][1]], // Số tài khoản ngân hàng
    ["deadlinePayment", savedData[20][1]], // Hạn chót thanh toán
    ["imageLink", savedData[21][1]], // Link ảnh trên mail
  ]);

  return result;
}

function getSavedDataCode() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lưu trữ");
  const savedData = sheet
    .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .getValues();

  const result = new Map([
    ["courseName", { value: savedData[0][1], key: savedData[0][2] }],
    ["startDate", { value: savedData[1][1], key: savedData[1][2] }], // Ngày bắt đầu
    ["endDate", { value: savedData[2][1], key: savedData[2][2] }], // Ngày kết thúc
    ["targetAudience", { value: savedData[3][1], key: savedData[3][2] }], // Đối tượng
    ["numberOfStudents", { value: savedData[4][1], key: savedData[4][2] }], // Số lượng thiền sinh
    ["busLocation", { value: savedData[5][1], key: savedData[5][2] }], // Địa điểm tập trung đi xe đoàn
    ["busMapLink", { value: savedData[6][1], key: savedData[6][2] }], // Link địa điểm tập trung đi xe đoàn
    ["busReadyTime", { value: savedData[7][1], key: savedData[7][2] }], // Thời gian tập trung xe đoàn
    ["busStartTime", { value: savedData[8][1], key: savedData[8][2] }], // Thời gian xe đoàn xuất phát
    ["arrivalTime", { value: savedData[9][1], key: savedData[9][2] }], // Thời gian có mặt tại thiền viện
    ["cancelDate", { value: savedData[10][1], key: savedData[10][2] }], // Hạn chót ngày huỷ đăng ký cho thiền sinh
    ["zaloGroupLink", { value: savedData[11][1], key: savedData[11][2] }], // Link nhóm Zalo
    ["contactName", { value: savedData[12][1], key: savedData[12][2] }], // Tên người liên hệ 1
    ["contactPhone", { value: savedData[13][1], key: savedData[13][2] }], // Số điện thoại 1
    ["contactName2", { value: savedData[14][1], key: savedData[14][2] }], // Tên người liên hệ 2
    ["contactPhone2", { value: savedData[15][1], key: savedData[15][2] }], // Số điện thoại 2
    ["busFee", { value: savedData[16][1], key: savedData[16][2] }], // Lệ phí xe đoàn
    ["bankName", { value: savedData[17][1], key: savedData[17][2] }], // Tên ngân hàng
    ["bankAccountName", { value: savedData[18][1], key: savedData[18][2] }], // Tên tài khoản ngân hàng
    ["bankAccountNumber", { value: savedData[19][1], key: savedData[19][2] }], // Số tài khoản ngân hàng
    ["deadlinePayment", { value: savedData[20][1], key: savedData[20][2] }], // Hạn chót thanh toán
    ["imageLink", { value: savedData[21][1], key: savedData[21][2] }], // Link ảnh trên mail
    ["formId", { value: savedData[22][1], key: 'FORM_ID' }], // ID form
    ["endingTime", { value: savedData[23][1], key: savedData[23][2] }], // Giờ kết thúc khoá tu
    ["numberOfDays", { value: savedData[24][1], key: savedData[24][2] }], // Số ngày tham gia tu tập
    ["announcementDate", { value: savedData[25][1], key: savedData[25][2] }], // Ngày thông báo
  ]);

  return result;
}

function formatDate(dateObj) {
  if (!dateObj) {
    return "";
  }

  const day = dateObj.getDate() || "00";
  const month = dateObj.getMonth() + 1 || "00";
  const year = dateObj.getFullYear() || "0000";
  return `${day}/${month}/${year}`;
}
