const SHEET_NAME = '임원신청서';
const PRINT_SHEET_NAME = '출력용_명단';

const HEADERS = [
  '연번',
  '성명',
  '사진',
  '1지망',
  '2지망',
  '생년월일',
  '주소',
  '핸드폰',
  '이메일',
  '경력',
  '서명',
  '추천인',
  '개인정보 수집ㆍ이용 동의',
  '신청일',
  '책임당원 여부'
];

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('국민의힘 책임당원협의회 서울서부본부 임원신청서')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 외부(깃허브 페이지 등) 폼에서 fetch 로 제출한 JSON을 받는 진입점.
 * - 폼 HTML을 Apps Script가 아닌 일반 웹호스팅에 올리면, 방문자는 script.google.com
 *   화면을 아예 보지 않으므로 구글 로그인/다중계정 문제가 발생하지 않는다.
 * - 요청은 Content-Type: text/plain 으로 보내 CORS 프리플라이트(OPTIONS)를 피한다.
 * - ContentService 응답에는 Access-Control-Allow-Origin 이 자동으로 붙어 브라우저가 결과를 읽을 수 있다.
 */
function doPost(e) {
  var out;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('요청 본문이 비어 있습니다.');
    }
    var data = JSON.parse(e.postData.contents);
    var result = submitForm(data);
    out = { ok: true, serialNo: result && result.serialNo, updated: !!(result && result.updated) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('임원신청 관리')
    .addItem('출력용 명단 생성', 'createPrintSheets')
    .addToUi();
}

/**
 * 3번 요구사항 반영: 신청서 제출 속도 개선
 * - 제출 시 부동 이미지 삽입 작업을 하지 않고 Drive 링크만 기록하여 1~2초 내에 응답 완료
 */
function submitForm(data) {
  if (!data) throw new Error('입력 정보를 확인해 주세요.');
  if (!data.isMember) throw new Error('국민의힘 책임당원 여부에 체크해 주세요.');
  if (!data.name) throw new Error('성명을 입력해 주세요.');
  if (!data.position1) throw new Error('희망직책 1지망을 선택해 주세요.');
  if (!data.position2) throw new Error('희망직책 2지망을 선택해 주세요.');
  if (data.position1 === data.position2) throw new Error('1지망과 2지망은 다른 직책을 선택해 주세요.');
  if (!data.birth) throw new Error('생년월일을 입력해 주세요.');
  if (!data.address) throw new Error('주소를 입력해 주세요.');
  if (!data.phone) throw new Error('핸드폰 번호를 입력해 주세요.');
  if (!data.email) throw new Error('이메일을 입력해 주세요.');
  if (!data.experience) throw new Error('경력을 입력해 주세요.');
  if (!data.photo) throw new Error('사진을 업로드해 주세요.');
  if (!data.signature) throw new Error('서명을 해 주세요.');
  if (!data.agree) throw new Error('개인정보 수집·이용에 동의해 주세요.');

  const sheet = prepareMainSheet_();

  // 중복 검사: 성명 + 핸드폰 + 생년월일이 모두 같은 기존 신청이 있으면
  // 새 행을 추가하지 않고 그 행을 최신 내용으로 갱신한다.
  // (응답 유실로 인한 재제출·정보 수정 재제출을 하나의 신청으로 유지)
  const existingRow = findMatchingRow_(sheet, data.name, data.phone, data.birth);
  const isUpdate = existingRow > 0;

  let serialNo;
  if (isUpdate) {
    serialNo = sheet.getRange(existingRow, 1).getValue() || (existingRow - 1);
  } else {
    const lastRow = sheet.getLastRow();
    serialNo = lastRow < 2 ? 1 : lastRow;
  }

  // 원본 파일은 구글 드라이브에 그대로 보존
  const photoUrl = savePhoto_(data.photo, data.name);
  const signatureUrl = saveSignature_(data.signature, data.name);

  const recommender = (data.recommenderAffiliation || data.recommenderName)
    ? ((data.recommenderAffiliation || '') + (data.recommenderAffiliation && data.recommenderName ? ' / ' : '') + (data.recommenderName || ''))
    : '';

  const rowData = [
    serialNo,
    data.name,
    photoUrl,
    data.position1,
    data.position2,
    data.birth,
    data.address,
    data.phone,
    data.email,
    data.experience,
    signatureUrl,
    recommender,
    '동의',
    data.applicationDate || '',
    data.isMember ? '예' : ''
  ];

  let targetRow;
  if (isUpdate) {
    targetRow = existingRow;
    sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
    targetRow = sheet.getLastRow();
  }

  setMediaCell_(sheet, targetRow, 3, photoUrl, '사진 원본');
  setMediaCell_(sheet, targetRow, 11, signatureUrl, '서명 원본');

  return { success: true, serialNo: serialNo, updated: isUpdate };
}

/**
 * 성명 + 핸드폰 + 생년월일이 모두 일치하는 기존 신청 행 번호를 반환한다.
 * 없으면 0. 비교 시 이름의 공백/대소문자, 번호·생년월일의 구분기호는 무시한다.
 */
function findMatchingRow_(sheet, name, phone, birth) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  // 2열(성명)~8열(핸드폰) 범위를 한 번에 읽어 비교
  const values = sheet.getRange(2, 2, lastRow - 1, 7).getValues();
  const targetName = normName_(name);
  const targetPhone = normDigits_(phone);
  const targetBirth = normDigits_(birth);
  if (!targetName || !targetPhone || !targetBirth) return 0;

  for (let i = 0; i < values.length; i++) {
    const rowName = normName_(values[i][0]);   // 성명 (2열)
    const rowBirth = normDigits_(values[i][4]); // 생년월일 (6열)
    const rowPhone = normDigits_(values[i][6]); // 핸드폰 (8열)
    if (rowName === targetName && rowPhone === targetPhone && rowBirth === targetBirth) {
      return i + 2;
    }
  }
  return 0;
}

function normName_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, '').toLowerCase();
}

function normDigits_(v) {
  return String(v == null ? '' : v).replace(/[^0-9]/g, '');
}

/**
 * 2번 요구사항 반영: 출력용 명단 생성
 * - '사진 원본 보기', '서명 원본 보기' 텍스트 완전 삭제
 * - 셀 위에 떠 있는 형태가 아닌 셀 내부 실제 이미지(CellImage)로 삽입
 */
function createPrintSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEET_NAME);

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('임원신청서 시트가 존재하지 않습니다.');
    return;
  }

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('제출된 신청서 데이터가 없습니다.');
    return;
  }

  let printSheet = ss.getSheetByName(PRINT_SHEET_NAME);
  if (printSheet) {
    ss.deleteSheet(printSheet);
  }
  printSheet = ss.insertSheet(PRINT_SHEET_NAME);

  printSheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  const sourceData = sourceSheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const printData = [];
  const photoUrls = [];
  const signatureUrls = [];

  for (let i = 0; i < sourceData.length; i++) {
    const row = sourceData[i];
    const sourceRow = i + 2;

    const photoUrl = getMediaUrlFromCell_(sourceSheet.getRange(sourceRow, 3));
    const signatureUrl = getMediaUrlFromCell_(sourceSheet.getRange(sourceRow, 11));

    // 텍스트를 완전히 빈칸("")으로 처리하여 문구 제거
    printData.push([
      row[0], row[1], '', row[3], row[4], row[5], row[6],
      row[7], row[8], row[9], '', row[11], row[12], row[13], row[14]
    ]);

    photoUrls.push(photoUrl);
    signatureUrls.push(signatureUrl);
  }

  printSheet.getRange(2, 1, printData.length, HEADERS.length).setValues(printData);
  setupPrintSheetStyle_(printSheet, printData.length);

  // 셀 내 실제 이미지 삽입
  for (let i = 0; i < printData.length; i++) {
    const targetRow = i + 2;

    if (photoUrls[i]) {
      const pId = extractDriveFileId_(photoUrls[i]);
      if (pId) {
        try {
          const img = SpreadsheetApp.newCellImage()
            .setSourceUrl('https://drive.google.com/uc?export=view&id=' + pId)
            .build();
          printSheet.getRange(targetRow, 3).setValue(img);
        } catch (e) {}
      }
    }

    if (signatureUrls[i]) {
      const sId = extractDriveFileId_(signatureUrls[i]);
      if (sId) {
        try {
          const img = SpreadsheetApp.newCellImage()
            .setSourceUrl('https://drive.google.com/uc?export=view&id=' + sId)
            .build();
          printSheet.getRange(targetRow, 11).setValue(img);
        } catch (e) {}
      }
    }

    printSheet.setRowHeight(targetRow, 110);
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('출력용 명단 생성이 완료되었습니다.');
}

function prepareMainSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setBackground('#f3f3f3').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function savePhoto_(dataUrl, name) {
  const match = String(dataUrl).match(/^data:(image\/[^;]+);base64,(.+)$/);
  if (!match) throw new Error('사진 파일을 확인할 수 없습니다.');
  const contentType = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  const ext = contentType === 'image/png' ? 'png' : 'jpg';
  const blob = Utilities.newBlob(bytes, contentType, '임원신청_사진_' + name + '_' + Date.now() + '.' + ext);
  const folder = getOrCreateFolder_('PHOTO_FOLDER_ID', '서울서부본부 임원신청 사진');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function saveSignature_(dataUrl, name) {
  const match = String(dataUrl).match(/^data:image\/png;base64,(.+)$/);
  if (!match) throw new Error('서명 이미지를 확인할 수 없습니다.');
  const bytes = Utilities.base64Decode(match[1]);
  const blob = Utilities.newBlob(bytes, 'image/png', '임원신청_서명_' + name + '_' + Date.now() + '.png');
  const folder = getOrCreateFolder_('SIGNATURE_FOLDER_ID', '서울서부본부 임원신청 서명');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateFolder_(propertyName, folderName) {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty(propertyName);
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      props.deleteProperty(propertyName);
    }
  }
  const folder = DriveApp.createFolder(folderName);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty(propertyName, folder.getId());
  return folder;
}

function setMediaCell_(sheet, row, column, url, label) {
  if (!url) return;
  const style = SpreadsheetApp.newTextStyle().setForegroundColor('#1155cc').setUnderline(true).build();
  const richText = SpreadsheetApp.newRichTextValue().setText(label).setLinkUrl(url).setTextStyle(style).build();
  sheet.getRange(row, column).setRichTextValue(richText).setNote(url);
}

function getMediaUrlFromCell_(cell) {
  const note = cell.getNote();
  if (note && /^https?:\/\//i.test(note.trim())) return note.trim();
  const rich = cell.getRichTextValue();
  if (rich && rich.getLinkUrl()) return rich.getLinkUrl();
  const val = String(cell.getValue() || '').trim();
  if (/^https?:\/\//i.test(val)) return val;
  return '';
}

function setupPrintSheetStyle_(sheet, dataCount) {
  sheet.getRange(1, 1, Math.max(1, dataCount + 1), HEADERS.length)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sheet.getRange(1, 1, 1, HEADERS.length).setBackground('#f3f3f3').setFontWeight('bold');
  const widths = [55, 75, 120, 120, 120, 100, 260, 120, 180, 220, 180, 150, 180, 100, 110];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
}

function extractDriveFileId_(url) {
  if (!url) return '';
  let match = String(url).match(/\/d\/([-\w]{25,})/);
  if (match) return match[1];
  match = String(url).match(/[?&]id=([-\w]{25,})/);
  if (match) return match[1];
  match = String(url).match(/^[-\w]{25,}$/);
  if (match) return match[0];
  return '';
}
