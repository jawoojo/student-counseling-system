// =========================================================================
// SCRIPT CONFIGURATION (설정)
// =========================================================================
const CONFIG = {
  // 시트 이름 설정
  masterListSheetName: "학생목록", // 학생 목록이 있는 시트 이름
  templateSheetName: "양식",       // 학생 시트의 원본 양식 시트 이름

  // '학생목록' 시트의 열 번호 설정 (A열=1, B열=2, ...)
  col: {
    number: 1,    // 번호 열
    name: 2,      // 이름 열
    email: 3,     // 이메일 열 (이메일 발송에 사용)
    sheetLink: 4, // 학생 파일 링크가 저장될 열
    lastSync: 5,  // 마지막 동기화 시간이 기록될 열
    status: 6     // 동기화 상태/오류가 기록될 열
  },

  // 구글 드라이브 폴더 설정
  studentFolder: "학생 상담카드",

  // 고정(관리용) 시트 이름 목록 (이 시트들은 정렬에서 제외됨)
  fixedSheets: ["학생목록", "양식"]
};


// =========================================================================
// SCRIPT 1: 사용자 메뉴 생성
// =========================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('학생 데이터 관리')
    .addItem('① 학생 데이터 가져오기 (Pull)', 'pullData_Menu')
    .addItem('② 현재 시트 학생에게 보내기 (Push)', 'pushData_Menu')
    .addSeparator()
    .addItem('③ 학생 파일 링크 이메일 보내기', 'sendStudentFileLinksByEmail_Menu') // 새 메뉴 추가
    .addSeparator()
    .addItem('자동 동기화 켜기 (파일 열 때)', 'enableAutoSyncOnOpen')
    .addItem('자동 동기화 끄기 (파일 열 때)', 'disableAutoSyncOnOpen')
    .addToUi();

  ui.createMenu('🚨 초기 설정 (최초 1회)')
    .addItem('모든 학생 파일 생성 및 초기화', 'createAndShareStudentSheets')
    .addToUi();
}


// =========================================================================
// SCRIPT 2: 파일 생성 & 초기 동기화 (★★ 안정성 및 성능 강화 버전 ★★)
// =========================================================================
function createAndShareStudentSheets() {
  const ui = SpreadsheetApp.getUi();
  const CONFIRM_TEXT = "초기설정";

  const response = ui.prompt(
    '‼️매우 중요: 실행 확인‼️',
    `이 작업은 모든 학생 파일을 새로 만들고, '${CONFIG.masterListSheetName}' 시트의 링크를 덮어씁니다.\n\n기존 학생 파일이 있다면 접근 권한을 잃을 수 있습니다. 계속하려면 아래 입력창에 '${CONFIRM_TEXT}'라고 정확히 입력한 후 [확인]을 누르세요.`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText().trim() !== CONFIRM_TEXT) {
    ui.alert('입력 내용이 다르거나 취소되었습니다. 작업이 실행되지 않았습니다.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("학생 파일 생성을 시작합니다...", "초기 설정", -1);

  const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);
  if (!templateSheet) {
    ui.alert(`오류: '${CONFIG.templateSheetName}' 시트를 찾을 수 없습니다. 시트 이름을 확인해주세요.`);
    return;
  }

  const studentListSheet = ss.getSheetByName(CONFIG.masterListSheetName);
  if (!studentListSheet) {
    ui.alert(`오류: '${CONFIG.masterListSheetName}' 시트를 찾을 수 없습니다. 시트 이름을 확인해주세요.`);
    return;
  }
  
  // 상태 열 헤더 추가
  studentListSheet.getRange(1, CONFIG.col.status).setValue("상태");

  const data = studentListSheet.getDataRange().getValues();
  
  // 드라이브 폴더 가져오기 또는 생성
  let targetFolder;
  const folders = DriveApp.getFoldersByName(CONFIG.studentFolder);
  if (folders.hasNext()) {
    targetFolder = folders.next();
  } else {
    targetFolder = DriveApp.createFolder(CONFIG.studentFolder);
  }

  // 루프 시작
  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const studentName = data[i][CONFIG.col.name - 1];
    
    // 학생 이름이 없으면 건너뛰기
    if (!studentName) continue;
    
    ss.toast(`${studentName} 학생 파일 생성 중...`, "초기 설정", 10);
    
    try {
      // 1. 새 스프레드시트 생성
      const newSpreadsheet = SpreadsheetApp.create(studentName);
      const newFile = DriveApp.getFileById(newSpreadsheet.getId());

      // 2. 지정된 폴더로 이동
      targetFolder.addFile(newFile);
      DriveApp.getRootFolder().removeFile(newFile);
      
      // 3. 양식 시트 복사 및 이름 변경
      const copiedSheet = templateSheet.copyTo(newSpreadsheet);
      newSpreadsheet.deleteSheet(newSpreadsheet.getSheets()[0]); // 기본 'Sheet1' 삭제
      copiedSheet.setName(studentName); // 복사된 시트 이름을 학생 이름으로 변경

      // 4. 공유 설정 (링크가 있는 모든 사용자가 수정 가능)
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      
      // 5. 마스터 시트에 링크와 상태 기록
      const sheetUrl = newSpreadsheet.getUrl();
      studentListSheet.getRange(row, CONFIG.col.sheetLink).setValue(sheetUrl);
      studentListSheet.getRange(row, CONFIG.col.status).setValue("파일 생성 완료");

    } catch (e) {
      // 오류 발생 시 마스터 시트에 상태 기록
      const errorMessage = `오류: ${e.message.substring(0, 100)}`;
      studentListSheet.getRange(row, CONFIG.col.status).setValue(errorMessage);
      console.error(`학생 파일 생성 오류 (${studentName}): ${e.toString()}`);
    }
  }

  ss.toast("자동 동기화 트리거를 설정합니다...", "초기 설정", -1);
  enableOnEditTrigger(); // 교사->학생 실시간 동기화 트리거
  enableAutoSyncOnOpen(true); // 학생->교사 자동 동기화 트리거 (파일 열 때)

  ss.toast("첫 데이터 동기화를 시작합니다. 모든 학생 탭이 생성됩니다...", "초기 설정", -1);
  pullDataCore(true); // isInitialSync = true

  ss.toast("학생 시트 순서를 정렬합니다...", "초기 설정", -1);
  sortStudentSheets();

  ui.alert(`초기 설정 완료!\n\n학생 파일 및 시트 생성이 완료되었으며, 번호 순으로 정렬되었습니다.\n\n이제 파일을 열 때마다 데이터가 자동으로 동기화됩니다.`);
}
// =========================================================================
// SCRIPT 3: 교사 -> 학생 동기화 (Push) (★★ 데이터 충돌 시 자동 Pull 후 Push ★★)
// =========================================================================
function onEdit(e) {
  const editedRange = e.range;
  const editedSheet = editedRange.getSheet();
  const editedSheetName = editedSheet.getName();

  // '학생목록'이거나 고정 시트이면 동기화 안함
  if (editedSheetName === CONFIG.masterListSheetName || CONFIG.fixedSheets.includes(editedSheetName)) {
    return;
  }
  
  // 여러 셀을 동시에 수정하면(예: 붙여넣기) 부하가 크므로 동기화 중단
  if (editedRange.getNumRows() > 5 || editedRange.getNumColumns() > 5) {
      SpreadsheetApp.getActiveSpreadsheet().toast("많은 셀이 변경되어 자동 Push 동기화는 실행되지 않습니다. [수동 보내기] 메뉴를 이용하세요.", "알림", 5);
      return;
  }

  // 실제 로직은 코어 함수에 위임. onEdit 이벤트 객체(e)를 전달
  pushDataCore(editedSheetName, false, e); 
}

function pushData_Menu() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();

  if (CONFIG.fixedSheets.includes(sheetName)) {
    SpreadsheetApp.getUi().alert(`'${sheetName}' 시트는 학생에게 보낼 수 없습니다.`);
    return;
  }
  
  const result = pushDataCore(sheetName, true, null); // isManual = true, event = null
  if (result) {
      SpreadsheetApp.getUi().alert(`'${sheetName}' 학생에게 현재 시트 내용을 성공적으로 보냈습니다.`);
  }
}

/**
 * 교사 -> 학생 데이터 동기화 (Push) 핵심 로직
 * @param {string} studentName 동기화할 학생 시트 이름
 * @param {boolean} isManual 수동 메뉴로 실행되었는지 여부
 * @param {Object} e onEdit 이벤트 객체 (선택사항)
 */
function pushDataCore(studentName, isManual = false, e = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const masterListSheet = ss.getSheetByName(CONFIG.masterListSheetName);
  if (!masterListSheet) return false;

  const data = masterListSheet.getDataRange().getValues();
  let studentInfo = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][CONFIG.col.name - 1] === studentName) {
      studentInfo = {
        row: i + 1,
        name: data[i][CONFIG.col.name - 1],
        url: data[i][CONFIG.col.sheetLink - 1],
        lastSync: data[i][CONFIG.col.lastSync - 1] ? new Date(data[i][CONFIG.col.lastSync - 1]) : null
      };
      break;
    }
  }
  
  if (!studentInfo || !studentInfo.url) {
    if(isManual) ui.alert(`오류: '${studentName}' 학생의 파일 링크를 '${CONFIG.masterListSheetName}' 시트에서 찾을 수 없습니다.`);
    return false;
  }

  try {
    const studentFileId = studentInfo.url.match(/[-\w]{25,}/);
    if (!studentFileId) throw new Error("유효하지 않은 파일 링크입니다.");

    const studentFile = DriveApp.getFileById(studentFileId[0]);
    const studentFileLastUpdated = studentFile.getLastUpdated();

    // ★★★★★ 학생 데이터 우선 충돌 해결 로직 ★★★★★
    if (studentInfo.lastSync && studentFileLastUpdated.getTime() > studentInfo.lastSync.getTime() + 1000) {
      
      ss.toast(`'${studentName}' 학생의 최신 데이터와 병합합니다...`, "자동 병합", 5);
      
      // 1. 학생의 최신 데이터를 자동으로 가져오기 (Pull)
      //    이 시점에서 교사의 수정사항은 학생 데이터로 덮어씌워집니다.
      pullDataCore(true, [studentName]);

      // 2. 교사의 변경사항을 다시 적용하는 로직을 '제거'하여 학생 데이터를 보존합니다.
      
      // 3. 사용자에게 결과를 알립니다.
      if (e && e.range) {
        const cellAddress = e.range.getA1Notation();
        ss.toast(`병합 완료: 학생이 먼저 수정한 내용이 '${cellAddress}' 셀에 유지되었습니다.`, "알림", 10);
      } else {
        ss.toast("병합 완료: 학생의 최신 내용이 반영되었습니다.", "알림", 10);
      }
      
      // 이후 로직은 자연스럽게 '학생 우선'으로 병합된 데이터를 학생 파일에 Push하게 됩니다.
    }

    // --- 데이터 Push 로직 ---
    const sourceSheet = ss.getSheetByName(studentName);
    const sourceData = sourceSheet.getDataRange().getValues();

    const studentSpreadsheet = SpreadsheetApp.openById(studentFileId[0]);
    const destinationSheet = studentSpreadsheet.getSheetByName(studentName) || studentSpreadsheet.getSheets()[0];
    
    destinationSheet.clearContents();
    destinationSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

    const syncTime = new Date();
    masterListSheet.getRange(studentInfo.row, CONFIG.col.lastSync).setValue(syncTime);
    masterListSheet.getRange(studentInfo.row, CONFIG.col.status).setValue(`Push 완료 (${syncTime.toLocaleTimeString()})`);
    return true;

  } catch (err) {
    const errorMessage = `Push 오류: ${err.message}`;
    masterListSheet.getRange(studentInfo.row, CONFIG.col.status).setValue(errorMessage);
    if(isManual) ui.alert(errorMessage);
    console.error(`Push 동기화 오류 (${studentName}): ${err.toString()}`);
    return false;
  }
}
// =========================================================================
// SCRIPT 4: 학생 -> 교사 동기화 (Pull) (★★ 성능 최적화 버전 ★★)
// =========================================================================
function pullData_Menu() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('알림', '변경된 학생의 데이터만 불러옵니다. 계속하시겠습니까?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) return;

  SpreadsheetApp.getActiveSpreadsheet().toast("학생 데이터 확인을 시작합니다...", "Pull 동기화", -1);
  const updatedCount = pullDataCore(false); // isInitialSync = false

  if (updatedCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("시트 순서를 정렬합니다...", "상태", 5);
    sortStudentSheets();
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`총 ${updatedCount}명의 학생 데이터가 업데이트되었습니다.`, "작업 완료", 10);
  if (updatedCount > 0) {
    ui.alert(`작업 완료: 총 ${updatedCount}명의 학생 데이터가 업데이트 및 정렬되었습니다.`);
  } else {
    ui.alert(`작업 완료: 업데이트할 새로운 데이터가 없습니다.`);
  }
}

function autoPullDataOnOpen() {
  const updatedCount = pullDataCore(false); // isInitialSync = false
  if (updatedCount > 0) {
    sortStudentSheets();
    SpreadsheetApp.getActiveSpreadsheet().toast(`자동 동기화 완료: 총 ${updatedCount}명 업데이트 및 정렬됨`, "완료", 10);
  }
}

/**
 * 학생 -> 교사 데이터 동기화 (Pull) 핵심 로직
 * @param {boolean} isSilent 사용자에게 알림(toast)을 최소화할지 여부
 * @param {Array<string>} [studentsToPull=null] 특정 학생의 이름 배열. null이면 전체 학생을 대상으로 함.
 */
function pullDataCore(isSilent = false, studentsToPull = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterListSheet = ss.getSheetByName(CONFIG.masterListSheetName);
  if (!masterListSheet) return 0;

  const data = masterListSheet.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const studentName = data[i][CONFIG.col.name - 1];
    
    // 특정 학생만 동기화하도록 지정된 경우, 해당 학생이 아니면 건너뜀
    if (studentsToPull && !studentsToPull.includes(studentName)) {
      continue;
    }

    const studentFileUrl = data[i][CONFIG.col.sheetLink - 1];
    const lastSyncTime = data[i][CONFIG.col.lastSync - 1] ? new Date(data[i][CONFIG.col.lastSync - 1]) : null;

    if (!studentName || !studentFileUrl) continue;

    if (!isSilent) masterListSheet.getRange(row, CONFIG.col.status).setValue("확인 중...");

    try {
      const studentFileId = studentFileUrl.match(/[-\w]{25,}/);
      if (!studentFileId) {
          masterListSheet.getRange(row, CONFIG.col.status).setValue("오류: 유효하지 않은 링크");
          continue;
      };

      const studentFile = DriveApp.getFileById(studentFileId[0]);
      const lastUpdatedTime = studentFile.getLastUpdated();

      // 시간 비교를 통한 업데이트 여부 결정
      let shouldUpdate = false;
      // studentsToPull이 지정된 경우는 강제 업데이트
      if (studentsToPull || !lastSyncTime || lastUpdatedTime.getTime() > lastSyncTime.getTime() + 1000) {
        shouldUpdate = true;
      }
      
      if (shouldUpdate) {
        if (!isSilent) ss.toast(`${studentName} 데이터 가져오는 중...`, "Pull 동기화", 10);
        
        const studentSpreadsheet = SpreadsheetApp.openById(studentFileId[0]);
        const sourceSheet = studentSpreadsheet.getSheetByName(studentName) || studentSpreadsheet.getSheets()[0];
        const sourceData = sourceSheet.getDataRange().getValues();

        let destinationSheet = ss.getSheetByName(studentName);

        if (destinationSheet) {
          destinationSheet.clearContents(); 
        } else {
          const templateSheet = ss.getSheetByName(CONFIG.templateSheetName);
          destinationSheet = templateSheet.copyTo(ss).setName(studentName);
        }
        
        destinationSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

        masterListSheet.getRange(row, CONFIG.col.lastSync).setValue(lastUpdatedTime);
        if (!isSilent) masterListSheet.getRange(row, CONFIG.col.status).setValue(`Pull 완료 (${lastUpdatedTime.toLocaleTimeString()})`);
        updatedCount++;
      } else {
          if (!isSilent) masterListSheet.getRange(row, CONFIG.col.status).setValue("최신 상태");
      }

    } catch (e) {
      const errorMessage = `Pull 오류: ${e.message.substring(0, 100)}`;
      masterListSheet.getRange(row, CONFIG.col.status).setValue(errorMessage);
      console.error(`Pull 동기화 오류 (${studentName}): ${e.toString()}`);
    }
  }
  return updatedCount;
}

// =========================================================================
// SCRIPT 5: 트리거 및 유틸리티 함수
// =========================================================================
/**
 * 학생 시트를 '학생목록'의 번호 순서대로 정렬합니다.
 */
function sortStudentSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterListSheet = ss.getSheetByName(CONFIG.masterListSheetName);
  if (!masterListSheet) return;

  // 1. '학생목록'에서 번호와 이름을 가져와 정렬 기준 생성
  const studentInfoList = masterListSheet.getRange(2, CONFIG.col.number, masterListSheet.getLastRow() - 1, CONFIG.col.name)
    .getValues()
    .map(row => ({ number: row[CONFIG.col.number - 1], name: row[CONFIG.col.name - 1] }))
    .filter(info => info.name && !isNaN(info.number))
    .sort((a, b) => a.number - b.number);

  // 2. 모든 시트를 맵으로 만들어 빠른 조회 준비
  const allSheets = ss.getSheets();
  const sheetMap = {};
  allSheets.forEach(sheet => {
    sheetMap[sheet.getName()] = sheet;
  });

  // 3. 고정 시트를 제외한 시트가 시작될 위치 계산
  let targetPosition = 1;
  allSheets.forEach(sheet => {
    if (CONFIG.fixedSheets.includes(sheet.getName())) {
      targetPosition++;
    }
  });

  // 4. 정렬된 목록 순서대로 시트 이동
  studentInfoList.forEach(studentInfo => {
    const sheetToSort = sheetMap[studentInfo.name];
    if (sheetToSort && sheetToSort.getIndex() !== targetPosition) {
      ss.setActiveSheet(sheetToSort);
      ss.moveActiveSheet(targetPosition);
    }
    // 시트가 존재할 경우에만 다음 위치로 이동
    if (sheetToSort) {
        targetPosition++;
    }
  });
}

/**
 * 프로젝트의 모든 트리거를 삭제합니다. (초기화용)
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}

/**
 * 교사->학생 동기화를 위한 onEdit 트리거를 생성합니다.
 */
function enableOnEditTrigger() {
  deleteAllTriggers(); // 기존 트리거를 모두 지우고 새로 설정
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

/**
 * 파일 열 때 학생->교사 동기화를 위한 onOpen 트리거를 생성합니다.
 */
function enableAutoSyncOnOpen(isSilent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 이미 동일한 트리거가 있는지 확인
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getEventType() === ScriptApp.EventType.ON_OPEN && trigger.getHandlerFunction() === "autoPullDataOnOpen") {
      if (!isSilent) SpreadsheetApp.getUi().alert("자동 동기화 기능이 이미 켜져 있습니다.");
      return;
    }
  }
  // 트리거 생성
  ScriptApp.newTrigger("autoPullDataOnOpen")
    .forSpreadsheet(ss)
    .onOpen()
    .create();
  if (!isSilent) SpreadsheetApp.getUi().alert("설정 완료! 이제부터 파일을 열 때마다 변경된 학생 데이터가 자동으로 동기화됩니다.");
}

/**
 * 파일 열 때 자동 동기화 트리거를 삭제합니다.
 */
function disableAutoSyncOnOpen() {
  const triggers = ScriptApp.getProjectTriggers();
  let wasTriggerDeleted = false;
  for (const trigger of triggers) {
    if (trigger.getEventType() === ScriptApp.EventType.ON_OPEN && trigger.getHandlerFunction() === "autoPullDataOnOpen") {
      ScriptApp.deleteTrigger(trigger);
      wasTriggerDeleted = true;
    }
  }
  if (wasTriggerDeleted) {
    SpreadsheetApp.getUi().alert("자동 동기화 기능이 꺼졌습니다.");
  } else {
    SpreadsheetApp.getUi().alert("현재 설정된 자동 동기화 기능이 없습니다.");
  }
}

// =========================================================================
// SCRIPT 6: 학생 파일 링크 이메일 발송 기능 (새로 추가)
// =========================================================================
/**
 * '학생목록' 시트에 있는 이메일 주소로 학생 개개인의 상담 카드 링크를 발송합니다.
 */
function sendStudentFileLinksByEmail_Menu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterListSheet = ss.getSheetByName(CONFIG.masterListSheetName);

  if (!masterListSheet) {
    ui.alert(`오류: '${CONFIG.masterListSheetName}' 시트를 찾을 수 없습니다. 시트 이름을 확인해주세요.`);
    return;
  }

  const response = ui.alert(
    '이메일 발송 확인',
    `'${CONFIG.masterListSheetName}' 시트에 있는 모든 학생의 이메일 주소로 상담 카드 링크를 발송하시겠습니까?`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('이메일 발송이 취소되었습니다.');
    return;
  }

  ss.toast("학생 파일 링크 이메일 발송을 시작합니다...", "이메일 발송", -1);

  const data = masterListSheet.getDataRange().getValues();
  let sentCount = 0;
  let failCount = 0;

  // 헤더 행을 건너뛰고 1번째 행부터 시작
  for (let i = 1; i < data.length; i++) {
    const row = i + 1; // 실제 시트 행 번호 (1부터 시작)
    const studentName = data[i][CONFIG.col.name - 1];
    const studentEmail = data[i][CONFIG.col.email - 1];
    const studentLink = data[i][CONFIG.col.sheetLink - 1];

    // 이름, 이메일, 링크 중 하나라도 없으면 건너뛰고 상태 업데이트
    if (!studentName || !studentEmail || !studentLink || !isValidEmail(studentEmail)) {
      let statusMessage = "이메일 발송 실패: ";
      if (!studentName) statusMessage += "이름 없음. ";
      if (!studentEmail) statusMessage += "이메일 없음. ";
      else if (!isValidEmail(studentEmail)) statusMessage += "유효하지 않은 이메일 주소. ";
      if (!studentLink) statusMessage += "링크 없음. ";
      masterListSheet.getRange(row, CONFIG.col.status).setValue(statusMessage.trim());
      failCount++;
      continue;
    }

    ss.toast(`${studentName} 학생에게 이메일 발송 중...`, "이메일 발송", 5);

    try {
      const subject = `${studentName} 학생 상담 카드 링크입니다.`;
      const body = `안녕하세요, ${studentName} 학생.\n\n` +
                   `상담 카드 링크를 보내드립니다. 이 링크를 통해 상담 내용을 확인하고 필요에 따라 수정할 수 있습니다.\n\n` +
                   `링크: ${studentLink}\n\n` +
                   `궁금한 점이 있다면 언제든지 문의해주세요.\n\n` +
                   `감사합니다.`;

      GmailApp.sendEmail(studentEmail, subject, body);
      masterListSheet.getRange(row, CONFIG.col.status).setValue(`이메일 발송 완료 (${new Date().toLocaleTimeString()})`);
      sentCount++;
    } catch (e) {
      const errorMessage = `이메일 발송 실패: ${e.message.substring(0, 100)}`;
      masterListSheet.getRange(row, CONFIG.col.status).setValue(errorMessage);
      console.error(`이메일 발송 오류 (${studentName}, ${studentEmail}): ${e.toString()}`);
      failCount++;
    }
  }

  ss.toast("이메일 발송 작업 완료!", "이메일 발송", 10);
  ui.alert(`이메일 발송 완료!\n\n총 ${sentCount}명에게 이메일을 성공적으로 보냈습니다.\n실패: ${failCount}명.`);
}

/**
 * 이메일 주소의 유효성을 간단히 검사합니다.
 * @param {string} email 이메일 주소
 * @returns {boolean} 유효하면 true, 그렇지 않으면 false
 */
function isValidEmail(email) {
  // 간단한 이메일 정규식 검사
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
