// =========================================================================
// SCRIPT CONFIGURATION (★★ 여기만 수정하세요! ★★)
// =========================================================================
const CONFIG = {
  /**
   * 1. 용어 설정 (이 스크립트 전체에서 사용될 용어를 정의합니다)
   * 예: 학생 상담록 -> item: "학생", listSheet: "학생 명단"
   * 예: 프로젝트 관리 -> item: "프로젝트", listSheet: "프로젝트 목록"
   */
  terms: {
    item: "항목", // 개별 데이터 단위의 명칭 (예: 학생, 프로젝트, 직원 등)
    itemFile: "개별 파일", // 각 항목에 대해 생성되는 파일의 명칭 (예: 상담 카드, 프로젝트 시트 등)
    listSheet: "데이터 목록" // 전체 항목의 정보가 담긴 시트의 명칭
  },

  /**
   * 2. 시트 이름 설정
   */
  sheetNames: {
    masterList: "데이터 목록", // 전체 항목 정보가 있는 시트 이름 (위 terms.listSheet와 일치시키는 것을 권장)
    template: "양식"        // 개별 파일의 원본이 될 시트 이름
  },

  /**
   * 3. '데이터 목록' 시트의 열 번호 설정 (A=1, B=2, ...)
   */
  cols: {
    id: 1,        // 고유 ID 또는 번호 열
    itemName: 2,  // 항목 이름 열 (이 이름으로 시트와 파일이 생성됨)
    email: 3,     // 이메일 수신자 주소 열
    fileLink: 4,  // 생성된 개별 파일 링크가 저장될 열
    lastSync: 5,  // 마지막 동기화 시간이 기록될 열
    status: 6     // 동기화 상태/오류가 기록될 열
  },

  /**
   * 4. 구글 드라이브 폴더 설정
   */
  folder: {
    // 개별 파일들이 저장될 폴더 이름입니다.
    // 이 폴더는 현재 스프레드시트가 있는 위치에 자동으로 생성됩니다.
    name: "배포된 개별 파일"
  },
  t
  /**
   * 5. 이메일 템플릿 설정
   * {{itemName}} -> 항목 이름으로 자동 대체됩니다.
   * {{itemLink}} -> 개별 파일 링크로 자동 대체됩니다.
   */
  email: {
    subject: "{{itemName}} 님의 개별 파일 링크입니다.",
    body: "안녕하세요, {{itemName}} 님.\n\n" +
          "요청하신 {{itemFile}} 링크를 보내드립니다. 이 링크를 통해 내용을 확인하고 필요에 따라 수정할 수 있습니다.\n\n" +
          "링크: {{itemLink}}\n\n" +
          "궁금한 점이 있다면 언제든지 문의해주세요.\n\n" +
          "감사합니다."
  },
  
  /**
   * 6. 기타 고급 설정
   */
  settings: {
    // 자동 동기화(Push)를 제한할 최대 셀 변경 개수 (행*열). 대용량 붙여넣기 시 서버 부하 방지용.
    syncCellLimit: 50, 
    // 정렬에서 제외할 고정 시트 이름 목록 (위에서 설정한 시트 이름들을 자동으로 포함)
    get fixedSheets() {
      return [CONFIG.sheetNames.masterList, CONFIG.sheetNames.template];
    }
  }
};

// =========================================================================
// SCRIPT 1: 사용자 메뉴 생성
// =========================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const itemTerm = CONFIG.terms.item; // "항목"

  ui.createMenu(`📄 ${itemTerm} 데이터 관리`)
    .addItem(`① ${itemTerm} 데이터 가져오기 (Pull)`, 'pullData_Menu')
    .addItem(`② 현재 시트 ${itemTerm}에게 보내기 (Push)`, 'pushData_Menu')
    .addSeparator()
    .addItem(`③ ${itemTerm} 파일 링크 이메일 보내기`, 'sendItemFileLinksByEmail_Menu')
    .addItem('④ 양식 업데이트', 'updateSheetPortion_Menu') // <<<==== 이 메뉴 추가
    .addSeparator()
    .addItem('자동 동기화 켜기 (파일 열 때)', 'enableAutoSyncOnOpen')
    .addItem('자동 동기화 끄기 (파일 열 때)', 'disableAutoSyncOnOpen')
    .addToUi();

  ui.createMenu('🚨 초기 설정 (최초 1회)')
    .addItem(`모든 ${itemTerm} 파일 생성 및 초기화`, 'createAndShareItemFiles')
    .addSeparator()
    .addItem('made by 상당고 장운종, 백수연', 'showMadeByInfo')
    .addToUi();
}

// =========================================================================
// SCRIPT 2: 파일 생성 & 초기 동기화 (★★ 상대 경로 및 범용성 강화 버전 ★★)
// =========================================================================
function createAndShareItemFiles() {
  const ui = SpreadsheetApp.getUi();
  const CONFIRM_TEXT = "초기설정";

  const response = ui.prompt(
    '‼️매우 중요: 실행 확인‼️',
    `이 작업은 모든 ${CONFIG.terms.item} 파일을 새로 만들고, '${CONFIG.sheetNames.masterList}' 시트의 링크를 덮어씁니다.\n\n기존 파일이 있다면 접근 권한을 잃을 수 있습니다. 계속하려면 아래 입력창에 '${CONFIRM_TEXT}'라고 정확히 입력한 후 [확인]을 누르세요.`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText().trim() !== CONFIRM_TEXT) {
    ui.alert('입력 내용이 다르거나 취소되었습니다. 작업이 실행되지 않았습니다.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast(`${CONFIG.terms.item} 파일 생성을 시작합니다...`, "초기 설정", -1);

  const templateSheet = ss.getSheetByName(CONFIG.sheetNames.template);
  if (!templateSheet) {
    ui.alert(`오류: '${CONFIG.sheetNames.template}' 시트를 찾을 수 없습니다. 시트 이름을 확인해주세요.`);
    return;
  }

  const masterListSheet = ss.getSheetByName(CONFIG.sheetNames.masterList);
  if (!masterListSheet) {
    ui.alert(`오류: '${CONFIG.sheetNames.masterList}' 시트를 찾을 수 없습니다. 시트 이름을 확인해주세요.`);
    return;
  }
  
  masterListSheet.getRange(1, CONFIG.cols.status).setValue("상태");

  const data = masterListSheet.getDataRange().getValues();
  
  // --- ★★ 파일 생성 위치(상대 경로) 로직 ★★ ---
  let targetFolder;
  const ssFile = DriveApp.getFileById(ss.getId());
  const parentFolders = ssFile.getParents();
  
  // 마스터 스프레드시트가 있는 첫 번째 부모 폴더를 가져옴
  if (parentFolders.hasNext()) {
    const parentFolder = parentFolders.next();
    const folders = parentFolder.getFoldersByName(CONFIG.folder.name);
    if (folders.hasNext()) {
      targetFolder = folders.next();
    } else {
      targetFolder = parentFolder.createFolder(CONFIG.folder.name);
    }
  } else {
    // 만약 파일이 최상위(My Drive)에 있다면, 최상위에 폴더 생성
    const folders = DriveApp.getFoldersByName(CONFIG.folder.name);
    if (folders.hasNext()) {
      targetFolder = folders.next();
    } else {
      targetFolder = DriveApp.createFolder(CONFIG.folder.name);
    }
  }
  // --- 로직 종료 ---

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const itemName = data[i][CONFIG.cols.itemName - 1];
    
    if (!itemName) continue;
    
    ss.toast(`${itemName} ${CONFIG.terms.itemFile} 생성 중...`, "초기 설정", 10);
    
    try {
      const newSpreadsheet = SpreadsheetApp.create(itemName);
      const newFile = DriveApp.getFileById(newSpreadsheet.getId());

      targetFolder.addFile(newFile);
      DriveApp.getRootFolder().removeFile(newFile);
      
      const copiedSheet = templateSheet.copyTo(newSpreadsheet);
      newSpreadsheet.deleteSheet(newSpreadsheet.getSheets()[0]);
      copiedSheet.setName(itemName);

      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      
      const sheetUrl = newSpreadsheet.getUrl();
      masterListSheet.getRange(row, CONFIG.cols.fileLink).setValue(sheetUrl);
      masterListSheet.getRange(row, CONFIG.cols.status).setValue("파일 생성 완료");

    } catch (e) {
      const errorMessage = `오류: ${e.message.substring(0, 100)}`;
      masterListSheet.getRange(row, CONFIG.cols.status).setValue(errorMessage);
      console.error(`${CONFIG.terms.item} 파일 생성 오류 (${itemName}): ${e.toString()}`);
    }
  }

  ss.toast("자동 동기화 트리거를 설정합니다...", "초기 설정", -1);
  enableOnEditTrigger();
  enableAutoSyncOnOpen(true);

  ss.toast("첫 데이터 동기화를 시작합니다. 모든 탭이 생성됩니다...", "초기 설정", -1);
  pullDataCore(true);

  ss.toast(`${CONFIG.terms.item} 시트 순서를 정렬합니다...`, "초기 설정", -1);
  sortItemSheets();

  ui.alert(`초기 설정 완료!\n\n${CONFIG.terms.item} 파일 및 시트 생성이 완료되었으며, 번호 순으로 정렬되었습니다.\n\n이제 파일을 열 때마다 데이터가 자동으로 동기화됩니다.`);
}

// =========================================================================
// SCRIPT 3: 관리자 -> 개별 파일 동기화 (Push)
// =========================================================================
function onEdit(e) {
  const editedRange = e.range;
  const editedSheet = editedRange.getSheet();
  const editedSheetName = editedSheet.getName();

  if (CONFIG.settings.fixedSheets.includes(editedSheetName)) {
    return;
  }
  
  // --- ▼▼▼ 요청하신 내용으로 수정된 부분 ▼▼▼ ---
  // 대량 데이터 변경을 감지하면
  if (editedRange.getNumRows() * editedRange.getNumColumns() > CONFIG.settings.syncCellLimit) {
      // 사용자에게 잠시 멈출 수 있음을 알림
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `많은 셀이 변경되어 자동으로 동기화를 시작합니다. 잠시 멈춤 현상이 발생할 수 있습니다.`, 
        "⏳ 자동 동기화 진행 중", 
        7 // 메시지 표시 시간 (7초)
      );
      
      // 메뉴 클릭과 동일한 효과를 내기 위해 pushData_Menu 함수를 직접 호출
      // ※ 주의: 이 작업은 30초의 실행 시간 제한을 받습니다.
      //         시간 초과 시 동기화가 중간에 멈출 수 있습니다.
      pushData_Menu_fromOnEdit(editedSheetName); // 별도의 onEdit용 함수 호출
      
      return; // 작업 완료 후 onEdit 함수 종료
  }
  // --- ▲▲▲ 여기까지 수정 ▲▲▲ ---

  pushDataCore(editedSheetName, false, e); 
}

/**
 * [신규] onEdit에서만 호출하기 위한 pushData_Menu 래퍼(wrapper) 함수.
 * pushData_Menu는 UI(알림창)를 포함하므로 onEdit에서 직접 호출 시 권한 오류가 발생할 수 있어
 * UI 부분을 제외하고 핵심 로직만 실행하도록 분리합니다.
 * @param {string} sheetName 동기화할 시트 이름
 */
function pushData_Menu_fromOnEdit(sheetName) {
  // 고정 시트인지 다시 한번 확인 (이중 안전장치)
  if (CONFIG.settings.fixedSheets.includes(sheetName)) {
    // onEdit에서는 alert를 사용할 수 없으므로 console.log로 기록
    console.warn(`'${sheetName}' 시트는 ${CONFIG.terms.item}에게 보낼 수 없습니다. (onEdit 호출)`);
    return;
  }
  
  // isManual=false로 설정하여, 완료 시 alert 창이 아닌 toast 메시지가 뜨도록 함
  const result = pushDataCore(sheetName, false, null); 
  
  if (result) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`[${sheetName}] 자동 동기화가 완료되었습니다.`, "✅ 완료", 5);
  } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(`[${sheetName}] 자동 동기화 중 오류가 발생했습니다. 상태 열을 확인하세요.`, "❌ 오류", 10);
  }
}

function pushData_Menu() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();

  if (CONFIG.settings.fixedSheets.includes(sheetName)) {
    SpreadsheetApp.getUi().alert(`'${sheetName}' 시트는 ${CONFIG.terms.item}에게 보낼 수 없습니다.`);
    return;
  }
  
  const result = pushDataCore(sheetName, true, null);
  if (result) {
      SpreadsheetApp.getUi().alert(`'${sheetName}' ${CONFIG.terms.item}에게 현재 시트 내용을 성공적으로 보냈습니다.`);
  }
}
/**
 * 교사 -> 학생 데이터 동기화 (Push) 핵심 로직 (★★ 불필요한 알림(toast) 제거 버전 ★★)
 */
function pushDataCore(itemName, isManual = false, e = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const masterListSheet = ss.getSheetByName(CONFIG.sheetNames.masterList);
  if (!masterListSheet) return false;

  const data = masterListSheet.getDataRange().getValues();
  let itemInfo = null;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][CONFIG.cols.itemName - 1]) === String(itemName)) {
      itemInfo = {
        row: i + 1,
        name: data[i][CONFIG.cols.itemName - 1],
        url: data[i][CONFIG.cols.fileLink - 1],
        lastSync: data[i][CONFIG.cols.lastSync - 1] ? new Date(data[i][CONFIG.cols.lastSync - 1]) : null
      };
      break;
    }
  }
  
  if (!itemInfo || !itemInfo.url) {
    // 수동 실행 시에만 오류 알림창 표시
    if (isManual) ui.alert(`오류: '${itemName}' ${CONFIG.terms.item}의 파일 링크를 '${CONFIG.sheetNames.masterList}' 시트에서 찾을 수 없습니다.`);
    return false;
  }

  try {
    const itemFileId = itemInfo.url.match(/[-\w]{25,}/);
    if (!itemFileId) throw new Error("유효하지 않은 파일 링크입니다.");

    const itemFile = DriveApp.getFileById(itemFileId[0]);
    const itemFileLastUpdated = itemFile.getLastUpdated();

    // ★ 데이터 충돌 해결 로직 (이 부분의 알림은 중요하므로 유지) ★
    if (itemInfo.lastSync && itemFileLastUpdated.getTime() > itemInfo.lastSync.getTime() + 1000) {
      ss.toast(`'${itemName}' ${CONFIG.terms.item}의 최신 데이터와 병합합니다...`, "자동 병합", 5);
      pullDataCore(true, [itemInfo.name]);

      if (e && e.range) {
        const cellAddress = e.range.getA1Notation();
        ss.toast(`병합 완료: ${CONFIG.terms.item}이 먼저 수정한 내용이 '${cellAddress}' 셀에 유지되었습니다.`, "알림", 10);
      } else {
        ss.toast(`병합 완료: ${CONFIG.terms.item}의 최신 내용이 반영되었습니다.`, "알림", 10);
      }
    }

    const sourceSheet = ss.getSheetByName(String(itemName));
    const sourceData = sourceSheet.getDataRange().getValues();

    const itemSpreadsheet = SpreadsheetApp.openById(itemFileId[0]);
    const destinationSheet = itemSpreadsheet.getSheetByName(String(itemInfo.name)) || itemSpreadsheet.getSheets()[0];
    
    destinationSheet.clearContents();
    destinationSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

    const syncTime = new Date();
    masterListSheet.getRange(itemInfo.row, CONFIG.cols.lastSync).setValue(syncTime);
    masterListSheet.getRange(itemInfo.row, CONFIG.cols.status).setValue(`Push 완료 (${syncTime.toLocaleTimeString()})`);
    
    // isManual 플래그가 true일 경우에만 성공 알림을 띄우는 로직은 pushData_Menu 함수에 이미 있으므로 여기서 별도 처리가 필요 없음
    
    return true;

  } catch (err) {
    const errorMessage = `Push 오류: ${err.message}`;
    if (itemInfo && itemInfo.row) {
        masterListSheet.getRange(itemInfo.row, CONFIG.cols.status).setValue(errorMessage);
    }
    // 수동 실행 시에만 오류 알림창 표시
    if(isManual) ui.alert(errorMessage);
    console.error(`Push 동기화 오류 (${itemName}): ${err.toString()}`);
    return false;
  }
}

// =========================================================================
// SCRIPT 4: 개별 파일 -> 관리자 동기화 (Pull)
// =========================================================================
function pullData_Menu() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('알림', `변경된 ${CONFIG.terms.item}의 데이터만 불러옵니다. 계속하시겠습니까?`, ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) return;

  SpreadsheetApp.getActiveSpreadsheet().toast(`${CONFIG.terms.item} 데이터 확인을 시작합니다...`, "Pull 동기화", -1);
  const updatedCount = pullDataCore(false);

  if (updatedCount > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("시트 순서를 정렬합니다...", "상태", 5);
    sortItemSheets();
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`총 ${updatedCount}개의 ${CONFIG.terms.item} 데이터가 업데이트되었습니다.`, "작업 완료", 10);
  if (updatedCount > 0) {
    ui.alert(`작업 완료: 총 ${updatedCount}개의 ${CONFIG.terms.item} 데이터가 업데이트 및 정렬되었습니다.`);
  } else {
    ui.alert(`작업 완료: 업데이트할 새로운 데이터가 없습니다.`);
  }
}

function autoPullDataOnOpen() {
  const updatedCount = pullDataCore(false);
  if (updatedCount > 0) {
    sortItemSheets();
    SpreadsheetApp.getActiveSpreadsheet().toast(`자동 동기화 완료: 총 ${updatedCount}개 업데이트 및 정렬됨`, "완료", 10);
  }
}

function pullDataCore(isSilent = false, itemsToPull = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterListSheet = ss.getSheetByName(CONFIG.sheetNames.masterList);
  if (!masterListSheet) return 0;

  const data = masterListSheet.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const itemName = data[i][CONFIG.cols.itemName - 1];
    
    if (itemsToPull && !itemsToPull.includes(itemName)) {
      continue;
    }

    const itemFileUrl = data[i][CONFIG.cols.fileLink - 1];
    const lastSyncTime = data[i][CONFIG.cols.lastSync - 1] ? new Date(data[i][CONFIG.cols.lastSync - 1]) : null;

    if (!itemName || !itemFileUrl) continue;

    if (!isSilent) masterListSheet.getRange(row, CONFIG.cols.status).setValue("확인 중...");

    try {
      const itemFileId = itemFileUrl.match(/[-\w]{25,}/);
      if (!itemFileId) {
          masterListSheet.getRange(row, CONFIG.cols.status).setValue("오류: 유효하지 않은 링크");
          continue;
      };

      const itemFile = DriveApp.getFileById(itemFileId[0]);
      const lastUpdatedTime = itemFile.getLastUpdated();

      let shouldUpdate = false;
      if (itemsToPull || !lastSyncTime || lastUpdatedTime.getTime() > lastSyncTime.getTime() + 1000) {
        shouldUpdate = true;
      }
      
      if (shouldUpdate) {
        if (!isSilent) ss.toast(`${itemName} 데이터 가져오는 중...`, "Pull 동기화", 10);
        
        const itemSpreadsheet = SpreadsheetApp.openById(itemFileId[0]);
        const sourceSheet = itemSpreadsheet.getSheetByName(itemName) || itemSpreadsheet.getSheets()[0];
        const sourceData = sourceSheet.getDataRange().getValues();

        let destinationSheet = ss.getSheetByName(itemName);

        if (destinationSheet) {
          destinationSheet.clearContents(); 
        } else {
          const templateSheet = ss.getSheetByName(CONFIG.sheetNames.template);
          destinationSheet = templateSheet.copyTo(ss).setName(itemName);
        }
        
        destinationSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

        masterListSheet.getRange(row, CONFIG.cols.lastSync).setValue(lastUpdatedTime);
        if (!isSilent) masterListSheet.getRange(row, CONFIG.cols.status).setValue(`Pull 완료 (${lastUpdatedTime.toLocaleTimeString()})`);
        updatedCount++;
      } else {
          if (!isSilent) masterListSheet.getRange(row, CONFIG.cols.status).setValue("최신 상태");
      }

    } catch (e) {
      const errorMessage = `Pull 오류: ${e.message.substring(0, 100)}`;
      masterListSheet.getRange(row, CONFIG.cols.status).setValue(errorMessage);
      console.error(`Pull 동기화 오류 (${itemName}): ${e.toString()}`);
    }
  }
  return updatedCount;
}

// =========================================================================
// SCRIPT 5: 트리거 및 유틸리티 함수
// =========================================================================
function sortItemSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterListSheet = ss.getSheetByName(CONFIG.sheetNames.masterList);
  if (!masterListSheet) return;

  const itemInfoList = masterListSheet.getRange(2, CONFIG.cols.id, masterListSheet.getLastRow() - 1, CONFIG.cols.itemName)
    .getValues()
    .map(row => ({ id: row[CONFIG.cols.id - 1], name: row[CONFIG.cols.itemName - 1] }))
    .filter(info => info.name && !isNaN(info.id))
    .sort((a, b) => a.id - b.id);

  const allSheets = ss.getSheets();
  const sheetMap = {};
  allSheets.forEach(sheet => {
    sheetMap[sheet.getName()] = sheet;
  });

  let targetPosition = CONFIG.settings.fixedSheets.length + 1;

  itemInfoList.forEach(itemInfo => {
    const sheetToSort = sheetMap[itemInfo.name];
    if (sheetToSort && sheetToSort.getIndex() !== targetPosition) {
      ss.setActiveSheet(sheetToSort);
      ss.moveActiveSheet(targetPosition);
    }
    if (sheetToSort) {
        targetPosition++;
    }
  });
}

function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() !== 'onOpen') { // onOpen은 남겨둘 수 있습니다.
        ScriptApp.deleteTrigger(trigger);
    }
  }
}

function enableOnEditTrigger() {
  deleteAllTriggers(); 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

function enableAutoSyncOnOpen(isSilent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getEventType() === ScriptApp.EventType.ON_OPEN && trigger.getHandlerFunction() === "autoPullDataOnOpen") {
      if (!isSilent) SpreadsheetApp.getUi().alert("자동 동기화 기능이 이미 켜져 있습니다.");
      return;
    }
  }
  ScriptApp.newTrigger("autoPullDataOnOpen")
    .forSpreadsheet(ss)
    .onOpen()
    .create();
  if (!isSilent) SpreadsheetApp.getUi().alert("설정 완료! 이제부터 파일을 열 때마다 변경된 데이터가 자동으로 동기화됩니다.");
}

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
// SCRIPT 6: 개별 파일 링크 이메일 발송 기능
// =========================================================================
function sendItemFileLinksByEmail_Menu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterListSheet = ss.getSheetByName(CONFIG.sheetNames.masterList);

  if (!masterListSheet) {
    ui.alert(`오류: '${CONFIG.sheetNames.masterList}' 시트를 찾을 수 없습니다.`);
    return;
  }

  const response = ui.alert('이메일 발송 확인', `'${CONFIG.sheetNames.masterList}' 시트의 모든 ${CONFIG.terms.item}에게 이메일을 발송하시겠습니까?`, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    ui.alert('이메일 발송이 취소되었습니다.');
    return;
  }

  ss.toast(`${CONFIG.terms.item} 파일 링크 이메일 발송을 시작합니다...`, "이메일 발송", -1);

  const data = masterListSheet.getDataRange().getValues();
  let sentCount = 0;
  let failCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const itemName = data[i][CONFIG.cols.itemName - 1];
    const itemEmail = data[i][CONFIG.cols.email - 1];
    const itemLink = data[i][CONFIG.cols.fileLink - 1];

    if (!itemName || !itemEmail || !itemLink || !isValidEmail(itemEmail)) {
      let statusMessage = "발송 실패: ";
      if (!itemName) statusMessage += "이름 없음. ";
      if (!itemEmail || !isValidEmail(itemEmail)) statusMessage += "유효하지 않은 이메일. ";
      if (!itemLink) statusMessage += "링크 없음. ";
      masterListSheet.getRange(row, CONFIG.cols.status).setValue(statusMessage.trim());
      failCount++;
      continue;
    }

    ss.toast(`${itemName} ${CONFIG.terms.item}에게 이메일 발송 중...`, "이메일 발송", 5);

    try {
      // CONFIG에서 이메일 템플릿 가져오기
      let subject = CONFIG.email.subject.replace(/{{itemName}}/g, itemName);
      let body = CONFIG.email.body.replace(/{{itemName}}/g, itemName)
                                 .replace(/{{itemLink}}/g, itemLink)
                                 .replace(/{{itemFile}}/g, CONFIG.terms.itemFile);

      GmailApp.sendEmail(itemEmail, subject, body);
      masterListSheet.getRange(row, CONFIG.cols.status).setValue(`이메일 발송 완료 (${new Date().toLocaleTimeString()})`);
      sentCount++;
    } catch (e) {
      const errorMessage = `이메일 발송 실패: ${e.message.substring(0, 100)}`;
      masterListSheet.getRange(row, CONFIG.cols.status).setValue(errorMessage);
      console.error(`이메일 발송 오류 (${itemName}, ${itemEmail}): ${e.toString()}`);
      failCount++;
    }
  }

  ss.toast("이메일 발송 작업 완료!", "이메일 발송", 10);
  ui.alert(`이메일 발송 완료!\n\n총 ${sentCount}건의 이메일을 성공적으로 보냈습니다.\n실패: ${failCount}건.`);
}

function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(String(email));
}
// =========================================================================
// SCRIPT 7: 양식 부분 업데이트 기능 (★★ 가장 안정적인 방식으로 전면 수정 ★★)
// =========================================================================

/**
 * 사용자에게 양식 업데이트를 시작할지 묻는 메뉴 함수
 * (주의: 이 기능은 이제 '부분'이 아닌 '전체 양식'을 기준으로 데이터를 병합합니다)
 */
function updateSheetPortion_Menu() {
  const ui = SpreadsheetApp.getUi();

  // 기능의 동작 방식을 명확히 안내
  const confirmResponse = ui.alert(
    '‼️ 양식 전체 업데이트 ‼️',
    `이 기능은 현재 '양식' 시트의 모양(서식, 셀 간격 등)을 모든 ${CONFIG.terms.itemFile}에 적용합니다.\n\n` +
    `각 ${CONFIG.terms.item}이 기존에 작성한 데이터는 최대한 보존하여 새 양식에 병합됩니다.\n\n` +
    `계속하시겠습니까? (이 작업은 되돌릴 수 없습니다)`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert('작업이 취소되었습니다.');
    return;
  }

  // 핵심 로직 실행 (더 이상 범위를 물어보지 않음)
  updateSheetPortionCore();
}

/**
 * 실제 양식 업데이트를 수행하는 핵심 로직 (★★ 마스터 시트 동기화 기능 최종 수정 ★★)
 */
function updateSheetPortionCore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  ss.toast('양식 업데이트를 시작합니다...', '준비 중', -1);

  // ... (함수의 앞부분은 이전과 동일하게 유지됩니다) ...
  const masterListSheet = ss.getSheetByName(CONFIG.sheetNames.masterList);
  const templateSheet = ss.getSheetByName(CONFIG.sheetNames.template);
  if (!masterListSheet || !templateSheet) { /* ... */ return; }
  const data = masterListSheet.getDataRange().getValues();
  let successCount = 0;
  let failCount = 0;
  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const itemName = data[i][CONFIG.cols.itemName - 1];
    const itemFileUrl = data[i][CONFIG.cols.fileLink - 1];
    if (!itemName || !itemFileUrl) continue;
    ss.toast(`(${i}/${data.length - 1}) '${itemName}' 파일 업데이트 중...`, '진행 중', 20);
    let itemSpreadsheet;
    try {
      // ... (개별 파일을 업데이트하는 로직은 이전과 동일) ...
      const itemFileId = itemFileUrl.match(/[-\w]{25,}/);
      if (!itemFileId) throw new Error("유효하지 않은 파일 링크");
      itemSpreadsheet = SpreadsheetApp.openById(itemFileId[0]);
      const oldSheet = itemSpreadsheet.getSheetByName(itemName) || itemSpreadsheet.getSheets()[0];
      if (!oldSheet) throw new Error(`기존 '${itemName}' 시트를 찾을 수 없음`);
      const oldData = oldSheet.getDataRange().getValues();
      const newSheet = templateSheet.copyTo(itemSpreadsheet);
      if (oldData.length > 0 && oldData[0].length > 0) {
        const requiredRows = oldData.length;
        const requiredCols = oldData[0].length;
        if(newSheet.getMaxRows() < requiredRows) { newSheet.insertRowsAfter(newSheet.getMaxRows(), requiredRows - newSheet.getMaxRows()); }
        if(newSheet.getMaxColumns() < requiredCols) { newSheet.insertColumnsAfter(newSheet.getMaxColumns(), requiredCols - newSheet.getMaxColumns()); }
        newSheet.getRange(1, 1, requiredRows, requiredCols).setValues(oldData);
      }
      itemSpreadsheet.deleteSheet(oldSheet);
      newSheet.setName(itemName);
      itemSpreadsheet.setActiveSheet(newSheet);
      itemSpreadsheet.moveActiveSheet(1);
      masterListSheet.getRange(row, CONFIG.cols.status).setValue(`양식 업데이트 완료 (${new Date().toLocaleTimeString()})`);
      successCount++;
    } catch (e) {
      // ... (에러 처리 로직도 이전과 동일) ...
      try { if(itemSpreadsheet){ const tempSheet = itemSpreadsheet.getSheetByName(templateSheet.getName()); if(tempSheet) itemSpreadsheet.deleteSheet(tempSheet); } } catch (cleanError) { console.error(`오류 정리 중 추가 오류 발생: ${cleanError.toString()}`); }
      const errorMessage = `업데이트 오류: ${e.message.substring(0, 100)}`;
      masterListSheet.getRange(row, CONFIG.cols.status).setValue(errorMessage);
      console.error(`'${itemName}' 파일 업데이트 오류: ${e.message}\n${e.stack}`);
      failCount++;
    }
  }

  // --- ▼▼▼ [최종 수정된 부분] ▼▼▼ ---
  if (successCount > 0) {
    // 이제 pullDataCore 대신, 새로 만든 forceSyncAllSheetsWithFormat 함수를 호출합니다.
    forceSyncAllSheetsWithFormat();
  } else {
    ss.toast('업데이트된 파일이 없어 동기화를 건너뜁니다.', '작업 완료', 10);
  }
  // --- ▲▲▲ [여기까지 수정] ▲▲▲ ---

  ui.alert(`양식 업데이트 완료!\n\n성공: ${successCount} 건\n실패: ${failCount} 건\n\n마스터 시트의 모든 탭이 최신 양식으로 동기화되었습니다.`);
}
/**
 * [신규] 모든 개별 파일의 시트를 마스터 시트로 '통째로' 가져와 동기화하는 함수
 * 데이터 뿐만 아니라 서식, 셀 간격 등 모든 모양을 동기화합니다.
 */
function forceSyncAllSheetsWithFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterListSheet = ss.getSheetByName(CONFIG.sheetNames.masterList);
  if (!masterListSheet) return;

  ss.toast('마스터 시트의 모든 탭을 최신 양식으로 동기화합니다...', '진행 중', -1);
  const data = masterListSheet.getDataRange().getValues();

  // 루프를 돌며 각 항목 시트를 강제로 업데이트
  for (let i = 1; i < data.length; i++) {
    const itemName = data[i][CONFIG.cols.itemName - 1];
    const itemFileUrl = data[i][CONFIG.cols.fileLink - 1];

    if (!itemName || !itemFileUrl) continue;
    
    try {
      // 1. 개별 파일에서 최신 양식이 적용된 소스 시트를 가져옴
      const itemSpreadsheet = SpreadsheetApp.openById(itemFileUrl.match(/[-\w]{25,}/)[0]);
      const sourceSheet = itemSpreadsheet.getSheetByName(itemName) || itemSpreadsheet.getSheets()[0];

      // 2. 마스터 시트에 있는 기존의 낡은 시트를 삭제
      const oldMasterSheet = ss.getSheetByName(itemName);
      if (oldMasterSheet) {
        ss.deleteSheet(oldMasterSheet);
      }
      
      // 3. 소스 시트를 마스터 시트로 통째로 복사
      const newMasterSheet = sourceSheet.copyTo(ss);
      
      // 4. 복사된 시트의 이름을 원래 이름으로 변경 (예: "사본 - 학생1" -> "학생1")
      newMasterSheet.setName(itemName);

    } catch (e) {
      console.error(`'${itemName}' 탭 동기화 실패: ${e.message}`);
      // 실패하더라도 상태 열에 기록
      masterListSheet.getRange(i + 1, CONFIG.cols.status).setValue(`탭 동기화 실패: ${e.message.substring(0, 50)}`);
    }
  }

  // 모든 작업이 끝난 후, 시트 순서를 다시 정렬
  sortItemSheets();
  ss.toast('마스터 시트의 탭 동기화가 완료되었습니다.', '완료', 10);
}