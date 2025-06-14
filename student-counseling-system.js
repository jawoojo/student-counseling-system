// =========================================================================
// SCRIPT 1: 파일 생성 & 초기 동기화 (★★시트 구조 변경 대응 버전★★)
// =========================================================================
function createAndShareStudentSheets() {
 const ui = SpreadsheetApp.getUi();
 const CONFIRM_TEXT = "초기설정";


 const response = ui.prompt(
   '‼️매우 중요: 실행 확인‼️',
   `이 작업은 모든 학생 파일을 새로 만들고, '학생목록' 시트의 링크를 덮어씁니다.\n\n계속하려면 아래 입력창에 '${CONFIRM_TEXT}'라고 정확히 입력한 후 [확인]을 누르세요.`,
   ui.ButtonSet.OK_CANCEL
 );


 if (response.getSelectedButton() !== ui.Button.OK || response.getResponseText().trim() !== CONFIRM_TEXT) {
   ui.alert('입력 내용이 다르거나 취소되었습니다. 작업이 실행되지 않았습니다.');
   return;
 }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
 ss.toast("학생 파일 생성을 시작합니다...", "초기 설정", -1);


 const templateSheet = ss.getSheetByName("양식");
 const studentListSheet = ss.getSheetByName("학생목록");
 const data = studentListSheet.getDataRange().getValues();
  const FOLDER_NAME = "학생 상담카드";
 let targetFolder;
 const folders = DriveApp.getFoldersByName(FOLDER_NAME);
 if (folders.hasNext()) {
   targetFolder = folders.next();
 } else {
   targetFolder = DriveApp.createFolder(FOLDER_NAME);
 }


 for (let i = 1; i < data.length; i++) {
   // ★★★ 구조 변경 대응: 열 인덱스 수정 ★★★
   // A열: 번호(0), B열: 이름(1)
   const studentName = data[i][1];
   if (!studentName) continue;
   ss.toast(`${studentName} 학생 파일 생성 중...`, "초기 설정", 5);


   const newSpreadsheet = SpreadsheetApp.create(studentName);
   const newFile = DriveApp.getFileById(newSpreadsheet.getId());
  
   targetFolder.addFile(newFile);
   DriveApp.getRootFolder().removeFile(newFile);


   if (templateSheet) {
     const copiedSheet = templateSheet.copyTo(newSpreadsheet);
     const defaultSheet = newSpreadsheet.getSheets()[0];
     newSpreadsheet.deleteSheet(defaultSheet);
     copiedSheet.setName(studentName);
   } else {
     newSpreadsheet.getSheets()[0].setName(studentName);
   }


   newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  
   const sheetUrl = newSpreadsheet.getUrl();
   // ★★★ 구조 변경 대응: 링크는 D열(4)에 저장 ★★★
   studentListSheet.getRange(i + 1, 4).setValue(sheetUrl);
 }


 ss.toast("자동 동기화 트리거를 설정합니다...", "초기 설정", -1);
 enableOnEditTrigger();
 enableAutoSyncOnOpen(true);


 ss.toast("첫 데이터 동기화를 시작합니다. 모든 학생 탭이 생성됩니다...", "초기 설정", -1);
 const updatedCount = pullDataCore(true);


 ss.toast("학생 시트 순서를 정렬합니다...", "초기 설정", -1);
 sortStudentSheets(); // 정렬 함수 호출


 ui.alert(`초기 설정 완료!\n\n총 ${updatedCount}명의 학생 파일 및 시트 생성이 완료되었으며, 번호 순으로 정렬되었습니다.\n\n이제부터 파일을 열 때마다 데이터가 자동으로 동기화됩니다.`);
}


// =========================================================================
// SCRIPT 2: 교사 -> 학생 동기화 (Push)
// =========================================================================
function syncSheetOnEdit(e) {
 const editedSheet = e.source.getActiveSheet();
 const editedSheetName = editedSheet.getName();
 const masterListSheetName = "학생목록";


 if (editedSheetName === masterListSheetName) return;


 const masterListSheet = e.source.getSheetByName(masterListSheetName);
 if (!masterListSheet) return;
  const data = masterListSheet.getDataRange().getValues();
 let targetStudentFileUrl = null;
 let targetRow = -1;


 for (let i = 1; i < data.length; i++) {
   // ★★★ 구조 변경 대응: 이름은 B열(1)과 비교, 링크는 D열(3)에서 가져옴 ★★★
   if (data[i][1] === editedSheetName) {
     targetStudentFileUrl = data[i][3];
     targetRow = i + 1;
     break;
   }
 }


 if (targetStudentFileUrl) {
   try {
     const studentSpreadsheet = SpreadsheetApp.openByUrl(targetStudentFileUrl);
     const newCopiedSheet = editedSheet.copyTo(studentSpreadsheet);
    
     const oldSheet = studentSpreadsheet.getSheetByName(editedSheetName);
     if (oldSheet) studentSpreadsheet.deleteSheet(oldSheet);
    
     newCopiedSheet.setName(editedSheetName);
     studentSpreadsheet.setActiveSheet(newCopiedSheet);
    
     // ★★★ 구조 변경 대응: 동기화 시간은 E열(5)에 기록 ★★★
     masterListSheet.getRange(targetRow, 5).setValue(new Date());


     // SpreadsheetApp.getActiveSpreadsheet().toast(`${editedSheetName} 학생 파일 동기화 완료! (시간 기록 초기화)`);
   } catch (error) {
     SpreadsheetApp.getUi().alert(`동기화 오류 발생: ${error.message}`);
   }
 }
}


// =========================================================================
// SCRIPT 3: 사용자 메뉴 생성 (기존과 동일)
// =========================================================================
function onOpen() {
 const ui = SpreadsheetApp.getUi();
 ui.createMenu('학생 데이터 관리').addItem('① [수동] 변경된 학생 데이터 불러오기', 'pullDataFromStudentFiles_Smart').addSeparator().addItem('② [설정] 자동 동기화 켜기', 'enableAutoSyncOnOpen').addItem('③ [설정] 자동 동기화 끄기', 'disableAutoSyncOnOpen').addToUi();
 ui.createMenu('🚨 초기 설정 (최초 1회)').addItem('모든 학생 파일 생성 및 초기화', 'createAndShareStudentSheets').addToUi();
}


// =========================================================================
// SCRIPT 4: 학생 -> 교사 동기화 (수동/자동 실행 함수)
// =========================================================================
function pullDataFromStudentFiles_Smart() {
 const ui = SpreadsheetApp.getUi();
 const response = ui.alert('알림', '변경사항이 있는 학생의 데이터만 불러옵니다. 계속하시겠습니까?', ui.ButtonSet.YES_NO);
 if (response != ui.Button.YES) return;
  SpreadsheetApp.getActiveSpreadsheet().toast("학생 데이터 확인을 시작합니다...", "상태", -1);
 const updatedCount = pullDataCore(false);
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
 const updatedCount = pullDataCore(true);
 if (updatedCount > 0) {
   sortStudentSheets();
   SpreadsheetApp.getActiveSpreadsheet().toast(`자동 동기화 완료: 총 ${updatedCount}명 업데이트 및 정렬됨`, "완료", 10);
 }
}


// =========================================================================
// SCRIPT 5: 동기화 핵심 로직
// =========================================================================
function pullDataCore(isSilent = false) {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const masterListSheet = ss.getSheetByName("학생목록");
 if (!masterListSheet) return 0;
  const data = masterListSheet.getDataRange().getValues();
 let updatedCount = 0;


 for (let i = 1; i < data.length; i++) {
   // ★★★ 구조 변경 대응: 열 인덱스 수정 ★★★
   const studentSheetName = data[i][1]; // 이름은 B열(1)
   const studentFileUrl = data[i][3];   // 링크는 D열(3)
   const lastSyncTime = data[i][4] ? new Date(data[i][4]) : null; // 동기화 시간은 E열(4)


   if (!studentSheetName || !studentFileUrl) continue;


   try {
     const studentFileId = studentFileUrl.match(/[-\w]{25,}/);
     if (!studentFileId) continue;


     const studentFile = DriveApp.getFileById(studentFileId[0]);
     const lastUpdatedTime = studentFile.getLastUpdated();


     let shouldUpdate = false;
     if (!lastSyncTime) {
       shouldUpdate = true;
     } else {
       const lastSyncSeconds = Math.floor(lastSyncTime.getTime() / 1000);
       const lastUpdatedSeconds = Math.floor(lastUpdatedTime.getTime() / 1000);
       if (lastUpdatedSeconds > lastSyncSeconds) {
         shouldUpdate = true;
       }
     }
    
     if (shouldUpdate) {
       if (!isSilent) ss.toast(`${studentSheetName} 데이터 가져오는 중...`);
      
       const studentSpreadsheet = SpreadsheetApp.openById(studentFileId[0]);
       const sourceSheet = studentSpreadsheet.getSheets()[0];
       const destinationSheet = ss.getSheetByName(studentSheetName);
       if (destinationSheet) ss.deleteSheet(destinationSheet);
      
       const newSheet = sourceSheet.copyTo(ss);
       newSheet.setName(studentSheetName);


       // ★★★ 구조 변경 대응: 동기화 시간은 E열(5)에 기록 ★★★
       masterListSheet.getRange(i + 1, 5).setValue(lastUpdatedTime);
       updatedCount++;
     }
   } catch (error) {
     console.error(`동기화 코어 오류 (${studentSheetName}): ${error.message}`);
     if (!isSilent) SpreadsheetApp.getUi().alert(`${studentSheetName} 학생 데이터 처리 중 오류: ${error.message}`);
   }
 }
 return updatedCount;
}


// =========================================================================
// SCRIPT 6: 트리거 관리 함수 (기존과 동일)
// =========================================================================
function enableOnEditTrigger() { /* ... */ }
function enableAutoSyncOnOpen(isSilent = false) { /* ... */ }
function disableAutoSyncOnOpen() { /* ... */ }
// (생략된 트리거 관리 함수 내용은 이전 코드와 완전히 동일합니다)
function enableOnEditTrigger(){const ss=SpreadsheetApp.getActiveSpreadsheet(),triggers=ScriptApp.getProjectTriggers();for(const trigger of triggers)if("syncSheetOnEdit"===trigger.getHandlerFunction())return;ScriptApp.newTrigger("syncSheetOnEdit").forSpreadsheet(ss).onEdit().create()}function enableAutoSyncOnOpen(isSilent=!1){const ss=SpreadsheetApp.getActiveSpreadsheet(),triggers=ScriptApp.getProjectTriggers();for(const trigger of triggers)if(trigger.getEventType()===ScriptApp.EventType.ON_OPEN&&"autoPullDataOnOpen"===trigger.getHandlerFunction())return void(isSilent||SpreadsheetApp.getUi().alert("자동 동기화 기능이 이미 켜져 있습니다."));ScriptApp.newTrigger("autoPullDataOnOpen").forSpreadsheet(ss).onOpen().create(),isSilent||SpreadsheetApp.getUi().alert("설정 완료! 이제부터 파일을 열 때마다 변경된 학생 데이터가 자동으로 동기화됩니다.")}function disableAutoSyncOnOpen(){const triggers=ScriptApp.getProjectTriggers();let wasTriggerDeleted=!1;for(const trigger of triggers)trigger.getEventType()===ScriptApp.EventType.ON_OPEN&&"autoPullDataOnOpen"===trigger.getHandlerFunction()&&(ScriptApp.deleteTrigger(trigger),wasTriggerDeleted=!0);wasTriggerDeleted?SpreadsheetApp.getUi().alert("자동 동기화 기능이 꺼졌습니다."):SpreadsheetApp.getUi().alert("현재 설정된 자동 동기화 기능이 없습니다.")}




// =========================================================================
// SCRIPT 7: 유틸리티 함수 (★★정렬 로직 변경★★)
// =========================================================================
function sortStudentSheets() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const masterListSheet = ss.getSheetByName("학생목록");
  if (!masterListSheet) return;


 // 1. ★★★ 정렬 기준(번호)과 시트 이름(이름)을 '학생목록'에서 가져옵니다. ★★★
 // getRange("A2:B")로 번호와 이름을 모두 가져옵니다.
 const studentInfoList = masterListSheet.getRange("A2:B" + masterListSheet.getLastRow())
   .getValues()
   .map(row => ({ number: row[0], name: row[1] })) // {번호, 이름} 객체로 만듭니다.
   .filter(info => info.name && !isNaN(info.number)); // 이름과 번호가 모두 유효한 경우만 필터링


 // 번호를 기준으로 오름차순 정렬합니다.
 studentInfoList.sort((a, b) => a.number - b.number);


 // 2. 모든 시트의 이름을 맵으로 만들어 빠른 조회를 가능하게 합니다.
 const allSheets = ss.getSheets();
 const sheetMap = {};
 allSheets.forEach(sheet => {
   sheetMap[sheet.getName()] = sheet;
 });


 // 3. 고정 시트를 제외한 학생 시트의 시작 위치를 결정합니다.
 let targetPosition = 3; // '학생목록', '양식' 시트 다음인 3번째 위치부터 시작


 // 4. 정렬된 목록 순서대로 시트를 이동시킵니다.
 studentInfoList.forEach(studentInfo => {
   const sheetToSort = sheetMap[studentInfo.name]; // 이름으로 시트를 찾습니다.
   if (sheetToSort) {
     if (sheetToSort.getIndex() !== targetPosition) {
       ss.setActiveSheet(sheetToSort);
       ss.moveActiveSheet(targetPosition);
     }
     targetPosition++;
   }
 });
}

