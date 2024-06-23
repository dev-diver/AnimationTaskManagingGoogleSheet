const LIBRARY_NAME = 'AM';
const LIBRARY_VERSION = 'HEAD'; // 최신 개발 버전 참조

function onClickBtn() {
  this[LIBRARY_NAME].syncWorkerToPart()
}

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  this[LIBRARY_NAME].onEditDo(e)
}