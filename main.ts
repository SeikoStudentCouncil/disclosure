const folderId = "1TuP8dRasRAY-M-1OPJh7wg8x50tx63Dx";
function function1() {
  const parent = DriveApp.getFolderById(folderId);
  const path = parent.getName();
  /*  for (const ans of getFileListGenerator(parent, path)) {
    console.log(ans);
  } */
  const sheet = SpreadsheetApp.getActiveSheet();
  const fileList = Array.from(getFileListGenerator(parent, path));
  fileList.unshift(["fileId", "oldFile", "Folder", "TimeStamp", "MimeType"]);
  sheet.getRange(1, 3, fileList.length, 5).setValues(fileList);
}
function getHyperlink(url: string, linkLabel: string) {
  return `=HYPERLINK("${url}","${linkLabel}")`;
}
function* getFileListGenerator(
  parent: GoogleAppsScript.Drive.Folder,
  path: string
): Generator<[string, string, string, string, string]> {
  const childFiles = parent.getFiles();
  while (childFiles.hasNext()) {
    const childFile = childFiles.next();
    yield [
      childFile.getId(),
      getHyperlink(childFile.getUrl(), childFile.getName()),
      getHyperlink(parent.getUrl(), path),
      "JST - " +
        Utilities.formatDate(
          childFile.getLastUpdated(),
          "JST",
          "yyyy/MM/dd (E) HH:mm:ss Z"
        ),
      childFile.getMimeType(),
    ];
  }
  const childFolders = parent.getFolders();
  while (childFolders.hasNext()) {
    const childFolder = childFolders.next();
    yield* getFileListGenerator(
      childFolder,
      path + "/" + childFolder.getName()
    );
  }
}
function onInstall(/* event: GoogleAppsScript.Events.AddonOnInstall */) {
  onOpen();
}
function onOpen(/* event: GoogleAppsScript.Events.SheetsOnOpen */) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("function1", "function1")
    .addItem("function2", "function2")
    .addToUi();
}
function function2() {
  const currentCell = SpreadsheetApp.getCurrentCell();
  const sheet = currentCell.getSheet();
  const newFileCells = sheet.getRange(currentCell.getRow(), 1, 1, 2);
  const oldFileCell = sheet.getRange(currentCell.getRow(), 2);
  const fileId: string = oldFileCell.getValue();
  const oldFile = DriveApp.getFileById(fileId);
  const newFile = oldFile.makeCopy();
  newFileCells.setValues([
    [newFile.getId(), getHyperlink(newFile.getUrl(), newFile.getName())],
  ]);
}
