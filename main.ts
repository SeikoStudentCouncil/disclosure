const scriptProperties = PropertiesService.getScriptProperties();
const oldFolderId = scriptProperties.getProperty("oldFolderId") as string;
const newFolderId = scriptProperties.getProperty("newFolderId") as string;
function function1() {
  const parent = DriveApp.getFolderById(oldFolderId);
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
function cd(currentDirectory: GoogleAppsScript.Drive.Folder, path: string) {
  return mkdir(currentDirectory, path.split("/").reverse());
}
/**
 * make folders with nested folders
 * pathList is like this:
 *
 *     const path = "root/parent/sub"
 *     const pathList = path.split("/").reverse()
 *     pathList==['sub', 'parent', 'root']
 */
function mkdir(
  currentDirectory: GoogleAppsScript.Drive.Folder,
  pathList: string[]
): GoogleAppsScript.Drive.Folder {
  if (!pathList.length) {
    return currentDirectory;
  }
  const childFolderName = pathList.pop() as string;
  const folderIterator = currentDirectory.getFoldersByName(childFolderName);
  const childFolder = folderIterator.hasNext()
    ? folderIterator.next()
    : currentDirectory.createFolder(childFolderName);
  return mkdir(childFolder, pathList);
}
function function2() {
  const newFolder = DriveApp.getFolderById(newFolderId);
  const rangeList = SpreadsheetApp.getActiveRangeList();
  const sheet = SpreadsheetApp.getActiveSheet();
  for (const range of rangeList.getRanges()) {
    const newFileCells = sheet.getRange(
      range.getRow(),
      1,
      range.getNumRows(),
      3
    );
    const oldFilesCells = sheet.getRange(
      range.getRow(),
      4,
      range.getNumRows(),
      3
    );
    const fileIds = oldFilesCells
      .getValues()
      .map((v: string[]): [string, string] => [v[0], v[2]]);
    newFileCells.setValues(
      fileIds.map(function ([fileId, path]): [string, string,string] {
        const oldFile = DriveApp.getFileById(fileId);
        const destination = cd(newFolder, path);
        const newFile = oldFile.makeCopy(
          oldFile.getName(),
          destination
        );
        return [
          newFile.getId(),
          getHyperlink(newFile.getUrl(), newFile.getName()),
          getHyperlink(destination.getUrl(),path)
        ];
      })
    );
  }
}
