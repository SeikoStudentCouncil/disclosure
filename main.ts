const scriptProperties = PropertiesService.getScriptProperties();
const oldFolderId = scriptProperties.getProperty("oldFolderId") as string;
const newFolderId = scriptProperties.getProperty("newFolderId") as string;

declare namespace GoogleAppsScript {
  namespace Drive {
    interface File {
      getTargetMimeType(): string | null;
    }
  }
}

interface FileObject {
  path: string;
  row: number | null;
}

function function1() {
  const parent = DriveApp.getFolderById(oldFolderId);
  const path = parent.getName();
  /*  for (const ans of getFileListGenerator(parent, path)) {
    console.log(ans);
  } */
  const sheet = SpreadsheetApp.getActiveSheet();
  const oldFilesMap = getOldFilesMap(sheet);
  const fileList = Array.from(
    getFileListGenerator(parent, path, oldFilesMap, sheet)
  );
  if (!oldFilesMap.has("oldfileId")) {
    fileList.unshift([
      "oldfileId",
      "timeStamp",
      "oldFile",
      "oldFolder",
      "timeStamp",
      "MimeType",
    ]);
  }
  if (fileList.length) {
    sheet
      .getRange(sheet.getLastRow() + 1, 4, fileList.length, 6)
      .setValues(fileList);
  }
}
function getOldFilesMap(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  return new Map(
    Array.from(
      (
        sheet.getRange("D:G").getValues() as [string, string, string, string][]
      ).entries()
    ).map((v) => [v[1][0], { path: v[1][3], row: v[0] + 1 } as FileObject])
  );
}
function getHyperlink(url: string, linkLabel: string) {
  return `=HYPERLINK("${url}","${linkLabel}")`;
}
function* getFileListGenerator(
  parent: GoogleAppsScript.Drive.Folder,
  path: string,
  oldFilesMap: Map<string, FileObject>,
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): Generator<[string, string, string, string, string, string]> {
  const childFiles = parent.getFiles();
  while (childFiles.hasNext()) {
    const childFile = childFiles.next();
    const targetMimeType = childFile.getTargetMimeType();
    try {
      if (targetMimeType) {
        // if shortcut
        const targetId = childFile.getTargetId() as string;
        if (targetMimeType === "application/vnd.google-apps.folder") {
          //if folder shortcut

          const childFolder = DriveApp.getFolderById(targetId);
          yield* getFileListGenerator(
            childFolder,
            path + "/" + childFolder.getName(),
            oldFilesMap,
            sheet
          );
        } else {
          // if file shortcut
          const targetChildFile = DriveApp.getFileById(targetId);
          const fileObject = oldFilesMap.get(targetId);
          if (fileObject) {
            if (fileObject.path === path) {
              continue;
            } else {
              if (fileObject.row) {
                sheet.deleteRow(fileObject.row);
              }
            }
          }
          oldFilesMap.set(targetId, { path, row: null });
          yield [
            targetId,
            String(targetChildFile.getLastUpdated().getTime()),
            getHyperlink(targetChildFile.getUrl(), targetChildFile.getName()),
            getHyperlink(parent.getUrl(), path),
            "JST - " +
              Utilities.formatDate(
                targetChildFile.getLastUpdated(),
                "JST",
                "yyyy/MM/dd (E) HH:mm:ss Z"
              ),
            targetChildFile.getMimeType(),
          ];
        }
      } else {
        const childFileId = childFile.getId();
        const fileObject = oldFilesMap.get(childFileId);
        if (fileObject) {
          if (fileObject.path === path) {
            continue;
          } else {
            if (fileObject.row) {
              sheet.deleteRow(fileObject.row);
            }
          }
        }
        oldFilesMap.set(childFileId, { path, row: null });
        yield [
          childFileId,
          String(childFile.getLastUpdated().getTime()),
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
    } catch (e) {
      console.log(e);
    }
  }
  const childFolders = parent.getFolders();
  while (childFolders.hasNext()) {
    const childFolder = childFolders.next();
    yield* getFileListGenerator(
      childFolder,
      path + "/" + childFolder.getName(),
      oldFilesMap,
      sheet
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
      4
    );
    const fileIds = (oldFilesCells.getValues() as string[][]).map(
      (v): [string, string] => [v[0], v[3]]
    );
    newFileCells.setValues(
      fileIds.map(function ([fileId, path]): [string, string, string] {
        const oldFile = DriveApp.getFileById(fileId);
        const destination = cd(newFolder, path);
        const newFile = oldFile.makeCopy(oldFile.getName(), destination);
        return [
          newFile.getId(),
          getHyperlink(newFile.getUrl(), newFile.getName()),
          getHyperlink(destination.getUrl(), path),
        ];
      })
    );
  }
}
