
function getfiles(filedetail){
  var allFiles = DriveApp.getFiles();
  var specificfiles = getFilesByName(filedetail);
  return allFiles.getContinuationToken()
}
;

function get_folders(){
  // idをもとにフォルダを取得
  const mst_folder_ID = '1lX_iSBFKSRA5z9OMPfhj1MiIMnIZLkiq';  
  var this_sheet = SpreadsheetApp.getActiveSheet();

  var name = ""
  var i = 0 //フォルダを処理する行位置
  var j = 0 //サブフォルダを出力する行位置
  do {
    //フォルダ一覧を取得
    var folders = DriveApp.searchFolders("'"+mst_folder_ID+"' in parents");
    //フォルダ一覧からフォルダを順に取り出す
    while(folders.hasNext()){
      //シートにフォルダ名称とIdを出力
      i++
      var folder = folders.next();
      this_sheet.getRange(i, 1).setValue(name + folder.getName());
      this_sheet.getRange(i, 2).setValue(folder.getId());
    }
/*
    //シートからフォルダを取得し次へ
    j++;
    name = this_sheet.getRange(j, 1).getValue() + " > ";
    key = this_sheet.getRange(j, 2).getValue();
*/
  } while (key != ""); //処理するフォルダがなくなるまで
}
;

// 指定したフォルダ以下に含まれるすべてのファイルを列挙
function getFiles(folderID){
  var this_sheet = SpreadsheetApp.getActiveSheet();
  // 「指定したフォルダ内に配置されている」という検索式を作る。
  var searchFileParams = Utilities.formatString("'%s' in parents", folderID);
  var searchFolderParams = searchFileParams;
  var folders;
  // フォルダ検索結果が空になるまで検索を続ける。

  while( (folders = DriveApp.searchFolders(searchFolderParams)).hasNext() )
  {
    // 最後に先頭の" or "を削除したいので先頭を示す"@"を入れておく
    searchFolderParams = "@";
    while( folders.hasNext() )
    {
      // このループでは検索式に" or 'XXX' in parents"を追加していく 
      var folder = folders.next();
      searchFolderParams += Utilities.formatString(" or '%s' in parents", folder.getId());
    }
    // でき上がった検索式の先頭から不要な"@ or "を削除する。
    // searchFolderParamsは次の階層のフォルダ検索式
    searchFolderParams = searchFolderParams.replace( "@ or ","" );
    // searchFileParamsはファイルを一気に検索する検索式
    searchFileParams += " or " + searchFolderParams;
  }
  // 指定フォルダ以下全階層のファイルを一気に検索する。
  return DriveApp.searchFiles(searchFileParams);
}

function getFiles2()
{
  // idをもとにフォルダを取得
  const mst_folder_ID = '1lX_iSBFKSRA5z9OMPfhj1MiIMnIZLkiq';  
  var this_sheet = SpreadsheetApp.getActiveSheet();
  var target_folder_name = this_sheet.getRange(5,7).getValue();
  var target_folder_id = DriveApp.getFolderById(mst_folder_ID).getFoldersByName("20"+target_folder_name.slice(0,4)).next().getFoldersByName(target_folder_name).next().getId()

  var i = 0 //ファイルを格納する行位置  
  var files = getFiles(target_folder_id)
  while(files.hasNext()){
    //シートにファイル名称とIdを出力
    i++
    var file = files.next();
    this_sheet.getRange(i, 1).setValue(file.getName());
    this_sheet.getRange(i, 2).setValue(file.getId());
    this_sheet.getRange(i, 3).setValue(file.getParents().next().getName());
    this_sheet.getRange(i, 5).setValue(file.getLastUpdated());
//    var Height = this_sheet.getImages(DriveApp.getFileById(file.getId()).getBlob()).getInherentHeight();
//    var Width = this_sheet.getImages(DriveApp.getFileById(file.getId()).getBlob()).getInherentWidth();
//    this_sheet.getRange(i, 6).setValue(Height,Width);
    //この関数だけ行と列が逆。。。
    this_sheet.insertImage(DriveApp.getFileById(file.getId()).getBlob(),7,i);
  } 
  }
;

/*
var 
// フォルダに属するファイルを全て取得
var files = folder.getFiles();

// filesをイテレートして, ファイル名をログに出力
while(files.hasNext()) {
  var file = files.next();
  Logger.log(file.getName());
*/
