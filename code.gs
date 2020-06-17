function CallCreativeImages(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // idをもとにマスタの共有フォルダを取得
  const mst_folder_ID = '1GJYcqjHXVDh8h7OxAm9ngPy834V_XbqN';
  // 探すフォルダの名前をセルからとる
  var target_folder_name = ss.getSheetByName('Wチェック').getRange(2,2).getValue().toString();
  // 探しにいくフォルダはYYYYMM > YYMMNNNの一つで、その配下の子フォルダ等は見に行く必要無し
  var target_folder_id =
    DriveApp.getFolderById(mst_folder_ID)
    .getFoldersByName("20"+target_folder_name.slice(0,4)).next()
    .searchFolders('title contains "'+target_folder_name+'"').next().getId();

  // 画像呼び込みシートを一から作り直す(GASでは画像を削除する関数がないため、シートの作り直しが必要)
  if (ss.getSheetByName('画像呼び込み') != null)
      {
        ss.deleteSheet(ss.getSheetByName('画像呼び込み'));
        ss.insertSheet('画像呼び込み');
      }
    else
      {
        ss.insertSheet('画像呼び込み');
      };

  // 画像呼び込みシートの体裁整える
  var paste_sheet = ss.getSheetByName('画像呼び込み');
  paste_sheet.insertColumns(1, 100);
  paste_sheet.getRange('A:A').setFontWeight('bold');


  // Part1. 画像を貼り付ける前に、ファイルの一覧を作る
  // 行を数えるためのカウンタiを設定
  var i = 0;

  paste_sheet.getRange(i + 1, 1).setValue('ファイル一覧');
  paste_sheet.getRange(i + 1, 2).setValue('ファイル名');
  paste_sheet.getRange(i + 1, 3).setValue('ファイル種別');
  paste_sheet.getRange(i + 1, 4).setValue('最終更新');
  paste_sheet.getRange(i + 1, 5).setValue('ファイルサイズ(KB)');
  paste_sheet.getRange(i + 1, 6).setValue('ファイルURL');

  // Part1.1 ファイルの情報収集と貼り付け
  var all_files = DriveApp.getFolderById(target_folder_id).getFiles();
  while (all_files.hasNext())
    {
      var file = all_files.next();
      paste_sheet.getRange(i + 2, 2).setValue(file.getName());
      paste_sheet.getRange(i + 2, 3).setValue(file.getMimeType());
      paste_sheet.getRange(i + 2, 4).setValue(file.getLastUpdated());
      paste_sheet.getRange(i + 2, 5).setValue(file.getSize() / 1000);
      paste_sheet.getRange(i + 2, 6).setFormula("=hyperlink(\""+file.getUrl()+"\",\"link\")");
      i++;
    };

  // Part1.2 フォルダの情報収集と貼り付け
  var all_folders = DriveApp.getFolderById(target_folder_id).getFolders();
  while (all_folders.hasNext())
    {
      var folder = all_folders.next();
      paste_sheet.getRange(i + 2, 2).setValue(folder.getName());
      paste_sheet.getRange(i + 2, 3).setValue('folder');
      paste_sheet.getRange(i + 2, 4).setValue(folder.getLastUpdated());
      paste_sheet.getRange(i + 2, 6).setFormula("=hyperlink(\""+folder.getUrl()+"\",\"link\")");
      i++;
    };


  // Part2 画像を呼び込む
  // 画像と映像のファイルとってくる(用件としてはjpg.png.gif.mp4だけで一応は大丈夫)
  // Part2.1 下準備
  // キャプチャはあとで除く(search queryで()とかを使った条件設定がわからなかった。。)
  var files = DriveApp.getFolderById(target_folder_id).searchFiles('mimeType contains "image/" or mimeType contains "video/"');

  // 各ブロックにおける貼り付けの行位置は変わらず、列位置だけが変動する
  // 縦方向だけでなく、横にもカウンタが必要(一つのcreative_typeで複数画像を表示するので)
  // 列位置に関しては、列挿入なども考えられるため、基準列を一度変数化(startC)してからそれを利用
  // C = Column, R = Row
  var startC = 2;
  // ここから先のcreative_typeの順番は、貼り付け先の行と非表示の行と連動するので、それらと統一する
  // rc = REC
  var rcC = startC;
  // wr = WRECTANGLE
  var wrC = startC;
  // ti = TOP IMPACT
  var tiC = startC;
  // ti = TOP IMPACT REMINDER
  var tirC = startC;
  // sh = SUPER HERO
  var shC = startC;
  // is = INTER SCROLLER
  var isC = startC;
  // bp = BIGPANEL
  var bpC = startC;
  // bt = BOTTOM
  var btC = startC;
  // sb = SUPER BANNER
  var sbC = startC;
  // hs = HOUSE
  var hsC = startC;
  // ブロックから漏れたものは別処理を施す(チェックリスト外シートに入れる)
  var leftC = startC;

  // キャプチャ判定のためのフラグを用意
  var capture_flg = 0;

  // 後述の画像サイズ取得のための逃し先のシートを作る
  if (ss.getSheetByName('getImageSize') != null)
    {
      ss.deleteSheet(ss.getSheetByName('getImageSize'));
      ss.insertSheet('getImageSize');
    }
  else
    {
      ss.insertSheet('getImageSize');
    };
  var hidden_sheet = ss.getSheetByName('getImageSize');

  while(files.hasNext())
    {
      // Part2.2 ファイル情報の整理
      var file = files.next();
      var fileName = file.getName();
      var fileId = file.getId();
      var fileByte = file.getSize();
      var fileLastUpdated = file.getLastUpdated();
      var fileMimeType = file.getMimeType();
      var fileAt2xflg = fileName.indexOf('@2x') != -1 ? 1 : 0;
      var fileThumbNailBlob = file.getThumbnail();
      var fileUrl = file.getUrl();

      // キャプチャは画像サイズが大きいとGASでinsertできない
      // https://qiita.com/Cesaroshun/items/f79e90dec82c8cec4676
      // ので、フラグ化して逃す
      if (fileName.indexOf('capture') >= 0 || fileName.indexOf('キャプチャ') >= 0)
          {
            capture_flg = 1;
          }
        // 以降はキャプチャ以外の素材に対する処理
        else
          {
            // 画像の縦横のサイズをとるメソッドがDriveAppに見つからない、、、
            // 苦肉の策として、DriveAppから持ってきたBlobをスプレッドシートに1回imageとして貼ってサイズを取得、変数として活用、を繰り返す

            // 画像ファイルのサイズを引っ張ってくるところ
            // DriveAppで上記できる様になり次第改修(優先度高)
            if (fileMimeType.indexOf('image/') != -1 && fileMimeType.indexOf('image/gif') == -1)
              {
                // 画像を仮シートに置く
                hidden_sheet.insertImage(DriveApp.getFileById(fileId).getBlob(),1,1);
                // 置いた画像からサイズを取得する
                // while文の中で使い回す際、画像がある時とない時がある(動画など)
                // その際、画像貼り付け様のカウンタを作成するのはめんどくさいので、都度貼って消してを繰り返すことにより、常に1枚しかない状態を保つ
                var fileImageTmp = hidden_sheet.getImages()[0];
                var fileHeight = fileImageTmp.getInherentHeight();
                var fileWidth = fileImageTmp.getInherentWidth();
                fileImageTmp.remove();
              };

            // Part2.3 画像の貼り付け
            // このブロックでサイズごとに変動するカウンタを設定して、file_iとjに落とす
            // iカウンタと分けたのは一番始めにディレクトリの中身一覧を作っているので行が可変になってしまうから。。。
            // ネーミングは要改修
            // addRowは、creative_typeによって貼り付け先の行を振り分けるためのもの
            var files_i = 0;
            var addRow = 0;
            var j = 0;

            // 頻出処理(貼り付け先の行確定はfunctionにしてしまう)
            function calc_Row(){eval('file_i = i + addRow + 2;')};

            // ここから、ファイルの縦横のサイズによってcreative_typeを判定する
            // 三項演算子等使うか悩んだものの、複数行の処理であること、可読性を維持することなどのためにif文で処理する
            // ここのaddRowの順番が、上述のcreative_typeの順番と同一でなければならない
            if (fileMimeType.indexOf('image/') == -1 || fileMimeType.indexOf('image/gif') >= 0)
                {
                  // 動画の時は条件にかかわらずリスト外
                  addRow = 201;
                  j = leftC;
                  leftC = leftC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('その他');
                }
              else if(fileWidth == 300 && fileHeight == 250)
                {
                  addRow = 1;
                  j = rcC;
                  rcC = rcC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('RECTANGLE');
                }
              else if(fileWidth == 300 && fileHeight == 600)
                {
                  addRow = 21;
                  j = wrC;
                  wrC = wrC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('WRECTANGLE')
                }
              else if(fileWidth == 828 && fileHeight == 752)
                {
                  addRow = 41;
                  j = tiC;
                  tiC = tiC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('TOPIMPACT')
                }
              else if(fileWidth == 828 && fileHeight == 360)
                {
                  addRow = 61;
                  j = tirC;
                  tirC = tirC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('TOPIMPACT REMINDER')
                }
              else if((fileWidth == 320 && fileHeight == 450) || (fileWidth == 1920 && fileHeight == 450))
                {
                  addRow = 81;
                  j = shC;
                  shC = shC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('SUPERHERO')
                }
              else if(fileWidth == 640 && fileHeight == 1386)
                {
                  addRow = 101;
                  j = isC;
                  isC = isC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('INTERSCROLLER')
                }
              else if(fileWidth == 640 && fileHeight == 360)
                {
                  addRow = 121;
                  j = bpC;
                  bpC = bpC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('BIGPANEL')
                }
              else if(fileWidth == 640 && fileHeight == 100)
                {
                  addRow = 141;
                  j = btC;
                  btC = btC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('BOTTOM')
                }
              else if(fileWidth == 728 && (fileHeight == 90||fileHeight == 91))
                {
                  addRow = 161;
                  j = sbC;
                  sbC = sbC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('SUPERBANNER')
                }
              else if(fileWidth == 640 && fileHeight == 1)
                {
                  addRow = 181;
                  j = hsC;
                  hsC = hsC + 5;
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('HOUSE')
                }
              else
                {
                  addRow = 201;
                  j = leftC;
                  leftC = leftC + 5
                  calc_Row();
                  paste_sheet.getRange(file_i, 1).setValue('その他')
                };

            // 隣接する列を先に定義する(あとで挿入する時とかにめんどくさいので)
            var jnext = j + 1;

            // 得られた情報を元に入力をしていく
            paste_sheet.getRange(file_i, j).setValue('ファイル名');
            paste_sheet.getRange(file_i+1, j).setValue('ファイルID');
            paste_sheet.getRange(file_i+2, j).setValue('ファイル容量(KB)');
            paste_sheet.getRange(file_i+3, j).setValue('ファイル最終更新日時');
            paste_sheet.getRange(file_i+4, j).setValue('@2xフラグ');
            paste_sheet.getRange(file_i+5, j).setValue('ファイル種別');
            paste_sheet.getRange(file_i+6, j).setValue('ファイルURL');

            paste_sheet.getRange(file_i, jnext).setValue(fileName);
            paste_sheet.getRange(file_i+1, jnext).setValue(fileId);
            paste_sheet.getRange(file_i+2, jnext).setValue(fileByte / 1000).setNumberFormat('0.0');
            paste_sheet.getRange(file_i+3, jnext).setValue(fileLastUpdated).setNumberFormat('yyyy-mm-dd hh:mm:ss');
            paste_sheet.getRange(file_i+4, jnext).setValue(fileAt2xflg);
            paste_sheet.getRange(file_i+5, jnext).setValue(fileMimeType);
            paste_sheet.getRange(file_i+6, jnext).setFormula("=hyperlink(\""+fileUrl+"\",\"link\")");

            if (fileMimeType.indexOf('image/') != -1 && fileMimeType.indexOf('image/gif') == -1)
              {
                paste_sheet.getRange(file_i+7, j).setValue('ファイル画像サイズ(横×縦)')
                paste_sheet.getRange(file_i+7, jnext).setValue(fileWidth+' × '+fileHeight);
                //この関数だけ行と列が逆。。。
                paste_sheet.insertImage(DriveApp.getFileById(fileId).getBlob(),jnext+1,file_i).setWidth(fileWidth*0.5).setHeight(fileHeight*0.5);
              };
          };
    };

  // Part3. 使っていない行を非表示にする
  // 使っていないものはstartCと開始Cが一致するのでそこから判定をする
  // ここの順番が、上述のcreative_typeの順番と同一でなければならない
  var creative_type = [rcC,wrC,tiC,tirC,shC,isC,bpC,btC,sbC,hsC]
  var creative_type_rowseq = [1,21,41,61,81,101,121,141,161,181,201]
    for(let creative_i=0; creative_i<creative_type.length; creative_i++)
      {
        if (creative_type[creative_i] == startC)
          {
            paste_sheet.hideRows(i + creative_type_rowseq[creative_i] + 2, 20)
          };
      };

  // キャプチャを含んでいるのであれば一応注意を出す
  if (capture_flg == 1)
    {
      Browser.msgBox('"キャプチャ/capture"がファイル名に含まれる画像は呼び込んでいません',Browser.Buttons.OK)
    };

  // 画像貼り付けの逃げ先となるシートを削除しておく
  if (ss.getSheetByName('getImageSize') != null)
    {
      ss.deleteSheet(ss.getSheetByName('getImageSize'));
    };
}
;
