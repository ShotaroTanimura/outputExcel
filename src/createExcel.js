const XLSX = require('xlsx');

function exportToExcel() {
    const ws_data = [
        ["S.No", "Name", "Age"],  // ヘッダー
        [1, "John", 25],
        [2, "Jane", 28]
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(ws_data);//データが配列の場合のシート作成
    const wb = XLSX.utils.book_new();//ワークブックの作成

    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");//ワークブックにシートを追加
    
    // Excelファイルをダウンロード
    XLSX.writeFile(wb, "output.xlsx");
}

// exportToExcel();

function overwriteExcel() {
  // 既存のExcelファイルを読み込む
  const workbook = XLSX.readFile('../doc/sample.xlsx');

  // 上書きするデータ
  const ws_data = [
      ["施設名", "所在地", "対象製品","点検実施者", "点検実施日"],
      ["東京のカフェ","東京丸の内","かつ丼", "東京さん", "2023/01/01"],
      [1,2,3,4,5]
  ];
  
  // データからワークシートを作成
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  
  // 既存のワークブックにワークシートを上書き（もしくは追加）
  workbook.Sheets['Sheet2'] = ws;

  // Excelファイルとして出力
  XLSX.writeFile(workbook, 'update.xlsx');
}

overwriteExcel();