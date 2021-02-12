const api_key = "自分のAPIキーを入力";
const spreadsheet = SpreadsheetApp.openById('スプレッドシートのIDを入力');
const sheet1 = spreadsheet.getSheetByName('検索結果');
const setting_sheet = spreadsheet.getSheetByName('設定');

function productDescribe(asin_list) {
  const asin = asin_list.join();
  const endpoint = 'https://api.keepa.com/product?domain=5&key=';
  const stats = '90';
  const response = UrlFetchApp.fetch(endpoint + api_key + '&asin=' + asin + '&stats=' + stats);
  const content = response.getContentText();
  const json = JSON.parse(content);
  // if (json.error) {
  //   throw new Error(json.error.message); //エラーが発生した場合、レスポンスは"error"というキーを含む
  // }
  const products = json.products;
  const ind = { AMAZON: 0, NEW: 1, SALES: 3, RATING: 16, COUNT_REVIEWS: 17, BUY_BOX_SHIPPING: 18 };　// CSVの行定義



  products.forEach((product, i) => {
    const startrow = 2;
    sheet1.getRange(startrow + i, 1).setValue(product.asin);
    sheet1.getRange(startrow + i, 2).setValue(product.title);
    sheet1.getRange(startrow + i, 3).setValue('=IMAGE("https://images-na.ssl-images-amazon.com/images/I/' + product.imagesCSV.split(',')[0] + '")');
    sheet1.getRange(startrow + i, 4).setValue(product.eanList);
    sheet1.getRange(startrow + i, 5).setValue(product.csv[ind['NEW']][product.csv[ind['NEW']].length - 1]);
    sheet1.getRange(startrow + i, 6).setValue(product.stats.salesRankDrops30);
    sheet1.getRange(startrow + i, 7).setValue(product.stats.salesRankDrops90);
    sheet1.getRange(startrow + i, 8).setValue('=HYPERLINK("https://www.amazon.co.jp/dp/' + product.asin + '","' + product.asin + '")');
    sheet1.getRange(startrow + i, 9).setValue('=HYPERLINK("https://keepa.com/#!product/5-' + product.asin + '","' + product.asin + '")');
    if (Array.isArray(product.eanList)) {
      if (product.eanList[0] != null) {
      sheet1.getRange(startrow + i, 10).setValue('=HYPERLINK("https://www.google.com/search?q=' + product.eanList[0] + '","' + product.eanList[0] + '")');
      };
      if (product.eanList[1] != null) {
      sheet1.getRange(startrow + i, 11).setValue('=HYPERLINK("https://www.google.com/search?q=' + product.eanList[1] + '","' + product.eanList[1] + '")');
      };
    };
  });
};

function productFinder(json) {
  const request = UrlFetchApp.fetch('https://api.keepa.com/query?domain=5&key=' + api_key,{
        'method': 'POST',
        'headers' : {
          'Connection' : 'keep-alive'
        },
        'payload': {'selection': JSON.stringify(json)}
      });
  const response = JSON.parse(request.getContentText());
  const asin_list = response.asinList;
  return asin_list;
};

function read_setting() {
  const data = setting_sheet.getDataRange().getValues();
  let delete_row = [0];    //削除したい行
  for(var i=0; i<delete_row.length; i++){
    data.splice(delete_row[i]-i, 1);
  };
  let delete_column = [0];  // 削除したい列
  for(let i=0; i<data.length; i++){    //このfor文で行を回す
    for(let j=0; j<delete_column.length; j++){
      data[i].splice(delete_column[j]-j, 1);
    }
  }
  return data;
};

function retrieveTokenStatus() {
  const request = UrlFetchApp.fetch('https://api.keepa.com/token?key=' + api_key);
  const response = JSON.parse(request.getContentText());
  Logger.log(response);
  Browser.msgBox(
    'トークン情報','現在のトークン数：' + response.tokensLeft + 
    '\\n１分間あたり回復するトークン数：' + response.refillRate + 
    '\\nトークン回復まで残り：' + Math.round(response.refillIn / 1000) + '秒', 
    Browser.Buttons.OK);
  return;
}

function main(){
  const confirmation = Browser.msgBox("確認","情報を取得してもよろしいですか？", Browser.Buttons.OK_CANCEL);
  if(confirmation == "cancel") {
    Browser.msgBox("操作をキャンセルしました。スクリプトを停止します。");
    exit;
  };
  SpreadsheetApp.getActiveSpreadsheet().toast('スクリプトを実行中', '実行中', 0);
  const extract = read_setting();
  const json_extract = extract.reduce((extract_association, [key, value]) => Object.assign(extract_association, {[key]: value}), {});
  const json_other = {
    "sort": [["current_SALES","asc"]],
    "productType": [0, 1]
  };
  const json_data = Object.assign(json_extract, json_other);
  sheet1.deleteRows(2, sheet1.getLastRow());
  const asin_list = productFinder(json_data);
  productDescribe(asin_list);
  SpreadsheetApp.getActiveSpreadsheet().toast('スクリプトの実行が終わりました', '実行終了', 1);
};

