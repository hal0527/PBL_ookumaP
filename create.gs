//シートの設定・新規spreadシート すべて自分設定する
function setAddOn(e){
  var userProperties = PropertiesService.getUserProperties();
  var sheet_check = spread_sheet_id;
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var section1 = CardService.newCardSection();
  var section2 = CardService.newCardSection();
  var user_id = '01';
  if(sheet_check == undefined || null){
    sheet_check = '登録したしたシートがありません。';
  } else{
    sheet_check = '登録したシートのID<br>'　+ sheet_check;
  }
　 //新規spreadシート,シート数を選択し、シートの名前を選択する
  var sheet_name = CardService.newTextInput()
                           .setFieldName('create_sheet_name')
                           .setTitle('ファイル名を入力してください。');
  var sheet_name_memo = CardService.newTextParagraph()
                              .setText('<b><font color="#ea9999">アドオンに対応した二つシートを設定しました。</font></b><br>'
                              +'・ 対応メールの設定シート<br>'+'・ プロジェクトを管理するシート');
  var sheet1_name = CardService.newTextInput()
                           .setFieldName('sheet1_name')
                           .setTitle('メール設定のシート名を入力してください。');
  var sheet2_name = CardService.newTextInput()
                           .setFieldName('sheet2_name')
                           .setTitle('タスクのシート名を入力してください。');
  var default_id = CardService.newTextParagraph()
                                  .setText(sheet_check);
  var sheet_id = CardService.newTextInput()
                             .setFieldName('sheet_id')
                             .setTitle('シートのキーIDを入力してください。');
  var button1 =  CardService.newTextButton()
                           .setText('新規シートを作成する')
                           .setBackgroundColor('blue')
                           .setOnClickAction(CardService.newAction()
                                                        .setFunctionName("CreateSpreadSheet"));                           
  var button2 =  CardService.newTextButton()
                           .setText('シートのIDを登録する')
                           .setBackgroundColor('blue')
                           .setOnClickAction(CardService.newAction()
                                                        .setFunctionName("SaveSheetId"));
  section1.addWidget(sheet_name);
  section1.addWidget(sheet_name_memo);
  section1.addWidget(sheet1_name);
  section1.addWidget(sheet2_name);
  section1.addWidget(button1);
  section2.addWidget(default_id);
  section2.addWidget(sheet_id);
  section2.addWidget(sheet1_name);
  section2.addWidget(sheet2_name);
  section2.addWidget(button2);
  var card1 = CardService.newCardBuilder()
                        .setHeader(CardService.newCardHeader()
                                              .setTitle('シートの作成'))
                        .addSection(section1)
                        .build();
  var card2 = CardService.newCardBuilder()
                        .setHeader(CardService.newCardHeader()
                                              .setTitle('シートの設定'))
                        .addSection(section2)
                        .build();
  return CardService.newUniversalActionResponseBuilder()
                    .displayAddOnCards([card1,card2])
                    .build();
}

function CreateSpreadSheet(e){
  var spreadsheet_name = e.formInput.create_sheet_name;
  var sheet1_name = e.formInput.sheet1_name;
  var sheet2_name = e.formInput.sheet2_name;
  var sheet = Sheets.newSpreadsheet();
  sheet.properties = Sheets.newSpreadsheetProperties();
  sheet.properties.title = spreadsheet_name; 
  var spreadsheet = Sheets.Spreadsheets.create(sheet);
  var get_spreadsheet_id = spreadsheet.spreadsheetId;
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('spread_sheet_id', get_spreadsheet_id);
  userProperties.setProperty('sheet1_name', sheet1_name);
  userProperties.setProperty('sheet2_name', sheet2_name);
  //表1,2,3の作成 
  var sheet_name = [sheet1_name,sheet2_name]
  for(var i = 0;i < sheet_name.length;i++){
  var resource = {'requests':[
                              {
                                "addSheet": {
                                  "properties": {
                                    "title": sheet_name[i]
                                  }
                                }
                              }
                            ]};
  Sheets.Spreadsheets.batchUpdate(resource, get_spreadsheet_id);
  }
  //表1,2,3の内容作成 
  var values1 = [
                  ['','Mail Adress','Subject','Content'],
                  ['Instruction'],
                  ['Feed sample'],
                  ['Check implemented tags'],
                  ['FTP account']
                ];
  var range1 = sheet_name[0] + '!a1:d8';
  var value_range1 = Sheets.newValueRange();
  value_range1.values = values1;
  var result1 = Sheets.Spreadsheets.Values.update(value_range1, get_spreadsheet_id, range1, {
      valueInputOption: 'RAW'});
  var values2 = [
                  ['campaign','sales','instruction','Feed sample','Check implemented tags','FTP account']        
                ];
  var range2 = sheet_name[1] + '!a1:f1';
  var value_range2 = Sheets.newValueRange();
  value_range2.values = values2;
  var result2 = Sheets.Spreadsheets.Values.update(value_range2, get_spreadsheet_id, range2, {
      valueInputOption: 'RAW'});
  return CardService.newActionResponseBuilder()
                    .setNotification(CardService.newNotification()
                    .setType(CardService.NotificationType.INFO)
                    .setText("シートを作成成功"))
                    .setNavigation(nav)
                    .build();    
}



function SaveSheetId(e){
  var sheet1_name = e.formInput.sheet1_name;
  var sheet2_name = e.formInput.sheet2_name;
  var sheet_id = e.formInput.sheet_id;
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('spread_sheet_id', sheet_id);
  userProperties.setProperty('sheet1_name', sheet1_name);
  userProperties.setProperty('sheet2_name', sheet2_name);

  return CardService.newActionResponseBuilder()
                    .setNotification(CardService.newNotification()
                    .setType(CardService.NotificationType.INFO)
                    .setText("シートを登録成功"))
                    .setNavigation(nav)
                    .build();    
  
}