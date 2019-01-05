//シートの設定・新規spreadシート すべて自分設定する
function setAddOn(e){
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var section1 = CardService.newCardSection();
  var section2 = CardService.newCardSection();
  var user_id = '01';
  var sheet_check = spread_sheet_id;
  //var sheet_check = PropertiesService.getUserProperties().getProperty('sheet_id');
  if(spread_sheet_id == undefined || null){
    sheet_check = '連携したシートがありません。';
  } else{
    sheet_check = '連携したシートがありました。';
  }
　 //新規spreadシート,シート数を選択し、シートの名前を選択する
  var sheet_name = CardService.newTextInput()
                           .setFieldName('create_sheet_name')
                           .setTitle('シート名を入力してください。');
  var default_id = CardService.newTextParagraph()
                              .setText(sheet_check);
  var user_name = CardService.newTextInput()
                             .setFieldName('user_name')
                             .setTitle('ユーザー名を入力してください。');
  var sheet_id = CardService.newTextInput()
                             .setFieldName('sheet_id')
                             .setTitle('シートのキーIDを入力してください。');
  var button1 =  CardService.newTextButton()
                           .setText('新規シートを作成する')
                           .setBackgroundColor('blue')
                           .setOnClickAction(CardService.newAction()
                                                        .setFunctionName("CreateSpreadSheet"));                           
  var button2 =  CardService.newTextButton()
                           .setText('ユーザーとIDを登録する')
                           .setBackgroundColor('blue')
                           .setOnClickAction(CardService.newAction()
                                                        .setFunctionName("SaveSheetId"));
  section1.addWidget(sheet_name);
  section1.addWidget(button1);
  section2.addWidget(default_id);
  section2.addWidget(user_name);
  section2.addWidget(sheet_id);
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
  var sheet_name = e.formInput.create_sheet_name;
  var sheet = Sheets.newSpreadsheet();
  sheet.properties = Sheets.newSpreadsheetProperties();
  sheet.properties.title = sheet_name; 
  var spreadsheet = Sheets.Spreadsheets.create(sheet);
  var get_spreadsheet_id = spreadsheet.spreadsheetId;
  var resource = {
      destinationSpreadsheetId: get_spreadsheet_id
    }
  var copy_spreadsheet = Sheets.Spreadsheets.Sheets.copyTo(resource, '1qGuIxLQJ0E0T2WjxpBxnfN_MxV8g1q0VDxQYVGM_7Wk', '0');
  var copy_spreadsheet = Sheets.Spreadsheets.Sheets.copyTo(resource, '1qGuIxLQJ0E0T2WjxpBxnfN_MxV8g1q0VDxQYVGM_7Wk', '143212783');
  var copy_spreadsheet = Sheets.Spreadsheets.Sheets.copyTo(resource, '1qGuIxLQJ0E0T2WjxpBxnfN_MxV8g1q0VDxQYVGM_7Wk', '1896783048');
  //1qGuIxLQJ0E0T2WjxpBxnfN_MxV8g1q0VDxQYVGM_7Wk
}

function SaveSheetId(e){
  var user_name = e.formInput.user_name;
  var user_email_address = Session.getActiveUser().getEmail();
  //var sheet_id = e.formInput.sheet_id;
  //var userProperties = PropertiesService.getUserProperties();
  //userProperties.setProperty('spread_sheet_id', sheet_id);
  var user_list = Sheets.Spreadsheets.Values.get(spread_sheet_id, 'sheet3!B2:B').values;
  var user_number = user_list.length + 1;
  var row_number = user_number + 1;
  var address_list = Sheets.Spreadsheets.Values.get(spread_sheet_id, 'sheet3!C2:C').values;
  var add_range = "Sheet3!A"+ row_number + ":D" + row_number;
  var user_login = {
      "values": [
        [user_number, user_name, user_email_address, "user"]
      ]
   };
   var add_user = Sheets.Spreadsheets.Values.update(user_login, spread_sheet_id, add_range, {
    valueInputOption: 'RAW'
  });
  
}