//シートの設定・新規spreadシート
function setAddOn(e){
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var section1 = CardService.newCardSection();
  var section2 = CardService.newCardSection();
  var user_id = '01';
  var sheet_check = PropertiesService.getUserProperties().getProperty('sheet_id');
  if(sheet_check !== undefined || null){
    var sheet_check = '今シートの連携がありません。';
  }
　 //新規spreadシート,シート数を選択し、シートの名前を選択する
  var sheet_name = CardService.newTextInput()
                           .setFieldName('sheet_name')
                           .setTitle('シート名を入力してください。');
  
                           
  var default_id = CardService.newTextParagraph()
                              .setText(sheet_check);
  var user_name = CardService.newTextInput()
                             .setFieldName('user_name')
                             .setTitle('input the name');
  var sheet_id = CardService.newTextInput()
                             .setFieldName('sheet_id')
                             .setTitle('input the id of sheet');
  var button =  CardService.newTextButton()
                           .setText('確認する')
                           .setBackgroundColor('blue')
                           .setOnClickAction(CardService.newAction()
                                                        .setFunctionName("SaveSheetId"));
  section2.addWidget(default_id);
  section2.addWidget(user_name);
  section2.addWidget(sheet_id);
  section2.addWidget(button);
  var card1 = CardService.newCardBuilder()
                        .setHeader(CardService.newCardHeader()
                                              .setTitle('シートの作成'))
                        //.addSection(section1)
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

function createSpreadSheet(){
  var sheet = Sheets.newSpreadsheet();
  sheet.properties = Sheets.newSpreadsheetProperties();
  sheet.properties.title = title;
  var spreadsheet = Sheets.Spreadsheets.create(sheet);
}

function SaveSheetId(e){
  var user_name = e.formInput.user_name;
  var user_email_address = Session.getActiveUser().getEmail();
  var sheet_id = e.formInput.sheet_id;
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('spread_sheet_id', sheet_id);
  
}

