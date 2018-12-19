var spread_sheet_id = '1DdCvhhFb-i3P3Px78sdww3qcv0o2iOpUvYMdM1gtK9M';
function buildAddOn(e) {
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var project_list = ProjectListData();
  var cards = []; 
  if (project_list.length > 0) {
    project_list.forEach(function(project_data) {
      cards.push(BuildCard(project_data));
    });
  } else {
    cards.push(CardService.newCardBuilder()
                          .setHeader(CardService.newCardHeader()
                                                .setTitle('No sheet data.')).build());
  }
  return cards;
} 
//データ保存、必要がないかも
function ProjectListData(){
  var project_lists = Sheets.Spreadsheets.Values.get(spread_sheet_id, 'sheet2!b2:c20');
  var recents = [];
  for(var i=0;i < project_lists.values.length; i++){
    recents.push({
      'id': i+2,
      'project_name': project_lists.values[i][0],
      'sales': project_lists.values[i][1]
    });
  }
   return recents;
}

function ProjectData(project_data_id){
  var start_step = 'd'+ project_data_id;
  var end_step = 'h'+ project_data_id;
  var range = 'sheet2!'+ start_step + ':' + end_step;
  var project_data = Sheets.Spreadsheets.Values.get(spread_sheet_id, range).values;
  return project_data;
}

//メールを発送機能が完成、しかしコードの整合が未完成
function BuildCard(project_data){
  var project_status = ProjectData(project_data.id);
  var card = CardService.newCardBuilder();
  var section = CardService.newCardSection();
  var step_data = Sheets.Spreadsheets.Values.get(spread_sheet_id, 'sheet2!d1:h1').values;
  var row_number = (project_data.id).toString();
  var checkboxGroup = CardService.newSelectionInput()
                                   .setType(CardService.SelectionInputType.CHECK_BOX)
                                   .setFieldName('check_box')
                                   .setOnChangeAction(CardService.newAction()
                                                                 .setFunctionName("StatusChange")
                                                                 .setParameters({row_number:row_number}));  
  var process_name = CardService.newSelectionInput()
                                   .setType(CardService.SelectionInputType.DROPDOWN)
                                   .setFieldName('process_name'); 
  for (var i = 0; i < step_data[0].length; i++) {
    var name = step_data[0][i];
    if(project_status == undefined){
      checkboxGroup.addItem(name, name, false);
    } else if(project_status[0][i] == 'Done'){
      checkboxGroup.addItem(name, name, true);
    } else {
      checkboxGroup.addItem(name, name, false);
    }
      process_name.addItem(name, name, false);
  }                                                                  
  var compose_action_1 = CardService.newAction()
                                 .setFunctionName('SendEmail')
                                 .setParameters({mail_type:'create'});
  var create_button = CardService.newTextButton()
                                .setText('新規メールで呼び出す')
                                .setComposeAction(compose_action_1, CardService.ComposedEmailType.REPLY_AS_DRAFT);
  var compose_action_2 = CardService.newAction()
                                 .setFunctionName('SendEmail')
                                 .setParameters({mail_type:'reply'});
  var reply_button = CardService.newTextButton()
                                .setText('返信メールで呼び出す')
                                .setComposeAction(compose_action_2, CardService.ComposedEmailType.REPLY_AS_DRAFT);
  var sheet_button = CardService.newTextButton()
    .setText("SPREADSHEETを開く")
    .setOpenLink(CardService.newOpenLink()
        .setUrl("https://docs.google.com/spreadsheets/d/1DdCvhhFb-i3P3Px78sdww3qcv0o2iOpUvYMdM1gtK9M")
        .setOpenAs(CardService.OpenAs.OVERLAY)
        .setOnClose(CardService.OnClose.RELOAD_ADD_ON));
  var process_title = CardService.newKeyValue()
    .setIconUrl("https://icon.png")
    .setContent("SELECT")
    .setButton(sheet_button);
  section.addWidget(process_title);
  section.addWidget(checkboxGroup);
  section.addWidget(process_name);
  section.addWidget(create_button);
  section.addWidget(reply_button);
  card.addSection(section);
  card.setHeader(CardService.newCardHeader().setTitle(project_data.project_name));
  return card.build();
}

//checkbox 確認、sheet入力完成　しかしcheckoff機能なし
function StatusChange(e){
  var checked_group = e.formInputs.check_box;
  var row_number = e.parameters.row_number;
  var line_number;
  Logger.log(e);
  var arr1 = [];
  var arr2 = [];
  for(var i = 100; i < 105; i++){
        arr1.push(String.fromCharCode(i));
  }

  for(var i = 0; i < checked_group.length; i++){
    var step_name = checked_group[i];
    switch (step_name)
    {  case "導入説明":
            line_number = "d";
            break;
        case "アカウント発行依頼":
            line_number = "e";
            break;
        case "サイト解析依頼":
            line_number = "f";
            break;
        case "実装確認依頼":
            line_number = "g";
            break;
        case "予算設定依頼":
            line_number = "h";
            break; 
    }
    arr2.push(line_number);
    Logger.log(line_number);
    var range = 'sheet2!' + line_number + row_number;
    var values = [
                    [
                      'Done'
                    ]
                 ];
    var valueRange = Sheets.newValueRange();
    valueRange.values = values;
    var result = Sheets.Spreadsheets.Values.update(valueRange, spread_sheet_id, range, {
    valueInputOption: 'RAW'}); 
  }
  var different = arr2.concat(arr1).filter(function (v) {
                return arr2.indexOf(v)===-1 || arr1.indexOf(v)===-1
            });
  for(var i=0; i < different.length; i++){
    var range = 'sheet2!' + different[i] + row_number;
    var values = [
                    [
                       ''
                    ]
                 ];
    var value_range = Sheets.newValueRange();
    valueRange.values = values;
    var result = Sheets.Spreadsheets.Values.update(value_range, spread_sheet_id, range, {
    valueInputOption: 'RAW'}); 
  }
}
function SendEmail(e){
 var process_name = e.formInput.process_name;
 var mail_type = e.parameters.mail_type;
 var row_number;
 switch (process_name)
    {  case "導入説明":
            row_number = "4";
            break;
        case "アカウント発行依頼":
            row_number = "5";
            break;
        case "サイト解析依頼":
            row_number = "6";
            break;
        case "実装確認依頼":
            row_number = "7";
            break;
        case "予算設定依頼":
            row_number = "8";
            break; 
    }
 var range = 'sheet1!b' + row_number + ':d' + row_number;
 var model_message = Sheets.Spreadsheets.Values.get(spread_sheet_id, range).values; 
 var mail = model_message[0][0];
 var subject = model_message[0][1];
 var main_body = model_message[0][2];
 if(mail_type == 'create'){
    var draft = GmailApp.createDraft(mail,subject,main_body);
    return CardService.newComposeActionResponseBuilder()
                      .setGmailDraft(draft).build();
 } else if(mail_type == 'reply'){
    var messageId = e.messageMetadata.messageId;
    var message = GmailApp.getMessageById(messageId);
    var draft = message.createDraftReply(main_body);
    return CardService.newComposeActionResponseBuilder()
                      .setGmailDraft(draft).build();
 }
}
function SendEmail1(e){
  var row_number = e.parameters.row_number;
  var range = 'sheet1!b'+row_number+':d'+row_number;
  var send_data = Sheets.Spreadsheets.Values.get(spread_sheet_id, range).values;
  var mail = send_data[0][0];
  var subject = send_data[0][1];
  var main_body = send_data[0][2];
  var draft = GmailApp.createDraft(mail,subject,main_body);
  return CardService.newComposeActionResponseBuilder()
                    .setGmailDraft(draft).build();

}

//１プロジェクト追加・２タスクの分配
function sheetAddOn(e){
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var section1 = CardService.newCardSection();
  var section2 = CardService.newCardSection();
  var sheet_id = PropertiesService.getUserProperties().getProperty('sheet_id');
  var projects_list = Sheets.Spreadsheets.Values.get(spread_sheet_id, 'sheet2!b2:b').values;
  var sales_list = Sheets.Spreadsheets.Values.get(spread_sheet_id, 'sheet3!b2:b').values;
  Logger.log(projects_list);
  Logger.log(sales_list);

 //プロジェクト追加
  var project_name = CardService.newTextInput()
                           .setFieldName('project_name')
                           .setTitle('プロジェクト名を入力してください。');
  var sales_name = CardService.newTextInput()
                             .setFieldName('sales_name')
                             .setTitle('担当者の名前を入力してください。');
  var button1 =  CardService.newTextButton()
                           .setText('確認する')
                           .setBackgroundColor('blue')
                           .setOnClickAction(CardService.newAction()
                                                        .setFunctionName("AddProject"));
  section1.addWidget(project_name);
  section1.addWidget(sales_name);
  section1.addWidget(button1);
  //タスクの分配
  var projects = CardService.newSelectionInput()
                         .setType(CardService.SelectionInputType.DROPDOWN)
                         .setFieldName('projects');
  for(var i = 0; i < projects_list.length;i++){
    projects.addItem(projects_list[i][0], projects_list[i][0], false);
  }
  var sales = CardService.newSelectionInput()
                         .setType(CardService.SelectionInputType.DROPDOWN)
                         .setFieldName('sales');
  for(var i = 0; i < sales_list.length;i++){
    sales.addItem(sales_list[i][0], sales_list[i][0],false);
  }
  var button2 =  CardService.newTextButton()
                           .setText('確認する')
                           .setBackgroundColor('blue')
                           .setOnClickAction(CardService.newAction()
                                                        .setFunctionName("GiveProject"));
  section2.addWidget(projects);
  section2.addWidget(sales);
  section2.addWidget(button2);
  
  //cardの作成
  var card1 = CardService.newCardBuilder()
                        .setHeader(CardService.newCardHeader()
                                              .setTitle('プロジェクトの追加'))
                        .addSection(section1)
                        .build();
  var card2 = CardService.newCardBuilder()
                        .setHeader(CardService.newCardHeader()
                                              .setTitle('タスクの分配'))
                        .addSection(section2)
                        .build();
  return CardService.newUniversalActionResponseBuilder()
                    .displayAddOnCards([card1,card2])
                    .build();
  
}
function AddProject(e){
  var project_name = e.formInput.project_name;
  var worker_name = e.formInput.sales_name;
  var nav = CardService.newNavigation().popToRoot();
  var project_num = Sheets.Spreadsheets.Values.get(spread_sheet_id, 'sheet2!B2:B').values.length;
  var row_num = Number(project_num) + 2;
  var range = 'sheet2!B' + row_num + ':C' + row_num;
  var values = [
    [project_name,worker_name]
  ];
  var valueRange = Sheets.newValueRange();
  valueRange.values = values;
  var result = Sheets.Spreadsheets.Values.update(valueRange, spread_sheet_id, range, {
    valueInputOption: 'RAW'
  });
  return CardService.newActionResponseBuilder()
                    .setNotification(CardService.newNotification()
                    .setType(CardService.NotificationType.INFO)
                    .setText("プロジェクト追加成功"))
                    .setNavigation(nav)
                    .build();

}

function GiveProject(){
  
}

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
