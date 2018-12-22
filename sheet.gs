var excel_key = "1Ecr-7B_PmZKvx1sdO2xk_VF1Rj3r_dsC9cS2lfTzDmc";

var arr_project = [];
var arr_work = [];
var arr_staff = [];

function getMasterData(){

  var list_1 = Sheets.Spreadsheets.Values.get(excel_key, 'project!A:B');
  var list_2 = Sheets.Spreadsheets.Values.get(excel_key, 'staff!A:B');
  var list_3 = Sheets.Spreadsheets.Values.get(excel_key, 'work!A:B');
  
  Logger.log("getMasterData 1");
  
  for(var i=0;i < list_1.values.length; i++){
   if (i==0) continue;
   arr_project.push({
      'project_id': list_1.values[i][0],
      'project_name': list_1.values[i][1],
      'sales': list_1.values[i][2]
    });
  }  
 Logger.log("getMasterData 2");
  for(var i=0;i < list_2.values.length; i++){
   if (i==0) continue;
   arr_staff.push({
      'staff_id': list_2.values[i][0],
      'staff_name': list_2.values[i][1],
      'sales': list_2.values[i][2]
    });
  }  
 
Logger.log("getMasterData 3"); 
  for(var i=0;i < list_3.values.length; i++){
   if (i==0) continue;
   arr_work.push({
      'work_id': list_3.values[i][0],
      'work_name': list_3.values[i][1],
      'sales': list_3.values[i][2]
    });
  }  
  
Logger.log("getMasterData 4");   
}

function buildAddOn(e) {
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  
// マスタデータの取得
  getMasterData();
  
//  プロジェクトリストを取得
//  var project_list = ProjectListData();
//  var cards = []; 
//  if (project_list.length > 0) {
//    project_list.forEach(function(project_data) {
//      cards.push(BuildCard(project_data));
//    });
//  } else {
//    cards.push(CardService.newCardBuilder()
//                          .setHeader(CardService.newCardHeader()
//                                                .setTitle('No sheet data.')).build());
//  }

  var cards = []; 
  if (arr_project.length > 0) {
    arr_project.forEach(function(project_data) {
      cards.push(BuildCard(project_data));
    });
  } else {
    cards.push(CardService.newCardBuilder()
                          .setHeader(CardService.newCardHeader()
                                                .setTitle('No sheet data.')).build());
  }

  return cards;
} 

function ProjectData(project_data_id){

  Logger.log("ProjectData 2");

  Logger.log("project_data_id:"+project_data_id);
  
  var start_step = 'd'+ project_data_id;
  var end_step = 'h'+ project_data_id;
  var range = 'sheet2!'+ start_step + ':' + end_step;
//  var project_data = Sheets.Spreadsheets.Values.get('1DdCvhhFb-i3P3Px78sdww3qcv0o2iOpUvYMdM1gtK9M', range).values;
//  var project_data = Sheets.Spreadsheets.Values.get(excel_key, range).values;

  var project_data = Sheets.Spreadsheets.Values.get(excel_key, 'todo!A:G').values;
  
  return project_data;
}

//メールを発送機能が完成、しかしコードの整合が未完成
function BuildCard(project_data){

  Logger.log("BuildCard 1");
  
  var project_status = ProjectData(project_data.project_id);
  var card = CardService.newCardBuilder();
  var section = CardService.newCardSection();
//  var step_data = Sheets.Spreadsheets.Values.get('1DdCvhhFb-i3P3Px78sdww3qcv0o2iOpUvYMdM1gtK9M', 'sheet2!d1:h1').values;
  var step_data = Sheets.Spreadsheets.Values.get(excel_key, 'sheet2!d1:h1').values;
  var row_number = (project_data.project_id).toString();
  var checkboxGroup = CardService.newSelectionInput()
                                   .setType(CardService.SelectionInputType.CHECK_BOX)
                                   .setFieldName('check_box')
                                   .setOnChangeAction(CardService.newAction()
                                                                 .setFunctionName("StatusChange")
                                                                 .setParameters({row_number:row_number})); 
  for (var i = 0; i < step_data[0].length; i++) {
    var name = step_data[0][i];
    if(project_status == undefined){
      checkboxGroup.addItem(name, name, false);
    } else if(project_status[0][i] == 'Done'){
      checkboxGroup.addItem(name, name, true);
    } else {
      checkboxGroup.addItem(name, name, false);
    }
  }
  
  var text1 = step_data[0][0] +'のテンプレートを呼び出す';
  var composeAction1 = CardService.newAction()
                                 .setFunctionName('SendEmail')
                                 .setParameters({row_number:'4'});
  var send_email1 = CardService.newTextButton()
                              .setText(text1)
                              .setComposeAction(composeAction1, CardService.ComposedEmailType.REPLY_AS_DRAFT);  
                              
  var text2 = step_data[0][1] +'のテンプレートを呼び出す';
  var composeAction2 = CardService.newAction()
                                 .setFunctionName('SendEmail')
                                 .setParameters({row_number:'5'});
  var send_email2 = CardService.newTextButton()
                              .setText(text2)
                              .setComposeAction(composeAction2, CardService.ComposedEmailType.REPLY_AS_DRAFT);   
  var text3 = step_data[0][2] +'のテンプレートを呼び出す';
  var composeAction3 = CardService.newAction()
                                 .setFunctionName('SendEmail')
                                 .setParameters({row_number:'6'});
  var send_email3 = CardService.newTextButton()
                              .setText(text3)
                              .setComposeAction(composeAction3, CardService.ComposedEmailType.REPLY_AS_DRAFT); 
  var text4 = step_data[0][3] +'のテンプレートを呼び出す';
   var composeAction4 = CardService.newAction()
                                 .setFunctionName('SendEmail')
                                 .setParameters({row_number:'7'});
  var send_email4 = CardService.newTextButton()
                              .setText(text4)
                              .setComposeAction(composeAction4, CardService.ComposedEmailType.REPLY_AS_DRAFT); 
  var text5 = step_data[0][4] +'のテンプレートを呼び出す';
   var composeAction5 = CardService.newAction()
                                 .setFunctionName('SendEmail')
                                 .setParameters({row_number:'8'});
  var send_email5 = CardService.newTextButton()
                              .setText(text5)
                              .setComposeAction(composeAction5, CardService.ComposedEmailType.REPLY_AS_DRAFT);                                                           
  var button = CardService.newTextButton()
    .setText("SPREADSHEETを開く")
    .setOpenLink(CardService.newOpenLink()
//        .setUrl("https://docs.google.com/spreadsheets/d/1DdCvhhFb-i3P3Px78sdww3qcv0o2iOpUvYMdM1gtK9M")
        .setUrl("https://docs.google.com/spreadsheets/d/1Ecr-7B_PmZKvx1sdO2xk_VF1Rj3r_dsC9cS2lfTzDmc")
        .setOpenAs(CardService.OpenAs.OVERLAY)
        .setOnClose(CardService.OnClose.RELOAD_ADD_ON));
  var composeAction = CardService.newAction()
      .setFunctionName('createReplyDraft');
  var composeButton = CardService.newTextButton()
      .setText('Compose Reply')
      .setComposeAction(composeAction, CardService.ComposedEmailType.REPLY_AS_DRAFT);
      
  section.addWidget(checkboxGroup);
  section.addWidget(send_email1);
  section.addWidget(send_email2);
  section.addWidget(send_email3);
  section.addWidget(send_email4);
  section.addWidget(send_email5);
  section.addWidget(button);
  //section.addWidget(composeAction);
  section.addWidget(composeButton);
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
//    var result = Sheets.Spreadsheets.Values.update(valueRange, '1DdCvhhFb-i3P3Px78sdww3qcv0o2iOpUvYMdM1gtK9M', range, {
   var result = Sheets.Spreadsheets.Values.update(valueRange, excel_key, range, {
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
    var valueRange = Sheets.newValueRange();
    valueRange.values = values;
//    excel_key
//    var result = Sheets.Spreadsheets.Values.update(valueRange, '1DdCvhhFb-i3P3Px78sdww3qcv0o2iOpUvYMdM1gtK9M', range, {
    var result = Sheets.Spreadsheets.Values.update(valueRange, excel_key, range, {
    valueInputOption: 'RAW'}); 
  }
}

function SendEmail(e){
  var row_number = e.parameters.row_number;
  var range = 'sheet1!b'+row_number+':d'+row_number;
//  var send_data = Sheets.Spreadsheets.Values.get('1DdCvhhFb-i3P3Px78sdww3qcv0o2iOpUvYMdM1gtK9M', range).values;
  var send_data = Sheets.Spreadsheets.Values.get(excel_key, range).values;
  var mail = send_data[0][0];
  var subject = send_data[0][1];
  var main_body = send_data[0][2];
  var draft = GmailApp.createDraft(mail,subject,main_body);
  return CardService.newComposeActionResponseBuilder()
                    .setGmailDraft(draft).build();
}

function logNamesAndMajors() {
//  var spreadsheetId = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms';
//  var spreadsheetId = '1Ecr-7B_PmZKvx1sdO2xk_VF1Rj3r_dsC9cS2lfTzDmc';
  var rangeName = 'Class Data!A2:E';
//  var values = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName).values;
  var values = Sheets.Spreadsheets.Values.get(excel_key, rangeName).values;
  if (!values) {
    Logger.log('No data found.');
  } else {
    Logger.log('Name, Major:');
    for (var row = 0; row < values.length; row++) {
      // Print columns A and E, which correspond to indices 0 and 4.
      Logger.log(' - %s, %s', values[row][0], values[row][4]);
    }
  }
}

 function createReplyDraft(e) {
//    var send_data = Sheets.Spreadsheets.Values.get('1DdCvhhFb-i3P3Px78sdww3qcv0o2iOpUvYMdM1gtK9M', 'sheet1!b4:d4').values;
    var send_data = Sheets.Spreadsheets.Values.get(excel_key, 'sheet1!b4:d4').values;
    var mail = send_data[0][0];
    var subject = send_data[0][1];
    var main_body = send_data[0][2];

    // Creates a draft reply.
    var messageId = e.messageMetadata.messageId;
    var message = GmailApp.getMessageById(messageId);
    var draft = message.createDraftReply(main_body
   
        
    );

    // Return a built draft response. This causes Gmail to present a
    // compose window to the user, pre-filled with the content specified
    // above.
    return CardService.newComposeActionResponseBuilder()
        .setGmailDraft(draft).build();
  }
