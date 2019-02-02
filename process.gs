var userProperties = PropertiesService.getUserProperties();
var nav = CardService.newNavigation().popToRoot();
var spread_sheet_id = userProperties.getProperty('spread_sheet_id');
var sheet1 = userProperties.getProperty('sheet1_name');
var sheet2 = userProperties.getProperty('sheet2_name');
if(spread_sheet_id == null || undefined){
  spread_sheet_id = '1M05QfoasPgJpBC4CL58SVp6zx-MrhqsSSMyV735otvQ';
  var sheet1 = 'templates';
  var sheet2 = 'campaigns';
}
Logger.log(spread_sheet_id);

function buildAddOn(e) {
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var project_list = ProjectListData();
  var cards = []; 
  if (project_list.length > 0) {
   cards.push(CardService.newCardBuilder()
                          .setHeader(CardService.newCardHeader()
                                                .setTitle('---進行中のプロジェクト---'))
                          .build());
    project_list.forEach(function(project_data) {
       if(BuildProjectCard(project_data) == 0){
         
       } else {
         cards.push(BuildProjectCard(project_data));
       }
    });
  } else {
    cards.push(CardService.newCardBuilder()
                          .setHeader(CardService.newCardHeader()
                                                .setTitle('No sheet data.')).build());
  }
  return cards;
} 
  //データ保存
function ProjectListData(){
  var range = sheet2 + '!a2:b30';
  var project_lists = Sheets.Spreadsheets.Values.get(spread_sheet_id, range);
  var recents = [];
  if(project_lists.values == undefined || null){
    recents.push({
      'id': 'none',
      'project_name': 'none',
      'sales': 'none'
    });
  } else {
       for(var i=0;i < project_lists.values.length; i++){
          recents.push({
            'id': i+2,
            'project_name': project_lists.values[i][0],
            'sales': project_lists.values[i][1]
          });
        }
  }
   return recents;
}

function ProjectData(project_data_id){
  var start_step = 'c'+ project_data_id;
  var end_step = 'g'+ project_data_id;
  if(project_data_id == 'none'){
    var project_data = [[]];
  }else {
    var range = sheet2 + '!'+ start_step + ':' + end_step;
    var project_data = Sheets.Spreadsheets.Values.get(spread_sheet_id, range).values;
  }
  return project_data;
}

//メールを発送機能
function BuildProjectCard(project_data){
  var range = sheet2 + '!c1:z1';
  //var range = 'sheet2!c1:z1';
  var finish_status = 0;
  var project_status = ProjectData(project_data.id);
  var card = CardService.newCardBuilder();
  var section = CardService.newCardSection();
  var step_data = Sheets.Spreadsheets.Values.get(spread_sheet_id, range).values;
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
      finish_status++;
    } else {
      checkboxGroup.addItem(name, name, false);
    }
      process_name.addItem(name, name, false);
  } 

    var compose_action_1 = CardService.newAction()
                                      .setFunctionName('SendEmail')
                                      .setParameters({mail_type:'create'});
    var create_button = CardService.newTextButton()
                                   .setText('新規メールに引用する')
                                   .setComposeAction(compose_action_1, CardService.ComposedEmailType.REPLY_AS_DRAFT);
    var compose_action_2 = CardService.newAction()
                                   .setFunctionName('SendEmail')
                                   .setParameters({mail_type:'reply'});
    var reply_button = CardService.newTextButton()
                                  .setText('返信メールに引用する')
                                  .setComposeAction(compose_action_2, CardService.ComposedEmailType.REPLY_AS_DRAFT);
    var process_title1 = CardService.newKeyValue()
                                    .setIcon(CardService.Icon.DESCRIPTION)
                                    .setContent("完了項目のみチェック");
    var process_title2 = CardService.newKeyValue()
                                    .setIcon(CardService.Icon.EMAIL)
                                    .setContent("項目の確認メールを選択する");
    var process_title3 = CardService.newKeyValue()
                                    .setIcon(CardService.Icon.OFFER)
                                    .setContent("シートを管理する")
                                    .setOpenLink(CardService.newOpenLink()
                                                          .setUrl("https://docs.google.com/spreadsheets/d/" + spread_sheet_id)
                                                          .setOpenAs(CardService.OpenAs.OVERLAY)
                                                          .setOnClose(CardService.OnClose.RELOAD_ADD_ON));      
    if(project_data.id == 'none'){
      var title = '*タスクがありません*';
      card.setHeader(CardService.newCardHeader().setTitle(title));
      return card.build();
    } else {
      var title = project_data.project_name + '(' +finish_status + '/'+ step_data[0].length + ') 担当者：' + project_data.sales;
    }                                                      
    section.addWidget(process_title1);
    section.addWidget(checkboxGroup);
    section.addWidget(process_title2);
    section.addWidget(process_name);
    section.addWidget(create_button);
    section.addWidget(reply_button);
    section.addWidget(process_title3);
    card.addSection(section);
    card.setHeader(CardService.newCardHeader().setTitle(title));
    if(finish_status == step_data[0].length){
      return 0;
    } else {
      return card.build();
    }
  
}

//進捗完成されば、タスク
function StatusChange(e){
  var checked_group = e.formInputs.check_box;
  var row_number = e.parameters.row_number;
  var line_number;
  Logger.log(e);
  var arr1 = [];
  var arr2 = [];
  for(var i = 99; i < 104; i++){
        arr1.push(String.fromCharCode(i));
  }
//自動化不足
  for(var i = 0; i < checked_group.length; i++){
    var step_name = checked_group[i];
    switch (step_name)
    {  case "導入説明":
            line_number = "c";
            break;
        case "アカウント発行依頼":
            line_number = "d";
            break;
        case "サイト解析依頼":
            line_number = "e";
            break;
        case "実装確認依頼":
            line_number = "f";
            break;
        case "予算設定依頼":
            line_number = "g";
            break; 
    }
    arr2.push(line_number);
    var range = sheet2 + '!' + line_number + row_number;
    //var range = 'sheet2!' + line_number + row_number;
    var values = [
                    [
                      'Done'
                    ]
                 ];
    var value_range = Sheets.newValueRange();
    value_range.values = values;
    var result = Sheets.Spreadsheets.Values.update(value_range, spread_sheet_id, range, {
    valueInputOption: 'RAW'}); 
  }
  var different = arr2.concat(arr1).filter(function (v) {
                return arr2.indexOf(v)===-1 || arr1.indexOf(v)===-1
            });
  for(var i=0; i < different.length; i++){
    var range = sheet2 + '!' + different[i] + row_number;
    var values = [
                    [
                       ''
                    ]
                 ];
    var value_range = Sheets.newValueRange();
    value_range.values = values;
    Logger.log(different);
    var result = Sheets.Spreadsheets.Values.update(value_range, spread_sheet_id, range, {
    valueInputOption: 'RAW'}); 
  }
}

//自動化不足
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
 var range = sheet1 + '!b' + row_number + ':d' + row_number;
 var model_message = Sheets.Spreadsheets.Values.get(spread_sheet_id, range).values; 
 var mail = model_message[0][0];
 var subject = model_message[0][1].toString();
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
  var range = sheet1 + '!b'+row_number+':d'+row_number;
  var send_data = Sheets.Spreadsheets.Values.get(spread_sheet_id, range).values;
  var mail = send_data[0][0];
  var subject = send_data[0][1];
  var main_body = send_data[0][2];
  var draft = GmailApp.createDraft(mail,subject,main_body);
  return CardService.newComposeActionResponseBuilder()
                    .setGmailDraft(draft).build();

}
