//１プロジェクト追加・２タスクの分配
function sheetAddOn(e){
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var section1 = CardService.newCardSection();
  var section2 = CardService.newCardSection();
  var sheet_id = PropertiesService.getUserProperties().getProperty('sheet_id');
  var range_pro = sheet2 + '!A2:A'; 
  var projects_list = Sheets.Spreadsheets.Values.get(spread_sheet_id, range_pro).values;
 
 //プロジェクト追加
  var project_name = CardService.newTextInput()
                           .setFieldName('project_name')
                           .setTitle('プロジェクト名を入力してください。');
  var sales_name = CardService.newTextInput()
                              .setTitle('担当者を入力してください。')
                              .setFieldName('sales_name'); 
  var button1 =  CardService.newTextButton()
                           .setText('プロジェクトを作成する')
                           .setBackgroundColor('blue')
                           .setOnClickAction(CardService.newAction()
                                                        .setFunctionName("AddProject"));
  //タスクの分配
  var projects = CardService.newSelectionInput()
                         .setType(CardService.SelectionInputType.DROPDOWN)
                         .setTitle('プロジェクトを選択してください。')
                         .setFieldName('projects');
  if(projects_list == undefined || null){
    projects.addItem('', '', false);
  } else {                       
    for(var i = 0; i < projects_list.length;i++){
      var project_range = sheet2 + '!B'+(i+2);
      projects.addItem(projects_list[i][0], project_range, false);
    }
  }
  var sales = CardService.newTextInput()
                         .setTitle('担当者を入力してください。')
                         .setFieldName('sales');
  
  var button2 =  CardService.newTextButton()
                           .setText('変更する')
                           .setBackgroundColor('blue')
                           .setOnClickAction(CardService.newAction()
                                                        .setFunctionName("GiveProject"));
  section1.addWidget(project_name);
  section1.addWidget(sales_name);
  section1.addWidget(button1);                                                      
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
  var range_pro = sheet2 + '!A2:A';
  var project_num = Sheets.Spreadsheets.Values.get(spread_sheet_id, range_pro).values;
  if(project_num == undefined || null){
    var row_num = 2;
  } else {
    var row_num = Number(project_num.length) + 2;
  }
  var range = sheet2 + '!A' + row_num + ':B' + row_num;
  var values = [
    [project_name,worker_name]
  ];
  var value_range = Sheets.newValueRange();
  value_range.values = values;
  var result = Sheets.Spreadsheets.Values.update(value_range, spread_sheet_id, range, {
    valueInputOption: 'RAW'
  });
  return CardService.newActionResponseBuilder()
                    .setNotification(CardService.newNotification()
                    .setType(CardService.NotificationType.INFO)
                    .setText("プロジェクト追加成功"))
                    .setNavigation(nav)
                    .build();

}

function GiveProject(e){
  var project_range = e.formInput.projects;
  var sales_name = e.formInput.sales;
  var values = [
                [sales_name]
               ];
  var value_range = Sheets.newValueRange();
  value_range.values = values;
  
  var result = Sheets.Spreadsheets.Values.update(value_range, spread_sheet_id, project_range, {
  valueInputOption: 'RAW'}); 
  return CardService.newActionResponseBuilder()
                    .setNotification(CardService.newNotification()
                    .setType(CardService.NotificationType.INFO)
                    .setText("担当者を修正成功"))
                    .setNavigation(nav)
                    .build();
}