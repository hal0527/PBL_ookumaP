//１プロジェクト追加・２タスクの分配
function sheetAddOn(e){
  var accessToken = e.messageMetadata.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  var section1 = CardService.newCardSection();
  var section2 = CardService.newCardSection();
  var sheet_id = PropertiesService.getUserProperties().getProperty('sheet_id');
  var projects_list = Sheets.Spreadsheets.Values.get(spread_sheet_id, 'sheet2!b2:b').values;
  var sales_list = Sheets.Spreadsheets.Values.get(spread_sheet_id, 'sheet3!b2:b').values;

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
    var project_range = 'sheet2!C'+(i+2);
    projects.addItem(projects_list[i][0], project_range, false);
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
  Logger.log(project_range);
    Logger.log(sales_name);
}