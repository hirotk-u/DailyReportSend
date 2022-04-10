//===========================
//doGet
//===========================
function doGet() {
  const htmlOutput = HtmlService.createTemplateFromFile("DailyReportSendIndex").evaluate();

  console.log("doGet 完了");

  return htmlOutput;
}

//===========================
//doPost
//===========================
function doPost(e) {
  //現在日付取得(メールタイトル用)
  const nowDate = new Date();
  const nowYear = nowDate.getFullYear();
  const nowMonth = nowDate.getMonth()+1;
  const nowDay = nowDate.getDate();

  //日報送信済チェック
  const hasSent = hasSentDailyReportMail(nowYear, nowMonth, nowDay);
  console.log(Utilities.formatString("hasSent=%s", hasSent));
  
  if(hasSent){
    console.log("日報送信済のためキャンセルしました");
    return ContentService.createTextOutput("日報送信済のためキャンセルしました");
  }

  //日報送信処理
  const ret = sendDailyReportMailDetail(e.parameters,
                                        nowMonth,
                                        nowDay);

  let msg = "";

  //-----------------
  //日報送信後処理
  //-----------------
  if(ret){
    msg = "日報処理が完了しました！";

    //送信履歴書き込み
    const sendDateStr = Utilities.formatString("%s/%s/%s", 
                                                nowYear, nowMonth, nowDay);
    
    const sp = SpreadsheetApp.getActiveSpreadsheet();
    const sendHistorySheet = sp.getSheetByName("sendHistory");
    sendHistorySheet.appendRow([sendDateStr]);

  }else{
    msg = "日報処理が失敗しました！";
  }

  console.log(msg);
  return ContentService.createTextOutput(msg);
}

//===========================
//hasSentDailyReportMail
//===========================
function hasSentDailyReportMail(nowYear, nowMonth, nowDay){
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sp.getSheetByName("sendHistory");
  const lastRowIdx = sheet.getLastRow();
  
  if(lastRowIdx === 0){
    return false;
  }

  const values = sheet.getRange(1, 1, sheet.getLastRow()).getDisplayValues();

  const nowDate = nowYear + "/" + nowMonth + "/" + nowDay;
  for(let val of values){
    if(val[0] === nowDate){
      return true;
    }
  }

  return false;
}

//===========================
//sendDailyReportMailDetail
//===========================
function sendDailyReportMailDetail(parameters, nowMonth, nowDay){
  try {
    const toAddress = getToAddress();
    const mailTitle = getMailTitle(nowMonth, nowDay);
    const mailBody = getMailBody(parameters.startTime,
                                  parameters.endTime,
                                  parameters.projectName,
                                  parameters.content);

    //メール送信
    GmailApp.sendEmail(toAddress, mailTitle, mailBody);

    console.log(Utilities.formatString("To[%s]", toAddress));
    console.log(Utilities.formatString("Title[%s]", mailTitle));
    console.log(Utilities.formatString("Body[\r\n%s\r\n]", mailBody));

  }catch(e){
    console.log(Utilities.formatString("[%s] ErrMsg[%s]",
                  arguments.callee.name, e.message));
    return false;
  }

  return true;
}

//===========================
//getToAddress
//===========================
function getToAddress(){
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sp.getSheetByName("sendTo");
  const values = sheet.getDataRange().getValues();

  return values.join(";");
}

//===========================
//getMailTitle
//===========================
function getMailTitle(nowMonth, nowDay){
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sp.getSheetByName("mailTitle");
  const baseTitle = sheet.getRange(1,1).getValue();

  //mm/ddを現在日付に変換
  const title = baseTitle.replace("mm", nowMonth).replace("dd", nowDay);

  return title;
}

//===========================
//getMailBody
//===========================
function getMailBody(startTime, endTime, projectName, content){
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const mailBodySheet = sp.getSheetByName("mailBody");
  const mailBodyValues = mailBodySheet.getRange(1, 1, mailBodySheet.getLastRow()).getDisplayValues();

  let mailBodyArr = [];

  //MailBody
  for(const row of mailBodyValues){
    let firstCell = row[0];
  
    //開始・終了時刻設定
    firstCell = firstCell.replace("[StartTime]", startTime);
    firstCell = firstCell.replace("[EndTime]", endTime);

    //案件名設定
    firstCell = firstCell.replace("[ProjectName]", projectName);

    //作業内容設定
    firstCell = firstCell.replace("[Content]", content);

    mailBodyArr.push(firstCell);
  }
  
  //署名設定
  const signatureSheet = sp.getSheetByName("signature");
  const signatureValues = signatureSheet.getRange(1, 1, signatureSheet.getLastRow()).getValues();

  mailBodyArr.push("");
  for(const row of signatureValues){
    mailBodyArr.push(row[0]);
  }

  return mailBodyArr.join("\r\n");
}
