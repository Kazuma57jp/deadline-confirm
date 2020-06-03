
//共通変数の宣言
var mySS=SpreadsheetApp.getActiveSpreadsheet();

var sheet0 = mySS.getSheetByName("DATA");                           //データのシート名
var column0 = sheet0.getDataRange().getLastColumn();                //最右列
var row0 = sheet0.getDataRange().getLastRow();                      //最下行
var var0 = sheet0.getRange(1,1,row0,column0).getValues();           //対象のデータの範囲
var var1 = sheet0.getRange(1,3,1,column0).getValues();     　 　　　 //FLG行
var flg = var1[0].indexOf(1);　　　  　　　　　　　　　　              　//1のFLGがあるかどうか検索


var label0 = sheet0.getRange("5:5").getValues();    　　　　　　　　　 // 項目名
  
var sheet1 = mySS.getSheetByName("TMP");　　　　　　　　　　　　　　　　 //メール送信テンプレート部
var row1 = sheet1.getDataRange().getLastRow();                      //TMPシート最下部

var sheet2 = mySS.getSheetByName("config");　　　　　　　　　　　　　　 //リダイレクト設定用シート
var sheet3 = mySS.getSheetByName("log");　　　　　　　　　　　　　　 　　//メールレスポンスログ蓄積用シート
var row3 = sheet3.getDataRange().getLastRow();                      //logシート最下部
var column3 = sheet3.getDataRange().getLastColumn();                //最右列
var var3 = sheet3.getRange(1,1,row3,column3).getValues();           //対象のデータの範囲

var sheet4 = mySS.getSheetByName("list");　　　　　　　　　　　　　　 　　//氏名検索用シート
var column4 = sheet4.getDataRange().getLastColumn();                //最右列
var row4 = sheet4.getDataRange().getLastRow();                      //最下行
var var4 = sheet4.getRange(1,1,row4,column4).getValues();           //対象のデータの範囲

var restime = Utilities.formatDate(new Date, 'JST', 'yyyy/MM/dd HH:mm'); //タイムスタンプ
var resName = Session.getActiveUser().getEmail(); //アクセスした人のメアド

var title = sheet1.getRange("C5").getValue();         　　　　　          　　　　 // リマインド
var add = sheet1.getRange("C11").getValue();         　　　　　          　　　　 // リマインド

var from = sheet1.getRange("C6").getValue();         　　　　　          　　　　　// 送信元
var name = sheet1.getRange("C7").getValue();         　　　　　          　　　　　// 名前
var cc   = sheet1.getRange("C9").getValue();         　　　　　　　　　　          // 宛先cc
var bcc  = sheet1.getRange("C10").getValue();              　　　　　　　　　　    // 宛先Bcc
    
var style1 = " style='background:#eeeeee ;border:1px #000000 solid;border-collapse:collapse ;table-layout: fixed;text-align: left;'";
 

var res1 = sheet2.getRange("D2").getValue();        　　　　　　　　　          　// 完了
var res2 = sheet2.getRange("D3").getValue();        　　　　　　　　　          　// 対応中
var res3 = sheet2.getRange("D4").getValue();        　　　　　　　　　          　// 対象外

var style1 = " style='background:#eeeeee ;border:1px #000000 solid;border-collapse:collapse ;table-layout: fixed;text-align: left;'";
var style2 = " style='height: 25px; width: 40px; font-weight: lighter; text-align: center; padding: 5px 20px; margin: 10px; border:1px solid #38aef0; color: #ffffff; background:#38aef0; font :bold;'";
var style3 = " style='height: 25px; width: 40px; font-weight: lighter; text-align: center; padding: 5px 20px; margin: 10px; border:1px solid #ffbb68; color: #ffffff; background:#ffbb68; font :bold;'";
var style4 = " style='height: 25px; width: 40px; font-weight: lighter; text-align: center; padding: 5px 20px; margin: 10px; border:1px solid #555555; color: #ffffff; background:#555555; font :bold;'";

//----------------------------------------------------------------------------------------------------------------------------------
//ツールバーへのメニュー追加

function onOpen() {

var entries = [  
    {name: "フォーム送信", functionName: "trigger1"},
];
mySS.addMenu("スクリプト実行", entries);
};

//----------------------------------------------------------------------------------------------------------------------------------  
//URLの生成

function getScriptUrl() {
//url生成

    //var url = ScriptApp.getService().getUrl();
    var url = sheet1.getRange("C12").getValue();       　　　　　　　　　          　  // スプレッドリンク
    return url;
}

//----------------------------------------------------------------------------------------------------------------------------------  
//ユーザー名の取得

function getusername(){

 var uname=Session.getActiveUser();
  
  return uname;

}

//----------------------------------------------------------------------------------------------------------------------------------  
//logシートに結果が帰ってきたら、Dataシートのステータスを更新するスクリプトを実行する。

function getNow() {

	var now = new Date();
	var year = now.getFullYear();
	var mon = now.getMonth()+1; //１を足すこと
	var day = now.getDate();
	var hour = now.getHours();
	var min = now.getMinutes();
	var sec = now.getSeconds();

	//出力用
	var udate = year + "/" + mon + "/" + day + "    " + hour + ":" + min + ":" + sec; 
    
	return udate;
}

//----------------------------------------------------------------------------------------------------------------------------------  
//logシートに結果が帰ってきたら、Dataシートのステータスを更新するスクリプトを実行する。

function doGet(e) {

  //GASのWebアプリケーションのURLに?name=○○とつけておくとその値をGAS側で受け取ることができる

  var name = e.parameter.name;　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　 //レスポンスされたステータス名を取得する。　　
  var no = Number(e.parameter.no) ;   　　　　　　　　　　　　　　　　　　　　　　       //レスポンスされたNoを取得する。
  var item = var0[4][no + 1];　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　//アイテム名を取得する。
  var title = sheet1.getRange("C5").getValue();         　　　　　          　　　　 //リマインド
  var URL = sheet1.getRange("C12").getValue();       　　　　　　　　　          　  //スプレッドリンク
    
  //configシートの中からリダイレクト先のURLを取得する
  var arrConfig = sheet2.getDataRange().getValues();
  
  var url = '';
  for(var i = 1; i < arrConfig.length; i++) {
    if(arrConfig[i][0] === name) {
      url = arrConfig[i][1];
      break;
  }
  }

   //urlが見つからなかった場合は強制的にエラーにする
   if(url === '') {
    throw 'urlが見つかりませんでした。configシートのリンクを正しく設定して下さい。';
  }
  
  //氏名を取得する。
  var vname = '';
  for(var j = 1; j < var4.length; j++) {
    if(var4[j][1] === resName) {
      vname = var4[j][0];
      break;
    }
  }
  
  if(name === "Complete"){name = "完了"};
  if(name === "Doing"){name = "対応中"};
  if(name === "Excluded"){name = "対象外"};
  
  //メールからの戻り値を配列に格納して、logシートに記入する。。
  var array = [restime, resName, name , e.parameter.no,vname,1 ];
  sheet3.appendRow(array);
    
  var tpl = HtmlService.createTemplateFromFile('result');
  tpl.title = title;
  tpl.name = name;
  tpl.vname = vname;
  tpl.no = no;
  tpl.item = item;
  tpl.URL = URL;
  return tpl.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME); 
  
}

//----------------------------------------------------------------------------------------------------------------------------------  
//logシートに結果が帰ってきたら、Dataシートのステータスを更新するスクリプトを実行する。

function Drawresult(){

  for(var i=1;i<row3;i++){　　　       　   //シートlogの行数分繰り返す　　　　　
    if (var3[i][5]!=1) continue;          //シートlogのflg行1じゃなかったらスキップする。
      resaddress = var3[i][1]               //シートlogの該当のi行の2列目（メールアドレス)を変数resaddressに格納する。
      ressts = var3[i][2]
      resno = var3[i][3]
      resname = var3[i][4]
      
        for(var k=6;k<row0;k++){           //シートDataの2列目の氏名に該当する行のステータスを更新する。
        //Browser.msgBox(var0[k][1]);
        if(var0[k][1] === resname){
        //Browser.msgBox("ミツケタゾ");
        sheet0.getRange(k+1,resno+2).setValue( ressts )
        sheet3.getRange(i+1,6).setValue( "OK" )　 //シートlogの該当業のflg列に名前検索が完了したこと（完了）を記入する。
        }
        }
   }
  
};


//----------------------------------------------------------------------------------------------------------------------------------
//フォーム送信スクリプトチェック

function trigger1(){
//FLG1があればFormsendを実行する。

  if (flg>=0){
  Formsend()}
  else
  {
Browser.msgBox("送信対象がありませんでした");
};    
}


//----------------------------------------------------------------------------------------------------------------------------------  
//メール送信スクリプト

function Formsend(){

    var repMoji;                                                                    // メール本文情報の取得　※concatで参照渡しを切断

    var to   = sheet1.getRange("C8").getValue();                                    // 宛先To
    
    for(var i=0;i<column0;i++){                                                    //左から3列目から開始

    if(var1[0][i] == 1){                                                         　//送信対象flgチェック
                                                        
      sheet0.getRange(3,i + 3).setValue( i + 1 )　　　　　　　　　　　　　　　　　　　　　　 //Noを入れる
      sheet0.getRange(7,i + 3).setNumberFormat('@');                               //書式を文字列にする。
      sheet0.getRange(8,i + 3).setNumberFormat('@');                               //書式を文字列にする。

    //Browser.msgBox(var0[4][i + 2])
    
    var No = var0[2][i + 2]       　　　　　　　　　　　　                             //No
    var No = i + 1               　　　　　　　　　　　　                             //No
    var FLG = var0[3][i + 2]  　　　　　　　　　　　　       　　　　　　　　　          　//対象者
    var item = var0[4][i + 2]       　　　　　　　　　　　　                           //コンテンツ名
    var link = var0[5][i + 2]      　　　　　　　　　　　　                            //コンテンツへのリンク
    var limit = var0[6][i + 2]       　　　　　　　　　　　　                          //回答期限
    var remind = var0[7][i + 2]       　　　　　　　　　　　　                         //リマインド日
    var ss   = sheet1.getRange("C12").getValue();                                    //スプレッドシートへのリンク
    
    subj = FLG + "【 No. "+ No + " " + item + " 】" + add
    
    var tableList = [];
    tableList.push("<p></p>");
  　tableList.push("<p style='color:#b2b2b2;font-size:12px'>*このメールはGASによる自動送信です。<br/>*このメールへの返信はご遠慮下さい。</p>");
  　tableList.push("<p><br/></p>");
  　tableList.push("<p>各位</p>");
  　tableList.push("<p></p>");
  　tableList.push("<p>お疲れ様です。</p>");
    tableList.push("<p></p>");
    tableList.push("<p>"+title+"</p>");
    tableList.push("<p>に新しいアイテムが追加されました。<br/>内容確認の上、対応結果の報告をお願いいたします。</p>");
    tableList.push("<p></p>");
    tableList.push("<p style='font :bold; font-size :14px; color:#cc0000; font-style: italic;'>" + FLG + "</p>");
    tableList.push("<p></p>");
    tableList.push("<p>　| アイテム名</p>");
    tableList.push("<p style='font :bold; font-size :16px;'>　　" + item + "</p>");
    tableList.push("<p></p>");
    tableList.push("<p>　| 対応期限</p>");
    tableList.push("<p style='font :bold; font-size :16px;'>　　" + limit + "</p>");
    tableList.push("<p></p>");
    tableList.push("<p>　| コンテンツリンク</p>");
    tableList.push("<p>　　" + link + "</p>");
    tableList.push("<p>　| 対応状況</p>");
    tableList.push("<p><br/></p>");
    tableList.push("<p><a href = " + res1 + No + " style = 'text-decoration: none;'><span" + style2 + " style = 'text-decoration: none;'>完了</span></a><a href=" + res2  + No +" style = 'text-decoration: none;'><span " + style3 + ">対応中</span></a><a href=" + res3  + No + " style = 'text-decoration: none;'><span " + style4 + ">対象外</span></a></p>");　
    tableList.push("<p><br/></p>");
    tableList.push("<p>　| スプレッドシートへのリンク</p>");
    tableList.push("<p></p>");    
    tableList.push("<a href='" + ss + "'>"+ title + "</a><br>")　
    tableList.push("<p><br/></p>");
    tableList.push("<p>以上です。</p>");　
    tableList.push("<p>よろしくお願いします。<br/></p>");　

    var htmlbody = tableList.push("<table><table" + style1 + ">");
    
    GmailApp.sendEmail(
    to, //宛先
    subj, //件名
    'htmlメールが表示できませんでした', //本文
    {
      from: from, 
      name : name ,
      cc : cc,
      bcc : bcc,
      htmlBody: tableList.join("\n")
    }
  　);
    
    // 送信済みフラグなどの処理  
    sheet0.getRange(1,i + 3).setValue("SEND")
    sheet0.getRange(2,i + 3).setValue(Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:MM:SS"))
       
    }    
   }
   
};


//----------------------------------------------------------------------------------------------------------------------------------  
//リマインドメール送信スクリプト

function trigger2(){

  Ddate = Utilities.formatDate(new Date, 'JST', 'yyyy/MM/dd'); 　　　　//今日の日付ね
  
  var varl = sheet0.getRange(7,3,1,column0).getValues();     　 　　　 //締め切り期限日のデータ行
  var lflg = varl[0].indexOf(Ddate); 　　　　　　　　　     　 　　　 　　　///締め切り期限日のデータ行に該当があるか検索する。
  
  var vark = sheet0.getRange(8,3,1,column0).getValues();     　 　　　 //リマインド日のデータ行
  var kflg = vark[0].indexOf(Ddate); 　　　　　　　　　     　 　　　 　　　//リマインド日のデータ行に該当があるか検索する。  
  
  var varr = sheet0.getRange(9,3,1,column0).getValues();     　 　　　 //リマインドFLG行
  var rflg = varr[0].indexOf('要'); 　　　　　　　　　　              　　//要のFLGがあるかどうか検索


//リマインド要の案件がそもそもあるか
  if (rflg>=0 || kflg==Ddate){
  //その中でもリマインド日が今日の日付のものがあるか
    
  　　//Browser.msgBox("いきます")
    remaidmailsend()
  　
  }
  
}

function remaidmailsend(){

  Ddate = Utilities.formatDate(new Date, 'JST', 'yyyy/MM/dd'); 　　　　//今日の日付ね

  for(var i=0;i<column0;i++){                                                       //左から3列目から開始
  
    if(var0[8][i] == '要'&& var0[7][i] == Ddate && var0[3][i] == '*全員必須') {                                               //1.dataの送信対象flgの列を特定する。
    //Browser.msgBox(i)
      //if(var0[7][i] == Ddate) {                                               //2.リマインド日の列を特定する。
      
    for(var k=9;k<row0;k++){              
      if(var0[k][i] == '' || var0[k][i] == '対応中') {                              //2.dataの送信対象flgの列の空白行を特定する。
       remindname = var0[k][1]　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　 //3.dataの送信対象氏名を取得する。
         sts = var0[k][i] 
       
        //氏名を取得する。
        var vname = '';
        for(var j = 1; j < row4; j++) {                                           //4.氏名からlistのメールアドレスを取得する。
        if(var4[j][0] === remindname) {
        remindaddress = var4[j][1];

    to = remindaddress //宛先
    
    var FLG = var0[3][i] 
    var No = var0[2][i]       　　　　　　　　　　　　                             //No
    var item = var0[4][i]       　　　　　　　　　　　　                           //コンテンツ名
    var FLG = var0[3][i]  　　　　　　　　　　　　       　　　　　　　　　          　//対象者
    var link = var0[5][i]      　　　　　　　　　　　　                            //コンテンツへのリンク
    var limit = var0[6][i]       　　　　　　　　　　　　                          //回答期限
    
    //Browser.msgBox(var0[4][i])
        
    subj = "まもなく対応期限です。 " + FLG + "【 No."+ No + " " + item + " 】" + add //件名
  
    var tableList = [];
    tableList.push("<p></p>");
  　tableList.push("<p style='color:#b2b2b2;font-size:12px'>*このメールはGASによる自動送信です。<br/>*このメールへの返信はご遠慮下さい。</p>");
  　tableList.push("<p><br/></p>");
  　tableList.push("<p>各位</p>");
  　tableList.push("<p></p>");
  　tableList.push("<p>お疲れ様です。</p>");
    tableList.push("<p></p>");
    tableList.push("<p>"+title+"</p>");
    tableList.push("<p>まもなく下記アイテムの対応期限です。<br/>内容確認の上対応お願いします。<br/>現在のステータスは以下の通りです。</p>");
    tableList.push("<p></p>");
    tableList.push("<p style='font :bold; font-size :14px; color:#cc0000; font-style: italic;'>" + FLG + "</p>");
    tableList.push("<p></p>");
    tableList.push("<p>　| アイテム名</p>");
    tableList.push("<p style='font :bold; font-size :16px;'>　　" + item + "</p>");
    tableList.push("<p></p>");
    tableList.push("<p>　| 対応期限</p>");
    tableList.push("<p style='font :bold; font-size :16px;'>　　" + limit + "</p>");
    tableList.push("<p></p>");
    tableList.push("<p>　| リンク</p>");
    tableList.push("<p>　　" + link + "</p>");
    tableList.push("<p>　| 対応状況</p>");
    tableList.push("<p>　　" + sts + "</p>");
    tableList.push("<p><br/></p>");
    tableList.push("<p><a href = " + res1 + No + " style = 'text-decoration: none;'><span" + style2 + " style = 'text-decoration: none;'>完了</span></a><a href=" + res2  + No +" style = 'text-decoration: none;'><span " + style3 + ">対応中</span></a><a href=" + res3  + No + " style = 'text-decoration: none;'><span " + style4 + ">対象外</span></a></p>");　
    tableList.push("<p><br/></p>");
    tableList.push("<p>以上です。</p>");　
    tableList.push("<p>よろしくお願いします。<br/></p>");　

    var htmlbody = tableList.push("<table><table" + style1 + ">");
   
   try{
   // 処理1
   
   
   GmailApp.sendEmail(
    to,
    subj,
    'htmlメールが表示できませんでした', //本文
    {
      from: from, 
      name : name ,
      cc : cc,
      bcc : bcc,
      htmlBody: tableList.join("\n")
    }
  　);
   }catch(e){
   };
   };
   };
   };
        
        //}
        }
       
      }
     }
   };
