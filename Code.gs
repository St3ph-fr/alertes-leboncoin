function lbc(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Données");
  var slog = ss.getSheetByName("Log");
  var i =0; var body = ""; var corps = "";
  var stop = false;
  while(sheet.getRange(2+i,2).getValue() != ""){
    var compteur = 0;
    body = "";
    stop = false;
    var rep = UrlFetchApp.fetch(sheet.getRange(2+i,2).getValue()).getContentText();
    if(rep.indexOf("Aucune annonce") < 0){
    var data = splitresult(rep);
    data = data.substring(data.indexOf("<a"));
    var firsta = data.substring(data.indexOf("<a") + 9 , data.indexOf(".htm", data.indexOf("<a") + 9) + 4);
    var holda = sheet.getRange(2+i,3).getValue();
    if(extractid(firsta) != holda && holda != ""){
    while(data.indexOf("<a") > 0 || stop == false){
      var a = data.substring(data.indexOf("<a") + 9 , data.indexOf(".htm", data.indexOf("<a") + 9) + 4);
      if(extractid(a) != holda){
        
        var title = data.substring(data.indexOf("title=") + 7 , data.indexOf("\"", data.indexOf("title=") + 7) );
var place = data.substring(data.indexOf("placement") + 11 , data.indexOf("</div>", data.indexOf("placement") + 11) );
// test à optimiser car c'est hyper bourrin [mlb]
var isPrice = String(data.substring(data.indexOf("price"), data.indexOf("price")+250)).match(/price/gi);
if (isPrice) {
var price = data.substring(data.indexOf("price") + 7 , data.indexOf("</div>", data.indexOf("price") + 7) );
} else {
var price = "";
}
var date = data.substring(data.indexOf("date") + 6 , data.indexOf("class=\"image\"", data.indexOf("date") + 6) - 5);
// test à optimiser car c'est hyper bourrin [mlb]
var isImage = String(data.substring(data.indexOf("image"), data.indexOf("image")+250)).match(/img/gi);
if (isImage) {
var image = data.substring(data.indexOf("class=\"image-and-nb\">") + 21, data.indexOf("class=\"nb\"", data.indexOf("class=\"image-and-nb\">") + 21) - 12);
} else {
var image = "";
}
body = body + "<li style='list-style:none;margin-bottom:20px; clear:both;background:#EAEBF0;border-top:1px solid #ccc;'><div style='float:left;width:90px;padding: 20px 20px 0 0;text-align: right;'>"+ date +"</div><div style='float:left;width:200px;padding:20px 0;'><a href=\"" + a + "\" '>"+ image +"</a> </div><div style='float:left;width:420px;padding:20px 0;'><a href=\"" + a + "\" style='font-size: 14px;font-weight:bold;color:#369;text-decoration:none;'>" + title + "</a> <div>" + place + "</div> <div style='line-height:32px;font-size:14px;font-weight:bold;'>" + price + "</div></div></li>";
if(data.indexOf("<a",10) > 0){
var data = data.substring(data.indexOf("<a",10));
}else{
stop = true;
}
}else{
stop = true;
}
compteur++;
}
corps = corps + "<p style='display:block;clear:both;padding-top:20px;font-size:14px;'> Votre recherche : <a href=\""+ sheet.getRange(2+i,2).getValue() + "\"> "+ sheet.getRange(2+i,1).getValue()+ "</a><ul>" + body + "</ul></p>";
slog.insertRowBefore(2);
slog.getRange("A2").setValue(sheet.getRange(2+i,1).getValue());
slog.getRange("B2").setValue(compteur-1);
slog.getRange("C2").setValue(new Date);
sheet.getRange(2+i,3).setValue(extractid(firsta));
}
}
i++;
}
if(corps != ""){
MailApp.sendEmail(ScriptProperties.getProperty('email'),"Alerte Lbc le " + myDate() + " à " + myTime(),corps,{ htmlBody: corps });
}
}

function setup(){
if(ScriptProperties.getProperty('email') == "" || ScriptProperties.getProperty('email') == null ){
Browser.msgBox("L'email du destintaire n'est pas définit. Allez dans le menu \"Lbc Alertes\" puis \"Gérer email\".");
}
var ss = SpreadsheetApp.getActiveSpreadsheet();
var i = 0;
var sheet = ss.getSheetByName("Données");
while(sheet.getRange(2+i,2).getValue() != ""){
if(sheet.getRange(2+i,3).getValue() == ""){
var rep = UrlFetchApp.fetch(sheet.getRange(2+i,2).getValue()).getContentText();
if(rep.indexOf("Aucune annonce") < 0){
var data = splitresult(rep);
sheet.getRange(2+i,3).setValue(extractid(data.substring(data.indexOf("<a") + 9 , data.indexOf(".htm", data.indexOf("<a") + 9) + 4)));
}else{
sheet.getRange(2+i,3).setValue(123);
}
}
i++;
}
}

function setupmail(){
if(ScriptProperties.getProperty('email') == "" || ScriptProperties.getProperty('email') == null ){
var quest = Browser.inputBox("Entrez votre email, le programme ne vérifie pas le contenu de cette boite.", Browser.Buttons.OK_CANCEL);
if(quest == "cancel"){
Browser.msgBox("Ajout email annulé.");
return false;
}else{
ScriptProperties.setProperty('email', quest);
Browser.msgBox("Email " + ScriptProperties.getProperty('email') + " ajouté");
}
}else{
var quest = Browser.inputBox("Entrez un email pour modifier l'email : " + ScriptProperties.getProperty('email') , Browser.Buttons.OK_CANCEL);
if(quest == "cancel"){
Browser.msgBox("Modification email annulé.");
return false;
}else{
ScriptProperties.setProperty('email', quest);
Browser.msgBox("Email " + ScriptProperties.getProperty('email') + " ajouté");
}
}
}

function extractid(id){
return id.substring(id.indexOf("/",25) + 1,id.indexOf(".htm"));
}
function splitresult(text){
var debut = text.indexOf("<div class=\"list-ads\">");
var fin = text.indexOf("<div class=\"list-gallery\">");
return text.substring(debut + "<div class=\"list-ads\">".length,fin);
}


function onOpen() {
var sheet = SpreadsheetApp.getActiveSpreadsheet();
var entries = [{
name : "Setup email",
functionName : "setupmail"
},{
name : "Setup recherche",
functionName : "setup"
},{
name : "Lancer manuellement",
functionName : "lbc"
}];
sheet.addMenu("Lbc Alertes", entries);
}

function myDate(){
var today = new Date();
//Browser.msgBox(today.getDate()+"/"+(today.getMonth()+1)+"/"+today.getFullYear());
  return today.getDate()+"/"+(today.getMonth()+1)+"/"+today.getFullYear();
}

function myTime(){
var temps = new Date();
  var h = temps.getHours();
  var m = temps.getMinutes();
  if (h<"10"){h = "0" + h ;}
  if (m<"10"){m = "0" + m ;}
  return h+":"+m;
  //Browser.msgBox(h+":"+m);
}
