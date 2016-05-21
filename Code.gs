   /*
 * ChangeLog
 * 6 Mars 2016 - Adaptation au nouveau site du Bon Coin, ainsi que quelques nettoyages
 * 7 Mars 2016 - Format d'email plus compact
 * 21 Mars 2016 - Correction message d'erreur si email destinataire non défini
 * 30 Mars 2016 - Identifie si la photo est manquante dans l'annonce, itération plus propre dans les annonces
 * 31 Mars 2016 - Correction regression dans le case de "Setup Recherche"
 * 20 Mai 2016 - Modifs proposées par Franck : ajout de l'heure dans le log (à partir de ligne 112) + ajout de l'image "https://www.leboncoin.fr/img/no-picture-adview.png" lorsque l'annonce n'a pas de photo (ligne 257) + ajout de la fonction purgeLog, qui permet de supprimer des lignes dans le log au dela du seuil défini par l'utilisateur
 */

var debug = false;

var menuLabel = "Lbc Alertes";
var menuMailSetupLabel = "Setup email";
var menuSearchSetupLabel = "Setup recherche";
var menuSearchLabel = "Lancer manuellement";
var menuLog = "Activer/Désactiver les logs";
var menuArchiveLog = "Archiver les logs";
var menuPurgeLog = "Purger le log";
var menuNumberOfRowsToKeepInLog = "Nombre de lignes à conserver dans le log lors d'une purge";

var scriptProperties = PropertiesService.getScriptProperties();

function lbc(sendMail){
  if(sendMail != false){
    sendMail = true;
  }

  var to = scriptProperties.getProperty('email');
  if(to == "" || to == null ){
    Browser.msgBox("L'email du destinataire n'est pas défini. Allez dans le menu \"" + menuLabel + "\" puis \"" + menuMailSetupLabel + "\".");
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Recherches");
    var slog = ss.getSheetByName("Log");
    var i = 0; var nbSearchWithRes = 0; var nbResTot = 0;
    var body = ""; var corps = ""; var announceHTML = ""; var bodyHTML = ""; var summary = ""; var searchURL = ""; var searchName = "";

    while((searchURL = sheet.getRange(2+i,2).getValue()) != "") {
      
      searchName = sheet.getRange(2+i,1).getValue();
      Logger.log("Recherche pour " + searchName);
      var nbRes = 0;
      body = "";
      announceHTML = "";

      var rep = UrlFetchApp.fetch(searchURL).getContentText("iso-8859-15");
      
      if(rep.indexOf("Aucune annonce") < 0) {
        
        var data = splitResult_(rep);
        data = data.substring(data.indexOf("<a"));
        
        var announceURL = extractA_(data);
        var firstID     = extractId_(announceURL);
        
        if(sendMail == null || sendMail == true) {
          
          var lastSavedID = sheet.getRange(2+i,3).getValue();
          
          if (firstID != lastSavedID) {
            
            var announceId = firstID;
          
            //While ID of announce is different from the saved one
            do {
              
              Logger.log("data = " + data);

              var endAnnounceMarker = "</section>";
              var endAnnounceMarkerPos = data.indexOf(endAnnounceMarker);
                
              if (endAnnounceMarkerPos > 0) {

                nbRes++;
                
                var title = extractTitle_(data);
                var place = extractPlace_(data);
                var price = extractPrice_(data);
                var vendpro = extractPro_(data);
                var date = extractDate_(data);
                var image = extractImage_(data, endAnnounceMarkerPos);

                announceHTML += "<tr style=\"height:1px; padding-bottom:10px;\"><td style=\"border-top:1px solid #ccc;\" colspan=\"2\"></td></tr>"
                announceHTML += "<tr><td style=\"width:200px;padding-right:20px;\"><a href=\"" + announceURL + "\" target=\"" + announceId + "\"><img src=\""+ image +"\"></a></td>";
                announceHTML += "<td style=\"align:left; padding-left:10px;\"><a href=\"" + announceURL + "\" target=\"" + announceId + "\" style=\"font-size: 14px;font-weight:bold;color:#369;text-decoration:none;\">";
                announceHTML += title + vendpro +"</a> <div>" + place + "</div><div>" + date + "</div><div style=\"font-size:14px;font-weight:bold;\">" + price + "</div>";
                announceHTML += "</td></tr>";
                               
                //Skip the block already analyzed
                data = data.substring(endAnnounceMarkerPos+endAnnounceMarker.length);
                
                announceURL = extractA_(data);
                announceId  = extractId_(announceURL);
              }
              else
                announceId = "";
              
            } while ((announceId != "") && (announceId != lastSavedID))

            if (nbRes > 0)
               nbSearchWithRes++;
            
            bodyHTML += "<p style=\"display:block;clear:both;padding-top:2px;font-size:14px;font-weight:bold;background:#F1F1F5;\">Recherche <a name=\""+ searchName;
            bodyHTML += "\" href=\""+ searchURL + "\"> "+ searchName +" (" + nbRes + ")</a></p>";
            bodyHTML += "<table border=\"0\" style=\"width:100%; vertical-align:middle; background:#FFFFFF;\"><tbody>" + announceHTML + "</tbody></table>";

            summary  += "<li><a href=\"#"+ searchName + "\">"+ searchName +" (" + nbRes + ")</a></li>"
                        
            if(scriptProperties.getProperty('log') == "true" || scriptProperties.getProperty('log') == null || scriptProperties.getProperty('log') == ""){
              slog.insertRowBefore(2);
              slog.getRange("A2").setValue(searchName);
              slog.getRange("B2").setValue(nbRes);
              var currentDate = new Date();
              slog.getRange("C2").setValue(currentDate.getDate() + "/" + currentDate.getMonth() + "/" + currentDate.getYear() + " - " + currentDate.toLocaleTimeString().replace(" CEST",""));
              //slog.getRange("C2").setValue(new Date);
            }
          }
        }
        sheet.getRange(2+i,3).setValue(firstID);
        nbResTot += nbRes;
      } else {
        //pas de résultat
        sheet.getRange(2+i,3).setValue(123);
      }
      i++;
    }
    
    if(nbSearchWithRes > 1) {
      //plusieurs recherches, on créé un summary
      summary = "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Accès rapide :</p><ul>" + summary + "</ul>";
      bodyHTML = summary + bodyHTML;
      debug_(summary);
    }
    
    debug_("Nb de res tot:" + nbResTot);
    //on envoie le mail?
    if(nbSearchWithRes > 0){
      var title = "Alerte leboncoin.fr : " + nbResTot + " nouveau" + (nbResTot>1?"x":"") + " résultat" + (nbResTot>1?"s":"");
      debug_("titre msg : " + title);
      corps = "Si cet email ne s’affiche pas correctement, veuillez sélectionner\nl’affichage HTML dans les paramètres de votre logiciel de messagerie.";
      //debug_("corps msg : " + corps);
      bodyHTML = "<body>" + bodyHTML + "</body>";
      debug_("bodyHTML msg : " + bodyHTML);
      
      if (bodyHTML.length > 200*1024) {
        // Email body size is limited to 200 KB. Too big body is then truncated to avoid script to fail
        bodyHTML = bodyHTML.substring(0, 200*1024-1);
      }
        
      MailApp.sendEmail(to,title,corps,{ htmlBody: bodyHTML });
      debug_("Nb mail journalier restant : " + MailApp.getRemainingDailyQuota());
    }
  }
}

/**
* Extrait l'id de l'annonce LBC
*/
function extractId_(data){
  
  var lastSlashPos = data.lastIndexOf("/");
  
  if (lastSlashPos < 0)
    return ""
  else
    return data.substring(lastSlashPos + 1,data.indexOf(".htm", lastSlashPos));
}

/**
* Extrait le lien de l'annonce
*/
function extractA_(data){
  
  var aPos = data.indexOf("<a");
  if (aPos < 0)
    return "";
  
  var found = data.substring(aPos + 9 , data.indexOf(".htm", aPos + 9) + 4);
  
  // Handle case when the URL doesn't start by http:
  if (found.indexOf("//") == 0)
    return "http:" + found;
  else
    return found;
}

/**
* Extrait le titre de l'annonce
*/
function extractTitle_(data){
  return data.substring(data.indexOf("title=") + 7 , data.indexOf("\"", data.indexOf("title=") + 7) );
}

/**
* Extrait vendeur pro
*/
function extractPro_(data){
  
  var proMarker = "<span class=\"ispro\">";
  var proStart  = data.indexOf(proMarker);
  var pro = data.substring(proStart + proMarker.length, data.indexOf("</span>", proStart + proMarker.length) );
  
  if(pro.indexOf("(pro)") > 0){
    return " (pro)";
  }else{
    return "";
  }
}

/**
* Extrait le lieu de l'annonce
*/
function extractPlace_(data){
  
  // Look for the 2nd "item_supp" block  
  var infoMarker = "<p class=\"item_supp\">";
  var info1pos = data.indexOf(infoMarker);
  var info2pos = data.indexOf(infoMarker, info1pos + infoMarker.length);
    
  return data.substring(info2pos + infoMarker.length, data.indexOf("</p>", info2pos) );
}

/**
* Extrait le prix de l'annonce
*/
function extractPrice_(data){

  var priceMarker = "<h3 class=\"item_price\">";
  var priceStart  = data.indexOf(priceMarker);
  if (priceStart < 0)
    return "";
  else
    return data.substring(priceStart + priceMarker.length, data.indexOf("</h3>", priceStart + priceMarker.length) );
}

/**
* Extrait la date de l'annonce
*/
function extractDate_(data){

  // Look for the 3rd "item_supp" block  
  var infoMarker = "<p class=\"item_supp\">";
  var info1pos = data.indexOf(infoMarker);
  var info2pos = data.indexOf(infoMarker, info1pos + infoMarker.length);
  var info3pos = data.indexOf(infoMarker, info2pos + infoMarker.length);
    
  return data.substring(info3pos + infoMarker.length, data.indexOf("</p>", info3pos) );
}

/**
* Extrait l'image de l'annonce
*/
function extractImage_(data, endAnnounceMarkerPos){
  
  var imgStartMarker = "data-imgSrc=";
  var imageStart = data.indexOf(imgStartMarker);
  if ((imageStart < 0) || (imageStart > endAnnounceMarkerPos)) {
    return "https://www.leboncoin.fr/img/no-picture-adview.png";
  }
  else {
    
    var imageEnd = data.indexOf("data-imgAlt=", imageStart);
    var image = data.substring(imageStart + imgStartMarker.length + 1, imageEnd - 2);
    image = image.replace("//","http://");
    return image;
  }
}

/**
* Extrait la liste des annonces
*/
function splitResult_(text){
  
  var debut = text.indexOf("<section id=\"listingAds\"");
  debut = text.indexOf("<li>", debut);
  var fin = text.indexOf("<div id=\"google_ads\"");
  return text.substring(debut,fin);
}

//Activer/Désactiver les logs
function dolog(){
  if(scriptProperties.getProperty('log') == "true"){
    scriptProperties.setProperty('log', false);
    Browser.msgBox("Les logs ont été désactivées.");
  }
  else if(scriptProperties.getProperty('log') == "false"){
    scriptProperties.setProperty('log', true);
    Browser.msgBox("Les logs ont été activées.");
  }
  else{
    scriptProperties.setProperty('log', false);
    Browser.msgBox("Les logs ont été désactivées.");
  }
}

function setup(){
  lbc(false);
}

function setupMail(){
  if(scriptProperties.getProperty('email') == "" || scriptProperties.getProperty('email') == null ){
    var quest = Browser.inputBox("Entrez votre email, le programme ne vérifie pas le contenu de cette boite.", Browser.Buttons.OK_CANCEL);
    if(quest == "cancel"){
      Browser.msgBox("Ajout email annulé.");
      return false;
    }else{
      scriptProperties.setProperty('email', quest);
      Browser.msgBox("Email " + scriptProperties.getProperty('email') + " ajouté");
    }
  }else{
    var quest = Browser.inputBox("Entrez un email pour modifier l'email : " + scriptProperties.getProperty('email') , Browser.Buttons.OK_CANCEL);
    if(quest == "cancel"){
      Browser.msgBox("Modification email annulé.");
      return false;
    }else{
      scriptProperties.setProperty('email', quest);
      Browser.msgBox("Email " + scriptProperties.getProperty('email') + " ajouté");
    }
  }
}

//Archiver les logs
function archivelog(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var slog = ss.getSheetByName("Log");
  var today  = new Date();
  var newname = "LogArchive " + today.getFullYear()+(today.getMonth()+1)+today.getDate();
  slog.setName(newname);
  var newsheet = ss.insertSheet("Log",1);
  newsheet.getRange("A1").setValue("Recherche");
  newsheet.getRange("B1").setValue("Nb Résultats");
  newsheet.getRange("C1").setValue("Date");
  newsheet.getRange(1,1,2,3).setBorder(true,true,true,true,true,true);
}

function setupNumberOfRowsToKeepInLog()
{
  if(ScriptProperties.getProperty('NumberOfRowsToKeepInLog') == "" || ScriptProperties.getProperty('NumberOfRowsToKeepInLog') == null ){
    var quest = Browser.inputBox("Indiquez le nombre de lignes à conserver dans le log lors d'une purge : ", Browser.Buttons.OK_CANCEL);
    if(quest == "cancel"){
      Browser.msgBox("Paramétrage du nombre de lignes à conserver dans le log annulé, valeur inchangée (= " + ScriptProperties.getProperty('NumberOfRowsToKeepInLog') + ")");
      return false;
    }else if(isNaN(quest)){
      Browser.msgBox("Vous devez entrer une valeur numérique entière");
      setupNumberOfRowsToKeepInLog()
    }else{
      if(quest == 0){quest = 1;} else if(quest != "") {quest++;} else {quest = null;} //On préserve la première ligne contenant les entêtes de colone
      ScriptProperties.setProperty('NumberOfRowsToKeepInLog', quest);
      Browser.msgBox("Nombre de lignes à conserver dans le fichier de log lors d'une purge paramétré à : " + ScriptProperties.getProperty('NumberOfRowsToKeepInLog') + " (Notez que la première ligne contenant les entêtes de colone sera conservée)");
      return true;
    }
  }else{
    var quest = Browser.inputBox("Indiquez le nombre de lignes à conserver dans le log lors d'une purge : (valeur actuelle = ", + ScriptProperties.getProperty('NumberOfRowsToKeepInLog') + ")" , Browser.Buttons.OK_CANCEL);
    if(quest == "cancel"){
      Browser.msgBox("Paramétrage du nombre de lignes à conserver dans le log annulé, valeur inchangée (= " + ScriptProperties.getProperty('NumberOfRowsToKeepInLog') + ")");
      return false;
    }else if(isNaN(quest)){
      Browser.msgBox("Vous devez entrer une valeur numérique entière");
       setupNumberOfRowsToKeepInLog()
    }else{
      if(quest == 0){quest = 1;} else if(quest != "") {quest++;} else {quest = null;} //On préserve la première ligne contenant les entêtes de colone
      ScriptProperties.setProperty('NumberOfRowsToKeepInLog', quest);
      Browser.msgBox("Nombre de lignes à conserver dans le fichier de log lors d'une purge paramétré à : " + ScriptProperties.getProperty('NumberOfRowsToKeepInLog') + " (Notez que la première ligne contenant les entêtes de colone sera conservée)");
      return true;
    }
  }
}

function purgeLog()
{
  if(ScriptProperties.getProperty('NumberOfRowsToKeepInLog') == "" || ScriptProperties.getProperty('NumberOfRowsToKeepInLog') == null ){
    if (setupNumberOfRowsToKeepInLog() == false){Browser.msgBox("Purge annulée, le nombre de lignes à conserver dans le log n'est pas paramétré");}
  }
  if(ScriptProperties.getProperty('NumberOfRowsToKeepInLog') != "" && ScriptProperties.getProperty('NumberOfRowsToKeepInLog') != null ){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log");
    var howmany = sheet.getLastRow() - ScriptProperties.getProperty('NumberOfRowsToKeepInLog')
    if(howmany > 0) {
      sheet.deleteRows(ScriptProperties.getProperty('NumberOfRowsToKeepInLog'), howmany);
    }
  }
}

function onOpen() {
var sheet = SpreadsheetApp.getActiveSpreadsheet();
var entries = [{
name : menuMailSetupLabel,
functionName : "setupMail"
},{
name : menuSearchSetupLabel,
functionName : "setup"
},
  null
,{
name : menuSearchLabel,
functionName : "lbc"
},
null
,{
name : menuLog,
functionName : "dolog"
},{
name : menuArchiveLog,
functionName : "archivelog"
},{
name : menuPurgeLog,
functionName : "purgeLog"
},{
name : menuNumberOfRowsToKeepInLog,
functionName : "setupNumberOfRowsToKeepInLog"
}
];
sheet.addMenu(menuLabel, entries);
}

function onInstall()
{
onOpen();
}

/**
* Retourne la date
*/
function myDate_(){
var today = new Date();
debug_(today.getDate()+"/"+(today.getMonth()+1)+"/"+today.getFullYear());
return today.getDate()+"/"+(today.getMonth()+1)+"/"+today.getFullYear();
}

/**
* Retourne l'heure
*/
function myTime_(){
var temps = new Date();
var h = temps.getHours();
var m = temps.getMinutes();
if (h<"10"){h = "0" + h ;}
if (m<"10"){m = "0" + m ;}
debug_(h+":"+m);
return h+":"+m;
}

/**
* Debug
*/
function debug_(msg) {
if(debug != null && debug) {
Browser.msgBox(msg);
}
}
