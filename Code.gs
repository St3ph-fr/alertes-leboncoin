/**
* global var section
*/
var debug = false;
var menuLabel = "Lbc Alertes";
var menuMailSetupLabel = "Setup email";
var menuSearchSetupLabel = "Setup recherche";
var menuSearchLabel = "Lancer manuellement";
var menuLog = "Activer/Désactiver les logs";
var menuArchiveLog = "Archiver les logs";

//Positionnement des log dans l'onglet log ? true=oui ; false=non
ScriptProperties.setProperty('log', false);

/**
* main function
*/
function lbc(sendMail){
  if(sendMail != false){
    sendMail = true;
  }
  var to = ScriptProperties.getProperty('email');
  if(to == "" || to == null ){
    Browser.msgBox("L'email du destinataire n'est pas défini. Allez dans le menu \"" + menuLabel + "\" puis \"" + menuMailSetupLabel + "\".");
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Données");
    var slog = ss.getSheetByName("Log");
    var i = 0; var nbSearchWithRes = 0; var nbResTot = 0;
    var body = ""; var corps = ""; var bodyHTML = ""; var corpsHTML = ""; var summary = ""; var searchURL = ""; var searchName = "";
    var stop = false;
    while((searchURL = sheet.getRange(2+i,2).getValue()) != ""){
      searchName = sheet.getRange(2+i,1).getValue();
      Logger.log("Recherche pour " + searchName);
      var nbRes = 0;
      body = "";
      bodyHTML = "";
      stop = false;
      
      var response = UrlFetchApp.fetch(searchURL).getContentText("iso-8859-15");
      if(response.indexOf("Aucune annonce") < 0){  
        //l'url répond un contenu
        var data = extractAdsList_(response);
        data = data.replace(/<aside class=/gi, "<bside class=");
        debug_("data1 : " + data);
        data = data.substring(data.indexOf("<a"));
        debug_("data2 : " + data);
        var firsta = extractA_(data);
        
        if(sendMail == null || sendMail == true) {
          var holda = sheet.getRange(2+i,3).getValue();
          //est-ce que l'annonce est identique à la dernière retournée au précédent lancement du script ?
          if(extractId_(firsta) > holda) {// && holda != ""){  //if(extractId_(firsta) != holda) {// && holda != ""){
            while(data.indexOf("<a") > 0 || stop == false){
              var url = extractA_(data);
              if(extractId_(url) != holda){//if(extractId_(a) > holda){
                var title = extractTitle_(data);
                var img = extractImage_(data);
                var place = extractPlace_(data);
                var date = extractDate_(data);
                var placeAndDate = place + " / " + date;
                var price = extractPrice_(data);
                var seller = extractSeller_(data);
                body = body + "<li><a href=\"" + url + "\">" + title + "</a> (" + price + " euros - " + place + ")</li>";
                bodyHTML = bodyHTML + "<li style=\"list-style:none;margin-bottom:20px;clear:both;background:#EAEBF0;border-top:1px solid #ccc;\"><div style=\"width:400px;padding:10px 0;\"><a href=\"" + url + "\"><img src=\""+ img +"\"/></a><div style=\"float:left;width:400px;padding:10px 0;\"><a href=\"" + url + "\" style=\"font-size:14px;font-weight:bold;color:#369;text-decoration:none;\">" + title + seller +"</a><div>" + placeAndDate +"</div><div style=\"line-height:18px;font-size:14px;font-weight:bold;\">" + price + "</div></div></li>";
                if(data.indexOf("<a",10) > 0){
                  debug_("data3 : " + data);
                  var data = data.substring(data.indexOf("<a",10));
                  debug_("data4 : " + data);
                }else{
                  stop = true;
                  debug_("stop : " + stop);
                }
                nbRes++;
              }else{
                stop = true;
              }
            }
            nbSearchWithRes++;
            corps = corps + "<p>Votre recherche : <a name=\""+ searchName + "\" href=\""+ searchURL + "\"> "+ searchName +" (" + nbRes + ")</a></p><ul>" + body + "</ul>";
            corpsHTML = corpsHTML + "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Votre recherche : <a name=\""+ searchName + "\" href=\""+ searchURL + "\"> "+ searchName +" (" + nbRes + ")</a></p><ul>" + bodyHTML + "</ul>";
            summary += "<li><a href=\"#"+ searchName + "\">"+ searchName +" (" + nbRes + ")</a></li>"
            if(ScriptProperties.getProperty('log') == "true" || ScriptProperties.getProperty('log') == null || ScriptProperties.getProperty('log') == ""){
              slog.insertRowBefore(2);
              slog.getRange("A2").setValue(searchName);
              slog.getRange("B2").setValue(nbRes);
              slog.getRange("C2").setValue(new Date);
            }
            //sheet.getRange(2+i,3).setValue(extractId_(firsta));
          }
        }
        sheet.getRange(2+i,3).setValue(extractId_(firsta));
        nbResTot += nbRes;
      } else {
        //l'url retourne "aucune annonce" => il n'y pas de résultat => on set une value quelconque dans la sheet de suivi
        sheet.getRange(2+i,3).setValue(123);
      }
      //on passe à la recherche suivante
      i++;
    }
    debug_("Nb Search with result:" + nbSearchWithRes)
    debug_("Nb result tot:" + nbResTot);
    
    //fin des recherches, assemblage du mail
    if(nbSearchWithRes > 1) {
      //plusieurs recherches retourne des résultats => création d'un sommaire
      summary = "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Accès rapide :</p><ul>" + summary + "</ul>";
      debug_(summary);
      //corps = summary + corps;
      corpsHTML = summary + corpsHTML;
    }
    if(corps != ""){
      var title = "Alertes leboncoin.fr : " + nbResTot + " nouveau" + (nbResTot>1?"x":"") + " résultat" + (nbResTot>1?"s":"");
      debug_("titre msg : " + title);
      corps = "Si cet email ne s’affiche pas correctement, veuillez sélectionner\nl’affichage HTML dans les paramètres de votre logiciel de messagerie.";
      debug_("corps msg : " + corps);
      corpsHTML = "<body>" + corpsHTML + "</body>";
      debug_("corpsHTML msg : " + corpsHTML);
      MailApp.sendEmail(to,title,corps,{ htmlBody: corpsHTML });
      debug_("Nb mail journalier restant : " + MailApp.getRemainingDailyQuota());
    }
  }
}

/**
* Extrait ID de l'annonce
*/
function extractId_(id){
  var extractAdsId = id.substring(id.indexOf("/",25) + 1,id.indexOf(".htm"));
  debug_("extractAdsId : " + extractAdsId);
  return extractAdsId;
}

/**
* Extrait URL de l'annonce
*/
function extractA_(data){
  var extractAdsUrl = data.substring(data.indexOf("<a href=\"") + 9 , data.indexOf(".htm", data.indexOf("<a href=\"") + 9) + 4);
  extractAdsUrl = extractAdsUrl.replace("//","http://");
  debug_("extractAdsUrl : " + extractAdsUrl);
  return extractAdsUrl;
}

/**
* Extrait Titre de l'annonce
*/
function extractTitle_(data){
  var extractAdsTitle = data.substring(data.indexOf("title=") + 7 , data.indexOf("\"", data.indexOf("title=") + 7));
  debug_("extractAdsTitle : " + extractAdsTitle);
  return extractAdsTitle;
}

/**
* Extrait Type de vendeur (pro/part)
*/
function extractSeller_(data){
  var extractTypeSeller = data.substring(data.indexOf("\"ad_offres\" : \"") + 15 , data.indexOf("\"", data.indexOf("\"ad_offres\" : \"") + 15));
  debug_("extractTypeSeller : " + extractTypeSeller);
  if (extractTypeSeller.indexOf("part") == -1) {
    extractTypeSeller = " (" + extractTypeSeller + ")"
    return extractTypeSeller;
  }
  return "";
}

/**
* Extrait Lieu de l'annonce
*/
function extractPlace_(data){
  //Raccourcissement de la longueur de "data" pour faciliter l'extraction du lieu
  data = data.substring(data.indexOf("</p>") + 4 , data.indexOf("</aside>", data.indexOf("</p>") + 4));
  var extractAdsPlace = data.substring(data.indexOf("item_supp\">") + 11 , data.indexOf("</p>", data.indexOf("item_supp\">") + 11));
  debug_("extractAdsPlace : " + extractAdsPlace);
  return extractAdsPlace;
}

/**
* Extrait Prix de l'annonce
*/
function extractPrice_(data){
  //Raccourcissement de la longueur de "data" pour conserver 1 seule annonce et faciliter l'extraction du prix
  data = data.substring(data.indexOf("item_imagePic") + 13 , data.indexOf("</aside>", data.indexOf("item_imagePic") + 13));
  //Est-ce qu'il y a un prix sur l'annonce ?
  if (data.indexOf("item_price\">") > -1) {
    var extractAdsPrice = data.substring(data.indexOf("item_price\">") + 12 , data.indexOf("</h3>", data.indexOf("item_price\">") + 12));
    debug_("extractAdsPrice : " + extractAdsPrice);
    return extractAdsPrice;
  }
  return "";
}

/**
* Extrait Date de l'annonce
*/
function extractDate_(data){
  //Raccourcissement de la longueur de "data" pour faciliter l'extraction de la date
  data = data.substring(data.indexOf("item_absolute\">") + 15 , data.indexOf("</aside>", data.indexOf("item_absolute\">") + 15));
  var extractAdsDate = data.substring(data.indexOf("item_supp\">") + 12 , data.indexOf("</p>", data.indexOf("item_supp\">") + 12));
  extractAdsDate = extractAdsDate.replace("," , " / ");
  debug_("extractAdsDate : " + extractAdsDate);
  return extractAdsDate;
}

/**
* Extrait Image de l'annonce
*/
function extractImage_(data){
  //Raccourcissement de la longueur de "data" pour conserver 1 seule annonce et faciliter l'extraction de l'image
  data = data.substring(data.indexOf("item_imagePic") + 13 , data.indexOf("</aside>", data.indexOf("item_imagePic") + 13));
  //Est-ce qu'il y a une image sur cette annonce ?
  if (data.indexOf("no-picture.png") == -1) {
    var extractAdsImg = data.substring(data.indexOf("data-imgSrc=\"") + 13, data.indexOf("\"", data.indexOf("data-imgSrc=\"") + 13));
    //debug_("extractAdsImg : " + extractAdsImg);
    extractAdsImg = extractAdsImg.replace("//","http://");
    debug_("extractAdsImg : " + extractAdsImg);
    return extractAdsImg;
  }
  debug_("Pas d'image sur l'annonce");
  return "";
}

/**
* Extrait la liste des annonces
*/
function extractAdsList_(text){
var debut = text.indexOf("<ul class=\"tabsContent dontSwitch block-white\">");
  debug_("position debut liste des annonces : " + debut);

var fin = text.indexOf("<!-- Check the utility of this part -->");
  debug_("position fin liste des annonces : " + fin);
  
return text.substring(debut + "<ul class=\"tabsContent dontSwitch block-white\">".length,fin);
}

function setup(){
  lbc(false);
}

/**
* Definition du mail du destinataire
*/
function setupMail(){
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

/**
* Activer/Désactiver les logs
*/
function dolog(){
  if(ScriptProperties.getProperty('log') == "true"){
    ScriptProperties.setProperty('log', false);
    Browser.msgBox("Les logs ont été désactivées.");
  }else if(ScriptProperties.getProperty('log') == "false"){
    ScriptProperties.setProperty('log', true);
    Browser.msgBox("Les logs ont été activées.");
  }else{
    ScriptProperties.setProperty('log', false);
    Browser.msgBox("Les logs ont été désactivées.");
  }
}

/**
* Archiver les logs
*/
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

/**
* Fonction surOuverture de la feuille => ajout des menus lbc
*/
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
}];
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
