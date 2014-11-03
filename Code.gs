/**
* Alertes LBC V2.2
*/

var debug = false;

/**
* Labels des menus
*/
var menuLabel = "Alertes";
var menuSearchSetupLabel = "Setup recherche";
var menuSearchLabel = "Lancer manuellement";
var menuLog = "Activer/Désactiver les logs";
var menuArchiveLog = "Archiver les logs";

/**
* Traitement LeBonCoin
*/
function lbc(sendMail){
  if(sendMail != false){
    sendMail = true;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Données");
  var slog = ss.getSheetByName("Log");
  var i = 0;
  var bodyHTML = ""; var corpsHTML = ""; var searchURL = ""; var searchName = "";
  var stop = false;
  
  var alertes = {};
  while((searchURL = sheet.getRange(2+i,2).getValue()) != ""){
    searchName = sheet.getRange(2+i,1).getValue();
    var dest = sheet.getRange(2+i,4).getValue();
    
    if (dest != "") {
      if (!alertes[dest]) {
        alertes[dest] = {};
        alertes[dest]["body"] = "";
        alertes[dest]["nbResult"] = 0;
        alertes[dest]["nbAnnonce"] = 0;
        alertes[dest]["menu"] = "";
      }
      
      Logger.log("Recherche pour " + searchName);
      stop = false;
      var rep = UrlFetchApp.fetch(searchURL).getContentText("iso-8859-15");
      if(rep.indexOf("Aucune annonce") < 0){
        // LBC à des résultats
        Logger.log("Présence de résultats");
        var data = splitResult_(rep);
        data = data.substring(data.indexOf("<a"));
        var firsta = extractA_(data);
        if(sendMail == null || sendMail == true) {
          var holda = sheet.getRange(2+i,3).getValue();
          
          var annonceBody = "";
          if(extractId_(firsta) != holda) {
            var nbRes = 0;
            while(data.indexOf("<a") > 0 || stop == false){
              
              
              var a = extractA_(data);
              if(extractId_(a) != holda){
                var title = extractTitle_(data);
                var place = extractPlace_(data);
                var price = extractPrice_(data);
                var vendpro = extractPro_(data);
                var date = extractDate_(data);
                var image = extractImage_(data);
                annonceBody = annonceBody + "<li style=\"list-style:none;margin-bottom:20px; clear:both;background:#EAEBF0;border-top:1px solid #ccc;\"><div style=\"float:left;width:90px;padding: 20px 20px 0 0;text-align: right;\">"+ date +"<div style=\"float:left;width:200px;padding:20px 0;\"><a href=\"" + a + "\">"+ image +"</a> </div><div style=\"float:left;width:420px;padding:20px 0;\"><a href=\"" + a + "\" style=\"font-size: 14px;font-weight:bold;color:#369;text-decoration:none;\">" + title + vendpro +"</a> <div>" + place + "</div> <div style=\"line-height:32px;font-size:14px;font-weight:bold;\">" + price + "</div></div></li>";
                
                if(data.indexOf("<a",10) > 0){
                  var data = data.substring(data.indexOf("<a",10));
                }else{
                  stop = true;
                }
                nbRes++;
                alertes[dest]["nbResult"] = alertes[dest]["nbResult"] + 1;
              }else{
                stop = true;
              }
            }
            
            
            alertes[dest]["body"] += "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Votre recherche : <a name=\""+ searchName + "\" href=\""+ searchURL + "\"> "+ searchName +" (" + nbRes + ")</a></p><ul>" + annonceBody + "</ul>";
            alertes[dest]["menu"] += "<li><a href=\"#"+ searchName + "\">"+ searchName +" (" + nbRes + ")</a></li>"
            
            alertes[dest]["nbAnnonce"] = alertes[dest]["nbAnnonce"] + 1;
            
            if(ScriptProperties.getProperty('log') == "true" || ScriptProperties.getProperty('log') == null || ScriptProperties.getProperty('log') == ""){
              slog.getRange(2+i, 1).setValue(searchName);
              slog.getRange(2+i, 2).setValue(nbRes);
              slog.getRange(2+i, 3).setValue(new Date);
              slog.getRange(2+i, 4).setValue(new Date);
              slog.getRange(2+i,1,1,4).setBorder(true,true,true,true,true,true);
            }
          } else {
            Logger.log("Pas de nouvelles annonces"); 
            if(ScriptProperties.getProperty('log') == "true" || ScriptProperties.getProperty('log') == null || ScriptProperties.getProperty('log') == ""){
              slog.getRange(2+i, 1).setValue(searchName);
              slog.getRange(2+i, 4).setValue(new Date);
              slog.getRange(2+i,1,1,4).setBorder(true,true,true,true,true,true);
            }
          }
        }
        sheet.getRange(2+i,3).setValue(extractId_(firsta));
      } else {
        Logger.log("Pas de résultat");
        // Pas de résultat
        sheet.getRange(2+i,3).setValue(123);
      }
      i++;
    }
  }
  
  for(var dest in alertes) {
    if(ScriptProperties.getProperty('log') == "true" || ScriptProperties.getProperty('log') == null || ScriptProperties.getProperty('log') == ""){
      slog.getRange("E2").setValue(MailApp.getRemainingDailyQuota());
    }
    //on envoie le mail?
    if(alertes[dest]["body"] != ""){
      if(alertes[dest]["nbAnnonce"] > 1) {
        // Plusieurs recherches, on créée un menu
        var menu = "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Accès rapide :</p><ul>" + alertes[dest]["menu"] + "</ul>";
        //corps = menu + corps;
        alertes[dest]["body"] = menu + alertes[dest]["body"];
        debug_(menu);
      }
      
      var nbResTot = alertes[dest]["nbResult"];
      var title = "Alerte leboncoin.fr : " + nbResTot + " nouveau" + (nbResTot>1?"x":"") + " résultat" + (nbResTot>1?"s":"");
      debug_("titre msg : " + title);
      var corps = "Si cet email ne s’affiche pas correctement, veuillez sélectionner\nl’affichage HTML dans les paramètres de votre logiciel de messagerie.";
      corpsHTML = "<body>" + alertes[dest]["body"] + "</body>";
      debug_("corpsHTML msg : " + corpsHTML);
      MailApp.sendEmail(dest,title,corps,{ htmlBody: corpsHTML });
    }
  }

}

function setup(){
  lbc(false);
}

/**
* Extrait l'id de l'annonce LBC
*/
function extractId_(id){
  return id.substring(id.indexOf("/",25) + 1,id.indexOf(".htm"));
}

/**
* Extrait le lien de l'annonce
*/
function extractA_(data){
  return data.substring(data.indexOf("<a") + 9 , data.indexOf(".htm", data.indexOf("<a") + 9) + 4);
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
  var pro = data.substring(data.indexOf("category") + 9 , data.indexOf("</div>", data.indexOf("category") + 9) );
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
  return data.substring(data.indexOf("placement") + 11 , data.indexOf("</div>", data.indexOf("placement") + 11) );
}

/**
* Extrait le prix de l'annonce
*/
function extractPrice_(data){
  // test à optimiser car c'est hyper bourrin [mlb]
  data = data.substring(0,data.indexOf("clear",10)); //racourcissement de la longueur de data pour ne pas aller chercher le prix du proudit suivant
  var isPrice = String(data.substring(data.indexOf("price"), data.indexOf("price")+250)).match(/price/gi);
  if (isPrice) {
    var price = data.substring(data.indexOf("price") + 7 , data.indexOf("</div>", data.indexOf("price") + 7) );
  } else {
    var price = "";
  }
  return price;
}

/**
* Extrait la date de l'annonce
*/
function extractDate_(data){
return data.substring(data.indexOf("date") + 6 , data.indexOf("class=\"image\"", data.indexOf("date") + 6) - 5);
}

/**
* Extrait l'image de l'annonce
*/
function extractImage_(data){
// test à optimiser car c'est hyper bourrin [mlb]
var isImage = String(data.substring(data.indexOf("image"), data.indexOf("image")+250)).match(/img/gi);
if (isImage) {
var image = data.substring(data.indexOf("class=\"image-and-nb\">") + 21, data.indexOf("class=\"nb\"", data.indexOf("class=\"image-and-nb\">") + 21) - 12);
} else {
var image = "";
}
return image;
}

/**
* Extrait la liste des annonces
*/
function splitResult_(text){
var debut = text.indexOf("<div class=\"list-lbc\">");
var fin = text.indexOf("<div class=\"list-gallery\">");
return text.substring(debut + "<div class=\"list-lbc\">".length,fin);
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
  
  newsheet.getRange("B1").setValue("Nouvelles annonces détectées");
  newsheet.getRange("C1").setValue("Date mail");
  newsheet.getRange("D1").setValue("Dernière exécution");
  newsheet.getRange("E1").setValue('Nombre de mails restants');
  newsheet.getRange(1,1,2,5).setBorder(true,true,true,true,true,true);
  newsheet.getRange(1,1,1,5).setFontWeight("bold");
  newsheet.getRange(1,1,1,5).setHorizontalAlignment("center");
  newsheet.getRange(1,1,1,5).setVerticalAlignment("middle");
  newsheet.getRange(1,1,1,5).setBackgroundRGB(243, 243, 243);
}


/**
* Initialisation du menu
*/
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
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

    }
  ];
  sheet.addMenu(menuLabel, entries);
}

/**
* Initialisation du menu
*/
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

