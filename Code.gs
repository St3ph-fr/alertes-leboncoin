/**************************************
***************************************
**
** Alertes LeBonCoin
** => http://justdocsit.blogspot.fr/2012/11/alerte-leboncoin-v2.html
** => https://plus.google.com/u/0/116856005769817085204/posts
**
***************************************
**************************************/


/**
* global var section
*/
var menuLabel = "Lbc Alertes";
var menuWizardSetupLabel = "Assitant d'installation";
var menuMailSetupLabel = "Setup email";
var menuSearchSetupLabel = "Setup recherche";
var menuSearchLabel = "Lancer manuellement";
var menuClearLogLabel = "Effacer les logs";

/**
* consts
*/
var autoClearLog = false;
var maxLogNb = 10000;
var dataSheetName = "Données";
var logSheetName = "Log";
var emailPropertyName = "email";
var firstLaunchPropertyName = "wizard";

/**
* Recherche sur LeBonCoin.fr !
*/
function searchLbc_(sendMail){
  log_("#############################\n# Execution de la recherche #\n# "+myDate_()+"@"+myTime_()+" #\n#############################");
  var to = ScriptProperties.getProperty(emailPropertyName);
  if(sendMail && (to == "" || to == null) ){
    log_("L'email du destinataire n'est pas définit. Allez dans le menu \"" + menuLabel + "\" puis \"" + menuMailSetupLabel + "\".", true);
  } else {
    sendMail = (sendMail) ? (to != "" && to != null) : false;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(dataSheetName);
    var slog = ss.getSheetByName(logSheetName);
    var searchCount = 0; var nbSearchWithRes = 0; var nbResTot = 0;
    var body = ""; var corps = ""; var bodyHTML = ""; var corpsHTML = ""; var menu = ""; var searchURL = ""; var searchName = "";
    var stop = false;
    while((searchURL = sheet.getRange(2+searchCount,2).getValue()) != ""){
      try{
        searchName = sheet.getRange(2+searchCount,1).getValue();
        log_(">Recherche pour " + searchName);
        log_(" => url : " + searchURL);
        var nbRes = 0;
        body = "";
        bodyHTML = "";
        stop = false;
        var rep = UrlFetchApp.fetch(searchURL).getContentText("iso-8859-15");
        if(rep.indexOf("Aucune annonce") < 0){
          //LBC à des résultats
          log_(">Parcours des résultats");
          var data = extractResults_(rep);
          var annonces = data.split("<a");
          if(annonces.length > 1) {
            var firstid = "";
            var oldid = sheet.getRange(2+searchCount,3).getValue();
            for (var annonceCount = 1; annonceCount < annonces.length; annonceCount++) {
              var annonce = "<a" + annonces[annonceCount]; //on rajoute "<a" à cause du split !!! (ok, on aurait pu faire sur "</a"   ;)
              var a = extractA_(annonce);
              if(firstid == "") {
                firstid = extractId_(a);
              }
              var currid = extractId_(a);
              if(currid != oldid) {
                if(sendMail) {
                  //extraction des informations de l'annonce
                  var title = extractTitle_(annonce);
                  var place = extractPlace_(annonce);
                  var price = extractPrice_(annonce);
                  var date = extractDate_(annonce);
                  var image = extractImage_(annonce);
                  body = body + "<li><a href=\"" + a + "\">" + title + "</a> (" + price + " euros - " + place + ")</li>";
                  bodyHTML = bodyHTML + "<li style=\"list-style:none;margin-bottom:20px; clear:both;background:#EAEBF0;border-top:1px solid #ccc;\"><div style=\"float:left;width:90px;padding: 20px 20px 0 0;text-align: right;\">"+ date +"<div style=\"float:left;width:200px;padding:20px 0;\"><a href=\"" + a + "\">"+ image +"</a> </div><div style=\"float:left;width:420px;padding:20px 0;\"><a href=\"" + a + "\" style=\"font-size: 14px;font-weight:bold;color:#369;text-decoration:none;\">" + title + "</a> <div>" + place + "</div> <div style=\"line-height:32px;font-size:14px;font-weight:bold;\">" + price + "</div></div></li>";
                }
                nbRes++;
              } else {
                break;
              }
            }
            sheet.getRange(2+searchCount,3).setValue(firstid);
            
            if(nbRes>0) {
              nbSearchWithRes++;
              nbResTot += nbRes;
              
              if(sendMail) {
                var searchNameNlz = searchName.replace(" ", "_");
                //création du corps du message
                var resLbl = searchName +" (" + nbRes + " résultat"+((nbRes>0)?"s":"")+")";
                corps = corps + "<p>Votre recherche : <a name=\""+ searchNameNlz + "\" href=\""+ searchURL + "\"> " + resLbl + "</a></p><ul>" + body + "</ul>";
                corpsHTML = corpsHTML + "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Votre recherche : <a name=\""+ searchName + "\" href=\""+ searchURL + "\"> "+ searchName +" (" + nbRes + ")</a></p><ul>" + bodyHTML + "</ul>";
                menu += "<li><a href=\"#"+ searchNameNlz + "\">"+ resLbl + "</a></li>"
                
                //ajout de la ligne dans les logs de la recherche
                slog.insertRowBefore(2);
                slog.getMaxRows()
                slog.getRange("A2").setValue(searchName);
                slog.getRange("B2").setValue(nbRes);
                slog.getRange("C2").setValue(new Date);
                //sheet.getRange(2+i,3).setValue(extractId_(firsta));
              }
            }
            log_(">Nb res: " + nbRes);
          }
        } else {
          //pas de résultat
          log_(">Pas de résultat");
          sheet.getRange(2+searchCount,3).setValue(123);
        }
      } catch (err) {
        log_("!!Erreur :\n"+err, true);
      }
      searchCount++;
      log_("----------------------------");
    }
    
    if(autoClearLog) {
      //nettoyage des logs
      clearLog(maxLogNb);
    }
    
    if(sendMail && nbSearchWithRes > 1) {
      //plusieurs recherche, on créé un menu
      menu = "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Accès rapide :</p><ul>" + menu + "</ul>";
      //corps = menu + corps;
      corpsHTML = menu + corpsHTML;
    }
    
    log_(">>Nb de res tot:" + nbResTot);
    log_(">>sendMail:" + sendMail + "(mail:"+to+")");
    //on envoie le mail?
    if(sendMail && nbResTot > 0) {
      var title = "Alerte leboncoin.fr : " + nbResTot + " nouveau" + (nbResTot>1?"x":"") + " résultat" + (nbResTot>1?"s":"");
      corps = "Si cet email ne s’affiche pas correctement, veuillez sélectionner\nl’affichage HTML dans les paramètres de votre logiciel de messagerie.";
      corpsHTML = "<body>" + corpsHTML + "</body>";
      
      try {
        log_(">>Envoie du mail à " + to);
        //Browser.msgBox(corpsHTML);
        MailApp.sendEmail(to,title,corps,{ htmlBody: corpsHTML });
      }
      catch (errMail) {
        log_("!!Erreur lors de l'envoie du mail :\n"+errMail, true);
      }
    }
    
    log_(">>Nb mail journailier restant : " + MailApp.getRemainingDailyQuota());
  }
}

function searchLbc(){
  searchLbc_(true);
}

function initSearchId(){
  searchLbc_(false);
  log_("Fin de l'initialisation des recherches.", true);
}

function setupMail(){
  log_("#############################\n#      MaJ de l'email       #\n#############################");
  var quest;
  if(ScriptProperties.getProperty(emailPropertyName) == "" || ScriptProperties.getProperty(emailPropertyName) == null ){
    quest = Browser.inputBox("Entrez votre email, le programme ne vérifie pas le contenu de cette boite.", Browser.Buttons.OK_CANCEL);
    if(quest == "cancel"){
      log_("Ajout email annulé.", true);
      return false;
    }
  }else{
    quest = Browser.inputBox("Entrez un email pour modifier l'email : " + ScriptProperties.getProperty(emailPropertyName) , Browser.Buttons.OK_CANCEL);
    if(quest == "cancel"){
      log_("Modification email annulé.", true);
      return false;
    }
  }
  //validation du format de l'email...
  /*if(!quest.validateMatches("")) {
      log_("Votre email n'est pas au bon format...", true);
      return false;
  }*/
  ScriptProperties.setProperty(emailPropertyName, quest);
  log_("Email " + ScriptProperties.getProperty(emailPropertyName) + " ajouté", true);
  return true;
}

/**
* supp des log en trop
*/
function clearLog(maxLogRows) {
  if(maxLogRows == null) {
    log_("#############################\n# Suppression des logs      #\n#############################");
    maxLogRows = 0;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var slog = ss.getSheetByName("Log");
  var start = 2+maxLogRows;
  var end = slog.getLastRow();
  if(start < end) {
    //supp des logs en trop
    slog.deleteRows(start, end);
    log_("Supp des logs de " + start + " à " + end);
  }
}

/**
* Extrait la liste des annonces
*/
function extractResults_(text){
  var debut = text.indexOf("<div class=\"list-lbc\">");
  var fin = text.indexOf("<div class=\"list-gallery\">");
  return text.substring(debut + "<div class=\"list-lbc\">".length, fin);
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
  var isPrice = String(data.substring(data.indexOf("price"), data.indexOf("price")+250)).match(/price/gi);
  if (isPrice) {
    var price = data.substring(data.indexOf("price") + 7 , data.indexOf("</div>", data.indexOf("price") + 7) );
  } else {
    var price = "";
  }
  return price;
}

/**
* Extrait tla date de l'annonce
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
* à l'ouverture du fichier
*/
function onOpen() {
  
  log_("#############################\n# Création du menu et aide  #\n#############################");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //création du menu
  var menuEntries = [];
  menuEntries.push({name: menuMailSetupLabel, functionName: "setupMail"});
  menuEntries.push({name: menuSearchSetupLabel, functionName: "initSearchId"});
  menuEntries.push(null); // line separator
  menuEntries.push({name: "=> " + menuSearchLabel, functionName: "searchLbc"});
  menuEntries.push(null); // line separator
  menuEntries.push({name: menuWizardSetupLabel, functionName: "wizard"});
  menuEntries.push({name: menuClearLogLabel, functionName: "clearLog"});
  ss.addMenu(menuLabel, menuEntries);
  
  //maj du mod'op
  var sheet = ss.getSheetByName(dataSheetName);
  for(var rowIndex = 10; rowIndex < 1000; rowIndex++) {
    if(sheet.getRange(rowIndex, 1).getValue() == "Comment ça marche ?") {
      sheet.getRange(rowIndex, 2).setValue("Allez sur la feuille \""+dataSheetName+"\", ajoutez les urls de vos recherche sur leboncoin dans la colonne B.");
      sheet.getRange(rowIndex+1, 2).setValue("Donnez un titre à votre recherche, colonne A.");
      sheet.getRange(rowIndex+2, 2).setValue("Dans le menu \""+menuLabel+"\" :\nCliquez sur \""+menuMailSetupLabel+"\", lors du premier lancement il faudra accepter les autorisations demandées par Google. C'est normal, acceptez et cliquez de nouveau sur \""+menuMailSetupLabel+"\".");
      sheet.getRange(rowIndex+3, 2).setValue("Dans le menu \""+menuLabel+"\" :\nCliquez sur \""+menuSearchSetupLabel+"\" pour initialiser les dernières annonces de LeBonCoin.fr (cette étape n'est pas obligatoire)");
      sheet.getRange(rowIndex+4, 2).setValue("Dans le menu \""+menuLabel+"\" :\nTestez le script en cliquant sur \""+menuSearchLabel+"\" => la colonne C \""+sheet.getRange(1, 3).getValue()+"\" se met à jour (si de nouvelles annonces ont été trouvées)");
      sheet.getRange(rowIndex+5, 2).setValue("Automatisez la recherche en allant dans le menu \"Outils/Éditeur de script.../Ressources/Déclencheurs du projet actuel...\"\nEt paramétrez la fréquence sur la fonction \"searchLbc()\"");
      break;
    }
  }
  
  wizard_(ScriptProperties.getProperty(firstLaunchPropertyName) == null);
}

/**
* à l'installation à partir de Script Gallery.
*/
function onInstall()
{
  onOpen();
}

function wizard() {
  wizard_(true);
}


/**
* Wizard d'installation du script
*/
function wizard_(executeWizard)
{
  if(executeWizard){
    log_("#############################\n#    Execution du wizard    #\n#############################");
    var quest = Browser.msgBox("Voulez-vous exécuter le wizard d'installation ?" , Browser.Buttons.YES_NO);
    if(quest == "yes"){
      if(setupMail()) {
        initSearchId();
      }
    }
  }
  ScriptProperties.setProperty(firstLaunchPropertyName, false);
}

/**
* Retourne la date
*/
function myDate_(){
  var today = new Date();
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
  return h+":"+m;
}

/**
* Logger
*/
function log_(msg, msgBox) {
  Logger.log(msg);
  if(msgBox != null && msgBox) {
    Browser.msgBox(msg);
  }
}
