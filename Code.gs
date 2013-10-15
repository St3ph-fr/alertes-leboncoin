var COLUMN_LABEL = 1;
var COLUMN_URL = 2;
var COLUMN_LAST_ID = 3;

var CELL_EMAIL = "B5";

var ROW_ANNONCE = 8;

function alerteLeBonCoin() {

  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadSheet.getSheetByName("Alertes");
  
  var rowIndex = ROW_ANNONCE;
  var url;
  
  // Boucle sur les URL de recherche
  while((url = dataSheet.getRange(rowIndex,COLUMN_URL).getValue()) != "") {
    var label = dataSheet.getRange(rowIndex,COLUMN_LABEL).getValue();
    var lastId = dataSheet.getRange(rowIndex,COLUMN_LAST_ID).getValue();
    
    // Code HTML de la réponse
    var options = {"contentType" : "text/xml; charset=iso-8859-1"};
    var rep = UrlFetchApp.fetch(url, options).getContentText("ISO-8859-1");
    
    // Vérifie s'il y a des annonces
    if(rep.indexOf("Aucune annonce") < 0) {
      // Supprime tous les retours chariot
      var cleanRegexp = new RegExp("\\n", "g");
      rep = rep.replace(cleanRegexp, " ");

      // Extrait la partie concernant les annonces
      var infos = extractInfo('list-lbc(.*?)list-gallery', rep);

      // Extrait les annonces une à une
      var regexpAnnonce = new RegExp('<a href="(.*?)" title="(.*?)">(.*?)<\/a>', 'gi');
      var annonce = "";
      
      // ID de l'annonce en cours
      var idAnnonce = lastId;

      // Corps du message
      var mail = "";
    
      // Compteur d'annonces
      var compteur = 0;
    
      // Boucle sur chaque annonce
      while ((annonce = regexpAnnonce.exec(infos)) != null) {
         // Url de l'annonce 
         var href = annonce[1];
         
         // Id de l'annonce
         idAnnonce = extractInfo('\/(\\d*)\.htm', href);
         
         // Vérifie que l'annonce n'a pas déjà été envoyée
         if(idAnnonce == lastId) {
           break;
         }
         
         // Sauve l'ID de l'annonce la plus récente
         if(compteur == 0) {
           dataSheet.getRange(rowIndex,COLUMN_LAST_ID).setValue(idAnnonce);
         }

          // Titre
          var title = annonce[2];
          
          // Contenu entre les balises <a>...</a>
          var content = annonce[3];
          
          // Extrait la date
          var date = "";
          var regexpDate = new RegExp('<div class="date">\\s*<div>(.+?)<\/div>\\s*<div>(.+?)<\/div>\\s*<\/div>', 'gi');
          var found = regexpDate.exec(content);
          if(found) {
            date = found[1] + '<br>' + found[2];
          }
          
          // Extrait le lieu
          var placement = extractInfo('"placement">(.*?)<\/div>', content);
          
          // Extrait le prix
          var price = extractInfo('"price">(.*?)\&', content);
          
          // Extrait l'image
          var image = extractInfo('img src="(.*?)"', content);
          
          // ouverture du li
          mail += '<li style="list-style:none;margin-bottom:20px;clear:both;background:#EAEBF0;border-top:1px solid #ccc;">';

          // Construction du message
          mail += '<div style="float:left;width:90px;padding: 20px 20px 0 0;text-align: right;">' + date + '</div>';
          
          // Ajoute l'image
          if(image !=null) {
            mail += '<div style="float:left;width:200px;padding:20px 0;"><a href="' + href + '"><img src="' + image + '"/></a></div>';
          }
          
          // Ajoute le titre
          mail += '<div style="float:left;width:auto;padding:20px 0;"><a href="' + href + '" style="font-size: 14px;font-weight:bold;color:#369;text-decoration:none;">' + title + '</a>';
          
          // Ajoute le lieu
          mail += '<div>' + placement + '</div>';
         
          // Ajoute le prix
          if(price != null) {
            mail += '<div style="line-height:32px;font-size:14px;font-weight:bold;">' + price + " €" + '</div>';        
          }

          // fermeture du li
          mail += '</li>';
          
          compteur++;
      }    
      
      // Envoi du mail s'il y a de nouvelles annonces
      if(compteur > 0) {
        mail = '<p><b>Recherche</b> : <a href="'+ url + '"> '+ label + '</a><ul>' + mail + '</ul></p>';
        var email = dataSheet.getRange(CELL_EMAIL).getValue();
        MailApp.sendEmail(email, 'Alerte LeBonCoin : ' + label + ' - ' + Utilities.formatDate(new Date(), 'GMT+2', 'dd-MM-yyyy HH:mm'), mail, { htmlBody: mail });
      }
    }
    
    rowIndex++;
  }
}

function extractInfo(regexp, content) {
   regexp = new RegExp(regexp, 'gi');
   var found = regexp.exec(content);
   if(found) {
     found = found[1];
   }
   return found;
}

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Lancer manuellement",
    functionName : "alerteLeBonCoin"
  }];
  sheet.addMenu("Lbc Alertes", entries);
};
