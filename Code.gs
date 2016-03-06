var nbMaxSearches = 10;
var rowSearchTitles = 1;
var rowResTitles = 9;


function reset()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.deleteRows(10, sheet.getLastRow()-9);
};


function indexOfAfter( data, str, start )
{
    var idx = data.indexOf(str, start);
    if ( idx > 0 ) idx += str.length;
    return idx;
};



/**
* global var section
*/
var menuLabel = "Lbc Alertes";
var menuMailSetupLabel = "Setup email";
var menuSearchLabel = "Lancer manuellement";
var menuLog = "Activer/Désactiver les logs";
var menuArchiveLog = "Archiver les logs";

function initSheets(ss)
{
  var sheets = [];
  
  for ( indexSheet = 1 ; indexSheet<=nbMaxSearches ; indexSheet+=1 )
  {
    var sheet = ss.getSheetByName("Recherche "+indexSheet);
    
    // Boucle sur chaque recherche
    if ( sheet != null )
    {
      var sheetValues = sheet.getDataRange().getValues();
      
      // Init des numéro de colonnes en fonction du libellé
      var searchColName = -1;
      var searchColUrl = -1;
      var searchColMinPrice = -1;
      var searchColMaxPrice = -1;
      var searchColLastExec = -1;
      var searchColNbResults = -1;
      var searchColNbResultsMatch = -1;
      for ( i = 1 ; i < 20 ; i += 1 ) {
        if ( sheetValues[rowSearchTitles-1][i-1] == "Libellé recherche" ) searchColName = i;
        if ( sheetValues[rowSearchTitles-1][i-1] == "Url" ) searchColUrl = i;
        if ( sheetValues[rowSearchTitles-1][i-1] == "Prix mini" ) searchColMinPrice = i;
        if ( sheetValues[rowSearchTitles-1][i-1] == "Prix maxi" ) searchColMaxPrice = i;
        if ( sheetValues[rowSearchTitles-1][i-1] == "Dernière execution" ) searchColLastExec = i;
        if ( sheetValues[rowSearchTitles-1][i-1] == "Nb résultats" ) searchColNbResults = i;
        if ( sheetValues[rowSearchTitles-1][i-1] == "Nb résultats avec critères" ) searchColNbResultsMatch = i;
      }
      if ( searchColName == -1 ) throw "Impossible de trouver la colonne " + "searchColName";
      if ( searchColUrl == -1 ) throw "Impossible de trouver la colonne " + "searchColUrl";
      if ( searchColMinPrice == -1 ) throw "Impossible de trouver la colonne " + "searchMinPrice";
      if ( searchColMaxPrice == -1 ) throw "Impossible de trouver la colonne " + "searchMaxPrice";
      if ( searchColLastExec == -1 ) throw "Impossible de trouver la colonne " + "searchColLastExec";
      if ( searchColNbResults == -1 ) throw "Impossible de trouver la colonne " + "searchColNbResults";
      if ( searchColNbResultsMatch == -1 ) throw "Impossible de trouver la colonne " + "searchColNbResultsMatch";
      
      var resColId = -1;
      var resColUrl = -1;
      var resColPrice = -1;
      var resColLastSeen = -1;
      var resColMailSent = -1;
      var resColMatchCriteria = -1;
      var resColPublishedDate = -1;
      
      for ( i = 1 ; i < 20 ; i += 1 ) {
        if ( sheetValues[rowResTitles-1][i-1] == "Ad id" ) resColId = i;
        if ( sheetValues[rowResTitles-1][i-1] == "Url" ) resColUrl = i;
        if ( sheetValues[rowResTitles-1][i-1] == "Prix" ) resColPrice = i;
        if ( sheetValues[rowResTitles-1][i-1] == "Last seen" ) resColLastSeen = i;
        if ( sheetValues[rowResTitles-1][i-1] == "Mail sent" ) resColMailSent = i;
        if ( sheetValues[rowResTitles-1][i-1] == "Match criteria" ) resColMatchCriteria = i;
        if ( sheetValues[rowResTitles-1][i-1] == "Date mise en ligne" ) resColPublishedDate = i;
      }
      if ( resColId == -1 ) throw "Impossible de trouver la colonne " + "resColId";
      if ( resColUrl == -1 ) throw "Impossible de trouver la colonne " + "resColUrl";
      if ( resColPrice == -1 ) throw "Impossible de trouver la colonne " + "resColPrice";
      if ( resColLastSeen == -1 ) throw "Impossible de trouver la colonne " + "resColLastSeen";
      if ( resColMailSent == -1 ) throw "Impossible de trouver la colonne " + "resColMailSent";
      if ( resColMatchCriteria == -1 ) throw "Impossible de trouver la colonne " + "resColMatchCriteria";
      if ( resColPublishedDate == -1 ) throw "Impossible de trouver la colonne " + "resColPublishedDate";
      
      sheets.push(
        {
          sheet: sheet,
          values: sheetValues,
          searchColName: searchColName,
          searchColUrl: searchColUrl,
          searchColMinPrice: searchColMinPrice,
          searchColMaxPrice: searchColMaxPrice,
          searchColLastExec: searchColLastExec,
          searchColNbResults: searchColNbResults,
          searchColNbResultsMatch: searchColNbResultsMatch,
          resColId: resColId,
          resColUrl: resColUrl,
          resColPrice: resColPrice,
          resColLastSeen: resColLastSeen,
          resColMailSent: resColMailSent,
          resColMatchCriteria: resColMatchCriteria,
          resColPublishedDate: resColPublishedDate
        }
      );
    }
  }
  return sheets;
}

function lbc(sendMail, throwErrorByMail)
{
  var to = ScriptProperties.getProperty('email');
  if ( sendMail  &&  (to == "" || to == null)  ) {
    //L'erreur suivante remonte jusqu'au navigateur puisqu'elle n'est pas catchée.
    throw new Error("L'email du destinataire n'est pas définit. Allez dans le menu \"" + menuLabel + "\" puis \"" + menuMailSetupLabel + "\".");
  }

  try
  {
    var now = new Date();
    if ( sendMail != false ) {
      sendMail = true;
    }

    
    
    var nbSearchWithRes = 0;
    var nbSearchWithResultsSinceLastEmail = 0;
    var nbResultWithCriteriaSinceLastEmail = 0;
    var nbResTot = 0;
    var nbResTotWithCriteria = 0;
    
    var corpsHTML = "";
    var menu = "";
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetObjArray = initSheets(ss);
    var slog = ss.getSheetByName("Log");
    
    for ( sheetsIndex = 0 ; sheetsIndex < sheetObjArray.length ; sheetsIndex += 1 )
    {
      var sheetObj = sheetObjArray[sheetsIndex];
      var sheet = sheetObj.sheet;
      var sheetValues = sheetObj.values;
      var searchName = sheetValues[rowSearchTitles+1-1][sheetObj.searchColName-1];
      var searchURL = sheetValues[rowSearchTitles+1-1][sheetObj.searchColUrl-1];
      searchURL = searchURL.replace(/sp=1/g, "sp=0");
      searchURL = searchURL.replace(/o=\d*/g, "o=1");
      //searchURL += "&sp=0&o=1"
      var prixMini = sheetValues[rowSearchTitles+1-1][sheetObj.searchColMinPrice-1];
      var prixMaxi = sheetValues[rowSearchTitles+1-1][sheetObj.searchColMaxPrice-1];
      Logger.log("Recherche pour " + searchName);
      
      var nbResForThisSearch = 0;
      var nbResultWithCriteriaForThisSearch = 0;
      var nbResultWithCriteriaSinceLastEmailForThisSearch = 0;
      //var currentPage = 1;
      var inserRowBefore = rowResTitles+1;
      
      var bodyHTMLForThisSearch = "";
      
      try {
        
        // Pour chaque page de résultat
        // Pour l'instant, le multi page est désactivé
        do
        {
          //var rep = UrlFetchApp.fetch(searchURL+"&o="+currentPage).getContentText("iso-8859-15");
          // dorénavant, on ne parcours que la première page de résultat.
          var rep = UrlFetchApp.fetch(searchURL).getContentText("iso-8859-15"); // ajoute sp=0 pour trier par date et &o=1 pour la première page
          if ( rep.match("Aucune annonce de professionnel n'a été trouvée!") ) throw "Dépassement de dernière page ! BUG !";
          
          if ( rep.indexOf("Aucune annonce") < 0 )
          {
            var dataList = splitResult_(rep); // enlever ce qu'il y a avant et après la liste des annonces
            
            // prendre la première annonce
            var idxOfAdStart = dataList.indexOf('<li>');
            if ( idxOfAdStart == -1 ) throw "Impossible de trouver <li> qui marque la première annonce. Il devrait y avoir au moins un résultat. searchURL="+searchURL;
            
            //if ( sheetsIndex == 0 ) {
            //  throw "DEBUG zeezez ezeeezez f zezeze zezeze zezez";
            //}
            
            // Pour chaque résultat dans la page
            while( idxOfAdStart > 0 )
            {
              var idxOfAdEnd = indexOfAfter(dataList, '</li>', idxOfAdStart);
              if ( idxOfAdEnd == -1 ) throw "Impossible de trouver </li> qui marque la fin de l'annonce. searchURL="+searchURL;
              
              var currentAdHref = "";
              var id = "";
              var title = "";
              var place = "";
              var priceAsString = "";
              var price = 0;
              var vendpro = "";
              var dateHtml = "";
              var dateZZZ = null;
              var image = "";
              var errorMsg = "";
              var errorLine = -1;
              
              var emailSent = false;
              
              try
              {
                var dataAd = dataList.substring(idxOfAdStart, idxOfAdEnd);
                
                
                currentAdHref = extractHref_(dataAd);
                id = extractId_(currentAdHref);
                title = extractTitle_(dataAd);
                place = extractPlace_(dataAd);
                priceAsString = extractPrice_(dataAd);
                price = parseInt(priceAsString.replace(/\s/g, ""));
                vendpro = extractPro_(dataAd);
                dateHtml = extractHtmlDate_(dataAd); // une chaine contenant, par exemple : Aujourd'hui, 16:47
                dateZZZ = extractDate_(dataAd);
                image = extractImage_(dataAd);
                
                //if ( id == "821936425" ) {
                //  var dummy = 0;
                //  throw "DEBUG sdlkfjsqdf sdflkqsdjfsdq f qsdflkqsdfqsd fqsdfdsqflkqsdjf sfsqdfqdslkf";
                //}
              }
              catch (e) {
                errorMsg = e;
                errorLine = e.lineNumber;
              }
              
              if ( errorMsg == '' )
              {
                var statusRowIndex = -1;
                for ( i = 0 ; statusRowIndex == -1  &&  i < sheetValues.length ; i+=1 ) {
                  if ( sheetValues[i][sheetObj.resColId-1] == id ) statusRowIndex = i+1;
                }
                if ( statusRowIndex == -1 ) {
                  sheetValues.splice(inserRowBefore-1, 0, [id, ""]);
                  statusRowIndex = inserRowBefore;
                  inserRowBefore += 1;
                }
                sheetValues[statusRowIndex-1][sheetObj.resColUrl-1] = "http://" + currentAdHref;
                var priceStored = sheetValues[statusRowIndex-1][sheetObj.resColPrice-1];
                if ( isNaN(price) && priceStored != ""  ||  !isNaN(price) && priceStored != price ) {
                  // price modified, re send email
                  sheetValues[statusRowIndex-1][sheetObj.resColMailSent-1] = "";
                }
                if ( !isNaN(price) ) {
                  sheetValues[statusRowIndex-1][sheetObj.resColPrice-1] = price;
                }else{
                  sheetValues[statusRowIndex-1][sheetObj.resColPrice-1] = "";
                }
                sheetValues[statusRowIndex-1][sheetObj.resColLastSeen-1] = now;
                if ( sheetObj.resColPublishedDate != -1 ) sheetValues[statusRowIndex-1][sheetObj.resColPublishedDate-1] = dateZZZ;
                
                var emailSent = sheetValues[statusRowIndex-1][sheetObj.resColMailSent-1] != "";
              }
              
              
              
              if ( !emailSent )
              {
                if ( errorMsg !== ''  ||  isNaN(price)  ||  (price >= prixMini && price <= prixMaxi) ) // si erreur, génération systématique de l'email
                { 
                  if ( errorMsg == '' ) // si erreur, statusRowIndex pourrait n'avoir pas été initalisé.
                  {
                    sheetValues[statusRowIndex-1][sheetObj.resColMatchCriteria-1] = "Yes";
                  }
                  
                  // bodyHTMLForThisSearch
                  // Ouverture de l'élément <li>
                  bodyHTMLForThisSearch += '<li style="list-style:none;margin-bottom:20px; clear:both;background:#EAEBF0;border-top:1px solid #ccc;">'
                  {
                    if ( errorMsg !== '' )
                    {
                      bodyHTMLForThisSearch += 'ERREUR : ' + errorMsg + '  - ligne:'+errorLine;
                    }
                    
                    // Création du div contenant l'image cliquable
                    bodyHTMLForThisSearch += '<div style="float:left;width:180px;padding:20px 0;"><a href="' + currentAdHref + '"><img src="'+ image +'"></a> </div>';
                    
                    // Ouverture du div contenant le titre, le lieu, le prix
                    bodyHTMLForThisSearch += "<div style=\"float:left;width:200px;padding:20px 0;\">"; //;width:420px
                    {
                      // Titre
                      bodyHTMLForThisSearch += '<a href="' + currentAdHref + '" style="font-size: 14px;font-weight:bold;color:#369;text-decoration:none;">' + title +'</a>';
                      // pro
                      bodyHTMLForThisSearch += '<div>' + vendpro +'</div>';
                      // Lieu
                      bodyHTMLForThisSearch += "<div>" + place + "</div>";
                      // Prix
                      bodyHTMLForThisSearch += '<div style="float:left;line-height:32px;font-size:14px;font-weight:bold;">' + price + '&nbsp;&euro;</div>';
                      // Création du div contenant la date
                      bodyHTMLForThisSearch += '<div style="float:right;line-height:32px;text-align:right;">' + dateHtml + '</div>';
                    }
                    // Fermeture  du div contenant le titre, le lieu, le prix
                    bodyHTMLForThisSearch += "</div>";
                  }
                  // Fermeture de <li>
                  bodyHTMLForThisSearch += "</li>";            
                  
                  nbResultWithCriteriaForThisSearch++;
                  nbResultWithCriteriaSinceLastEmailForThisSearch++;
                  nbResultWithCriteriaSinceLastEmail++;
                  nbResTotWithCriteria++;
                  if ( errorMsg == '' ) // si erreur, statusRowIndex pourrait n'avoir pas été initalisé.
                    sheetValues[statusRowIndex-1][sheetObj.resColMailSent-1] = "Yes";
                }
                else{
                  if ( errorMsg == '' ) // si erreur, statusRowIndex pourrait n'avoir pas été initalisé.
                    sheetValues[statusRowIndex-1][sheetObj.resColMatchCriteria-1] = "No";
                }
              }
              
              nbResForThisSearch++;
              nbResTot++;
              
              if ( corpsHTML.length + bodyHTMLForThisSearch.length > 180000 ) {
                if ( nbResultWithCriteriaSinceLastEmail > 0 ) {
                  nbSearchWithResultsSinceLastEmail += 1;
                  corpsHTML = corpsHTML + "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Votre recherche : <a name=\""+ searchName + "\" href=\""+ searchURL + "\"> "+ searchName +" (" + nbResultWithCriteriaSinceLastEmailForThisSearch + ")</a></p><ul>" + bodyHTMLForThisSearch + "</ul>";
                  menu += "<li><a href=\"#"+ searchName + "\">"+ searchName +" (" + nbResultWithCriteriaSinceLastEmailForThisSearch + ")</a></li>"
                  if ( sendMail ) sendResEmail(ss, to, menu, corpsHTML, nbSearchWithResultsSinceLastEmail, nbResultWithCriteriaSinceLastEmail);
                }
                bodyHTMLForThisSearch = ""
                corpsHTML = "";
                menu = "";
                nbSearchWithResultsSinceLastEmail = 0;
                nbResultWithCriteriaSinceLastEmail = 0;
                nbResultWithCriteriaSinceLastEmailForThisSearch = 0;
              }
              
              // prendre l'annonce suivante
              var idxOfAdStart = dataList.indexOf('<li>', idxOfAdEnd);
            } //while( idxOfAdStart > 0 )
          } // if ( rep.indexOf("Aucune annonce") < 0 )
          
          //var noMore = !rep.match("Page suivante")  ||  rep.match("<li class=\"page\">\s*Page suivante\s*</li>");
          //currentPage += 1;
          // dorénavant, on ne parcours que la première page de résultat.
          var noMore = true;
        }
        while (!noMore);
      }
      catch(e)
      {
        bodyHTMLForThisSearch = '';
        bodyHTMLForThisSearch += '<li style="list-style:none;margin-bottom:20px; clear:both;background:#EAEBF0;border-top:1px solid #ccc;">'
        {
          // Ouverture du div
          bodyHTMLForThisSearch += "<div style=\"float:left;width:200px;padding:20px 0;\">"; //;width:420px
          {
            bodyHTMLForThisSearch += '<div>Erreur ' + e + ' ligne:' + e.lineNumber +'</div>';
          }
          // Fermeture du div
          bodyHTMLForThisSearch += "</div>";
        }
        // Fermeture de <li>
        bodyHTMLForThisSearch += "</li>";
        nbResForThisSearch = 1;
        nbResultWithCriteriaForThisSearch = 1;
        nbResultWithCriteriaSinceLastEmailForThisSearch = 1;
        nbResultWithCriteriaSinceLastEmail++;
        nbResTotWithCriteria++;
        
      }
        
      if ( nbResultWithCriteriaSinceLastEmail > 0 ) {
        nbSearchWithResultsSinceLastEmail++;
        corpsHTML = corpsHTML + "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Votre recherche : <a name=\""+ searchName + "\" href=\""+ searchURL + "\"> "+ searchName +" (" + nbResultWithCriteriaSinceLastEmailForThisSearch + ")</a></p><ul>" + bodyHTMLForThisSearch + "</ul>";
        menu += "<li><a href=\"#"+ searchName + "\">"+ searchName +" (" + nbResultWithCriteriaSinceLastEmailForThisSearch + ")</a></li>"
      }
      
      sheetValues[rowSearchTitles+1-1][sheetObj.searchColLastExec-1] = new Date; // date dernière execution
      sheetValues[rowSearchTitles+1-1][sheetObj.searchColNbResults-1] = nbResForThisSearch;
      sheetValues[rowSearchTitles+1-1][sheetObj.searchColNbResultsMatch-1] = nbResultWithCriteriaForThisSearch;
      
      // Log
      if (ScriptProperties.getProperty('log') == "true" || ScriptProperties.getProperty('log') == null || ScriptProperties.getProperty('log') == "")
      {
        if ( nbResTot > 0 ) {
          slog.insertRowBefore(2);
          slog.getRange("A2").setValue(searchName);
          slog.getRange("B2").setValue(nbResForThisSearch);
          slog.getRange("C2").setValue(nbResultWithCriteriaForThisSearch);
          slog.getRange("D2").setValue(new Date);
        }
      }
      //sheet.getRange(1, 1, sheetValues.length, sheetValues[0].length).setValues(sheetValues);
    } // for ( sheetsIndex = 0 ; sheetsIndex < sheetObjArray.length ; sheetsIndex += 1 )
    
    //on envoie le mail?
    if ( sendMail ) sendResEmail(ss, to, menu, corpsHTML, nbSearchWithResultsSinceLastEmail, nbResultWithCriteriaSinceLastEmail);
    nbSearchWithResultsSinceLastEmail = 0;
    nbResultWithCriteriaSinceLastEmail = 0;

    // write modif
    for ( sheetsIndex = 0 ; sheetsIndex < sheetObjArray.length ; sheetsIndex += 1 )
    {
      var sheetObj = sheetObjArray[sheetsIndex];
      sheetObj.sheet.getRange(1, 1, sheetObj.values.length, sheetObj.values[0].length).setValues(sheetObjArray[sheetsIndex].values);
    }
  }
  catch (e)
  {
    if ( throwErrorByMail == true ) {
      var error = "Erreur :" + e;
      var dummy = 0;
      MailApp.sendEmail(to, "Alert LBC Exception", "line " + e.lineNumber + "  -  " + e);
    } else {
      throw e;
    }
  }
}

function sendResEmail(sheetObjArray, to, menu, corpsHTML, nbSearchWithRes, nbResTotWithCriteria)
{
    if (nbSearchWithRes > 1)
    {
      //plusieurs recherche, on créé un menu
      menu = "<p style=\"display:block;clear:both;padding-top:20px;font-size:14px;\">Accès rapide :</p><ul>" + menu + "</ul>";
      corpsHTML = menu + corpsHTML;
    }
    if( corpsHTML != "" )
    {
      var title = "Alerte leboncoin.fr : " + nbResTotWithCriteria + " nouveau" + (nbResTotWithCriteria>1?"x":"") + " résultat" + (nbResTotWithCriteria>1?"s":"");
      var corpsAsText = "Si cet email ne s’affiche pas correctement, veuillez sélectionner\nl’affichage HTML dans les paramètres de votre logiciel de messagerie.";
      corpsHTML = "<body>" + corpsHTML + "</body>";
      var l = corpsHTML.length;
      MailApp.sendEmail(to,title,corpsAsText,{ htmlBody: corpsHTML });
    }
}

function checkNow()
{
  lbc(true, true);
}

function setupMail()
{
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
      Browser.msgBox("Email " + ScriptProperties.getProperty('email') + " modifié");
    }
  }
}

/**
* Extrait l'id de l'annonce LBC
*/
function extractId_(id)
{
  var returnValue = id.substring(id.lastIndexOf("/") + 1,id.indexOf(".htm"));
  return returnValue;
}

/**
* Extrait le lien de l'annonce
*/
function extractHref_(data)
{
  var hrefTag = 'href="//';
  var endOfUrl = ".htm";
  var idxHref = data.indexOf(hrefTag); // ne fonctionnera plus si leboncoin ajoute un espace avant ou après le '='
  if ( idxHref == -1 ) {
    throw new Error("Attribut \"" + hrefTag + "\" non trouvé dans balise a");
  }
  var idxDothtm = data.indexOf(endOfUrl, idxHref);
  if ( idxHref == -1 ) {
    throw new Error("Fin d'URL (\"" + endOfUrl + "\") non trouvé dans balise a");
  }
  var res = data.substring(idxHref + hrefTag.length , idxDothtm + endOfUrl.length);
  return res;
}

/**
* Extrait le titre de l'annonce
*/
function extractTitle_(data)
{
  var idxStart = indexOfAfter(data, "title=\"");
  var idxEnd = data.indexOf("\"", idxStart);
  var returnValue = data.substring(idxStart , idxEnd);
  return returnValue;
}

/**
* Extrait vendeur pro
*/
function extractPro_(data)
{
  if(data.indexOf("(pro)") > 0){
    return " (pro)";
  }else{
    return "";
  }
}

/**
* Extrait le lieu de l'annonce
*/
function extractPlace_(data)
{
  var idx1 = indexOfAfter(data, '<p class="item_supp">');
  var idxStart = indexOfAfter(data, '<p class="item_supp">', idx1);
  var idxEnd = data.indexOf('</p>', idxStart);
  var returnValueBrut = data.substring(idxStart , idxEnd);
  var returnValue = returnValueBrut.trim();
  return returnValue;
}

/**
* Extrait le prix de l'annonce
*/
function extractPrice_(data)
{
  var idxStart = indexOfAfter(data, '<h3 class="item_price">');
  var idxEnd = data.indexOf('&nbsp;', idxStart);
  var returnValue = data.substring(idxStart, idxEnd);
  return returnValue;
}

/**
* Extrait la date de l'annonce.
* Return : un div contenant 2 div. Ex : <div> <div>Aujourd'hui</div> <div>16:47</div> </div>
*/
function extractHtmlDate_(data)
{
  var idx1 = indexOfAfter(data, '<p class="item_supp">');
  var idx2 = indexOfAfter(data, '<p class="item_supp">', idx1);
  var idxStart = indexOfAfter(data, '<p class="item_supp">', idx2);
  var idxUrgent = indexOfAfter(data, '<span class="item_supp emergency"><i class="icon-star"></i>Urgent</span>', idxStart);
  if ( idxUrgent > 0 ) {
    idxStart = idxUrgent;
  }
  var idxEnd = data.indexOf('</p>', idxStart);
  var returnValue = data.substring(idxStart , idxEnd).trim();
  return returnValue;
}

/**
* Extrait la date de l'annonce.
* Return : un div contenant 2 div. Ex : <div> <div>Aujourd'hui</div> <div>16:47</div> </div>
*/
function extractDate_(data)
{
  var dataText = extractHtmlDate_(data);
  var idxComma = dataText.indexOf(',');
  var dayAndMonth = dataText.substring(0, idxComma);
  var hourAndMinute = dataText.substring(idxComma+1).trim();
  
  var day;
  var month;
  if ( dayAndMonth == "Aujourd'hui" ) {
    var now = new Date();
    day = now.getDate(); // 1-31
    month = now.getMonth(); // 0-11
  }
  else
    if ( dayAndMonth == "Hier" ) {
      var now = new Date();
      now.setHours(-24);
      day = now.getDate(); // 1-31
      month = now.getMonth(); // 0-11
    }else
    {
      day = dayAndMonth.substring(0, dayAndMonth.indexOf(" "));
      month = dayAndMonth.substring(dayAndMonth.indexOf(" ")+1);
      if ( month == "jan" ) month = 0;
      else if ( month == "fév" ) month = 0;
      else if ( month == "mars" ) month = 1;
      else if ( month == "avr" ) month = 2;
      else if ( month == "mai" ) month = 3;
      else if ( month == "juin" ) month = 4;
      else if ( month == "juil" ) month = 5;
      else if ( month == "août" ) month = 6;
      else if ( month == "sept" ) month = 7;
      else if ( month == "oct" ) month = 8;
      else if ( month == "nov" ) month = 9;
      else if ( month == "déc" ) month = 10;
      else throw "Erreur décodage mois. " + dayAndMonth;
    }
  var hour = hourAndMinute.substring(0, hourAndMinute.indexOf(":"));
  var minute = hourAndMinute.substring(hourAndMinute.indexOf(":")+1);
  
  var now = new Date();
  var year = now.getFullYear();
  
  var adDate = new Date(now.getFullYear(), month, day, hour, minute, 0);
  //  var formattedDate = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
  return adDate;
}

/**
* Extrait l'image de l'annonce
*/
function extractImage_(data)
{
  if ( data.indexOf('no-picture.png') > 0 )
  {
    return "static.leboncoin.fr/img/no-picture.png";
  }
  var idx1 = indexOfAfter(data, '<div class="item_image">');
  var idxStart = indexOfAfter(data, 'imgSrc="//', idx1)
  var idxEnd = data.indexOf('"', idxStart);
  var returnValue = data.substring(idxStart, idxEnd);
  return returnValue;
}

/**
* Extrait la liste des annonces
*/
function splitResult_(text)
{
  var idxStart = indexOfAfter(text, '<ul class="tabsContent dontSwitch block-white">');
  var idxEnd = text.indexOf('</ul>', idxStart);
  return text.substring(idxStart, idxEnd);
}

//Activer/Désactiver les logs
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


function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name : menuMailSetupLabel,
      functionName : "setupMail"
    },
    null,
    {
      name : menuSearchLabel,
      functionName : "checkNow"
    },
    null,
    {
      name : menuLog,
      functionName : "dolog"
    },
    {
      name : menuArchiveLog,
      functionName : "archivelog"
    }
  ];
  sheet.addMenu(menuLabel, entries);
}

function onInstall()
{
  onOpen();
}
