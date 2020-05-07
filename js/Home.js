var arrSheetTraitement = [];

function init(){
      console.log("Start load json files ...");
	  document.getElementById("gif_patenter").style.display = "none";
      var data = getUrlJsonSync("./json/00_list_type_import.json");
	  localStorage.setItem("Liste_Contrat",JSON.stringify(data));
      var Liste_Contrat= localStorage.getItem("Liste_Contrat");
      var Liste_Contrat =JSON.parse(Liste_Contrat);
      var ListeC = document.getElementById('ListeContrat');

	  // Chargement de la liste des fichiers JSON contenu dans 00_list_type_import.json
      for (var i = 0 ; i < Liste_Contrat.objects.length ; i++) {
        var Contrat = Liste_Contrat.objects[i].name;
        var opt = document.createElement('option');
        opt.appendChild(document.createTextNode(Contrat) );
        opt.value = 'option value'; 
        ListeC.appendChild(opt); 
      }
      onChangeJSON();
	  //
      var data1 = getUrlJsonSync("./json/Contrat_LP.json");
      localStorage.setItem("JsonFile",JSON.stringify(data1));
}
///////////////////////////////
function onChangeJSON(){
    console.log("Load JSON file");
	var strDivId = "accordionId";
    IdL=document.getElementById("ListeContrat").options.selectedIndex
    var Lc= localStorage.getItem("Liste_Contrat");
    var Lc =JSON.parse(Lc);
    var JsonFile=Lc.objects[IdL].json
	var data = getUrlJsonSync(JsonFile);
	localStorage.setItem("JsonFile",JSON.stringify(data));
	clearTableOnglet(strDivId);
	addHtml(strDivId,"Description",data.Description);
	//data.Colonnes[j].Formule
	var arrayListOnglet = new Array();
    for (i=0; i<data.Onglets.length;i++){
		var c1=data.Onglets[i].Titre
		var strTxt = data.Onglets[i].Description + "<BR>"
		if(!(data.Onglets[i].Colonnes===undefined)){
			for (j=0; j<data.Onglets[i].Colonnes.length;j++){
				strTxt = strTxt + "<li>" + data.Onglets[i].Colonnes[j].Nom +" : " + data.Onglets[i].Colonnes[j].Aide +"</li>" ;
			}
		}
		addHtml(strDivId,c1,strTxt);
		arrayListOnglet.push(c1);
		//insertRowOnglet(c1);
		//var L1 = document.getElementById("Lbl" + i);
		//L1.textContent=c1;
    }
	localStorage.setItem("arrayListOnglet",JSON.stringify(arrayListOnglet));
	localStorage.setItem("arrayListOngletSelected",JSON.stringify(arrayListOnglet));
}

function addHtml(strDivId,strTitre,strText){
	var html = document.getElementById(strDivId);
	var start = "<a style=\"height:35px;width:100%;text-align: left;\" class=\"btn btn-primary\" data-toggle=\"collapse\" href=\"#multiCollapse"+strTitre+"\" role=\"button\" aria-expanded=\"true\" aria-controls=\"multiCollapse"+strTitre+"\">";
	var milde = "</a><div class=\"collapse multi-collapse\" id=\"multiCollapse"+strTitre+"\"><div class=\"card card-body\">";
	var end = "</div></div>";
	var addRow = start+strTitre+milde+strText+end;
	html.innerHTML = html.innerHTML+addRow;
}
function insertRowOnglet(title){
	var Table = document.getElementById("listeOnglet");
	var addRow = "<tr><th scope=\"row\"><input onchange=\"onChangeCheked(this);\" class=\"form-check-input\" type=\"checkbox\" value=\""+title+"\" checked=\"true\" id=\"idCheck"+title+"\"></th><td><label id=\"Lb"+title+"\">"+title+"</label></td></tr>";
	Table.innerHTML = Table.innerHTML+addRow;
}
function clearTableOnglet(strID){
	var html = document.getElementById(strID);
	html.innerHTML = "";
}
function onChangeCheked(checkbox){
	var arrayListOnglet = JSON.parse(localStorage.getItem("arrayListOngletSelected"));
	if(checkbox.checked){
		arrayListOnglet.push(checkbox.value);
	}else{
		arrayListOnglet.splice(arrayListOnglet.indexOf(checkbox.value), 1);
	}
	localStorage.setItem("arrayListOngletSelected",JSON.stringify(arrayListOnglet));
}
/////////////////////////////

function getExcelColonneStr(strHeader,strAide,strValue,cptRow) {
	var tableau = [];
	var tHeader = [];
	var tAide = [];
	var tValue = [];
	tHeader.push(strHeader);
	tAide.push(strAide);
	tValue.push(strValue);
	tableau.push(tHeader);
	tableau.push(tAide);
	for(var k = 1; k < cptRow; k++){
		tableau.push(tValue);
	}
    return tableau;
}
function getExcelColonneFormuleStr(strValue,cptRow) {
	var tableau = [];
	var tValue = [];
	tValue.push(strValue);
	for(var k = 0; k < cptRow; k++){
		tableau.push(tValue);
	}
	console.log("tableauFormule = " + tableau);
	return tableau;
}
 function getUrlJsonSync(url){
    var jqxhr = $.ajax({
        type: "GET",
        url: url,
        dataType: 'json',
        cache: false,
        async: false
    });

    // 'async' has to be 'false' for this to work
    var response = jqxhr.responseJSON;
    return response;
} 

function loadValue(){
	document.getElementById("gif_patenter").style.display = "block";
	arrSheetTraitement = JSON.parse(localStorage.getItem("arrayListOngletSelected"));
	loadFile(arrSheetTraitement.shift());
}
/** 
 * @description
 * @param
 * @return
 */
function fusionRefFile(){
	var myFile = document.getElementById("file");
	var reader = new FileReader();

	reader.onload = (event) => {
		Excel.run((context) => {
			// strip off the metadata before the base64-encoded string
			var startIndex = reader.result.toString().indexOf("base64,");
			var workbookContents = reader.result.toString().substr(startIndex + 7);

			var sheets = context.workbook.worksheets;
			sheets.addFromBase64(
				workbookContents,
				null, // get all the worksheets
				Excel.WorksheetPositionType.after, // insert them after the worksheet specified by the next parameter
				sheets.getActiveWorksheet() // insert them after the active worksheet
			);
			return context.sync();
		});
	};

	// read in the file as a data URL so we can parse the base64-encoded string
	reader.readAsDataURL(myFile.files[0]);
}
/** 
 * @description
 * @param
 * @return
 */
function loadFile(sheetName){
	console.log("Start loadFile");
	var arrayListOnglet = JSON.parse(localStorage.getItem("arrayListOnglet"));
	Excel.run(function (context) {
		var arrayListOnglet = JSON.parse(localStorage.getItem("arrayListOngletSelected"));
		var sheets = context.workbook.worksheets;
		var sheet;
		var table;
		var tables = context.workbook.tables;
		var headerTable;
		var bodyTable;
		sheets.load("items/name");
		return context.sync()
			//***** Vérification de l'existance des onglets
			.then( function () {
				for (i = 0 ; i < arrayListOnglet.length ; i++){ 
					var C1 = true;//document.getElementById("idCheck" + arrayListOnglet[i]).checked
					var L1 = arrayListOnglet[i];
					if (C1 ==true){
						var Find= false;
						for (var j in sheets.items) {
							AddSheet=sheets.items[j].name;
							if (AddSheet === L1 ){Find=true};                     
							}
						if (Find==false){ 
							console.log(L1 + " créé");
							var varSheet = sheets.add(L1);
							varSheet.load("name, position");
						}
					}
				}
				sheets.load("items/name");
				tables.load("items/name");
				return context.sync();
			})
			//***** Création des tableaux
			.then(function () {
				var tablesName =[];
				for(var i = 0; i < tables.items.length; i++) 
                { 
					tablesName.push(tables.items[i].name);
                }
				sheets.items.forEach( function (varSheet) {
					if(varSheet.name==sheetName){
						sheet=varSheet;
					}
				});
				sheet.activate();
				if(!(tablesName.indexOf(sheet.name)>-1)){
					table = sheet.tables.add(sheet.getUsedRange(), true);
					table.name = sheet.name;
				}else{
					table = sheet.tables.getItem(sheet.name);
				}
				headerTable = table.getHeaderRowRange().load("values");
				bodyTable = table.getDataBodyRange().load("values");
				sheets = context.workbook.worksheets;
				sheets.load("items/name");
				table.columns.load("items/name");
				return context.sync(table);
			})
			//***** Création des entête de colonne
			.then(function (table){
				var jsonFile =JSON.parse(localStorage.getItem("JsonFile"));
				localStorage.setItem(sheetName+"_tableHeader",JSON.stringify(headerTable.values));
				localStorage.setItem(sheetName+"_tableValue",JSON.stringify(bodyTable.values));
				var headerTabCount = headerTable.values[0].length;
				var valueTabCount = bodyTable.values.length;
				console.log("headerTabCount = "+headerTabCount);
				console.log("valueTabCount = "+valueTabCount);
				var onglet = jsonFile.Onglets.find(Onglets => {
					   return Onglets.Titre == sheet.name
				})
				
					var tableau = [];
					var colLength ;
					var colHeader;
					var colAide;
					var colVal;
					var colForm;
					var keys;
					var jsondat;
					if(onglet.URLJSONData===undefined){
						colLength = onglet.Colonnes.length;
					}else{
						jsondata = getUrlJsonSync(onglet.URLJSONData);
						keys = Object.keys(jsondata.Data[0]);
						colLength =keys.length;
					}
					for(var j = 0; j < colLength; j++){
						if(onglet.URLJSONData===undefined){
							colHeader = onglet.Colonnes[j].Nom;
							colAide = "";//onglet.Colonnes[j].Aide;
							if(!(onglet.Colonnes[j].Value===undefined)){
								colVal = onglet.Colonnes[j].Value;
								colAide = colVal;
							}							
						}else{
							colHeader = keys[j];
							colAide = "";
							colVal = "";
						}
						var col;
						if(!(headerTable.values[0].indexOf(colHeader)>-1)){
							col = table.columns.add(null,getExcelColonneStr(colHeader,colAide,colVal,valueTabCount));
						}else{
							col = table.columns.getItem(colHeader);
						}
						if(!(onglet.Colonnes===undefined) && !(onglet.Colonnes[j].Formule===undefined)){
							col.getDataBodyRange().formulasLocal = getExcelColonneFormuleStr(onglet.Colonnes[j].Formule,valueTabCount);
						}
						tableau = [];
					}
					//***** Suppression de la colonne "Colonne1" créé par défaut si l'onglet était vide // ATTENTION: Si Excel est dans une autre langue cela ne fonctionne plus 
					if(valueTabCount==1 && table.columns.items[0].name=="Colonne1"){
						var column = table.columns.items[0];
						column.delete();
					}
				return context.sync();
			})
			.then(function () {
				if(arrSheetTraitement.length>0){
					loadFile(arrSheetTraitement.shift());
				}else{
					document.getElementById("gif_patenter").style.display = "none";
				}
			})
			;
	})//.catch(errorHandlerFunction);
}

/** 
 * @description
 * @param
 * @return
 */
function Start(){
  console.log("Start");

  var Nbli = Nblx();
  var Colj = Coly();

  setTimeout(function(){ 
    
    Nbli=localStorage.getItem("Nbli");
    Colj=localStorage.getItem("Colj");
    console.log(" Test Nbl : " + Nbli);
    console.log(" Test Nbc : " + Colj);
    
  //  Nbli = 22;
    Colj= 5;

      // for (Colj=0; Colj < 3; Colj++){
          Clear(Nbli,Colj);
          Verif(Nbli,Colj);
      // }  
      console.log("This is the end"); 
  }, 2000);
 
}

/** 
 * @description Effacement des cellules coloriées
 * @param
 * @return
 */
function Clear(Nbli,Colj){
    console.log("Start Clear");
    var RangeT = RangeTrait(Nbli,Colj);

    Excel.run(function (context) {

      console.log("Clear : " + RangeT);
        var Range2 = context.workbook.worksheets.getActiveWorksheet().getRange(RangeT).load("values,address"); 
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        // context.workbook.comments.getItemByCell("Data!A2").delete();
        //Range.comments.delete();
        return context.sync().then(function () { 
          console.log(RangeT);
          Range2.format.fill.color= "#FFFFFF" //"white";
          console.log("Clear fini"); 
        }); 
     })//.catch(function (error) { 
       // console.log(error); 
     //});
}

/** 
 * @description Convertion ligne colonne en Adresse Range
 * @param
 * @return
 */
function RangeTrait(NbLigne,NumCol){
  var TextCol = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","R","S","T","U","V","W","X","Y","Z"];
  var Colonne = TextCol[NumCol];
  var RangeT = Colonne + "1:" + Colonne + NbLigne;
  return RangeT;
}

/** 
 * @description Click sur le bouton vérifier
 * @param
 * @return
 */
function Verif(Nbli,Colj){
  Excel.run(function (context) {
    console.log("Start vérif");

    var NbError = 0;
    var RangeT = RangeTrait(Nbli,Colj);
    var _range = context.workbook.worksheets.getActiveWorksheet().getRange(RangeT).load("values,address");

      return context.sync().then(function () { 

        GetJson(Colj);
        console.log("Nom Colo : " + Name_Colonne);
        var rangeH = context.workbook.worksheets.getActiveWorksheet().getCell(0,Colj); 
        var Colo_ex =_range.values[0][0];
        console.log("Colo Ex : " + Colo_ex);
         if (Name_Colonne !== Colo_ex){
          rangeH.format.fill.color = 'yellow';
         }

          for (var i=1; i<Nbli;i++){
            var Cellv=_range.values[i][0];
            var CellC = ("Prénom" + ":" + Cellv);
            var range = context.workbook.worksheets.getActiveWorksheet().getCell(i,Colj);
            //var Result =validate.validators.datetime(Cellv, {datetime: true});  
              var Result = validate.single(Cellv, Constraints);
            if (typeof Result !== 'undefined'){
                  //  console.log("Erreur ligne :" + ( i+ 1)); 
                     range.format.fill.color = 'red';
                     NbError ++; 
                     // var comments = context.workbook.comments;
                     // comments.add(range, "Erreur 99");
                    var Error1 = Result[0];
                   // console.log(Error1);
              }
          } 
        console.log("Vérif Terminé : " + NbError + " Erreurs" + " - Colonne : " + Colj);
        //MsgBox(); 
      }); 
   })//.catch(function (error) { 
     // console.log(error); 
   //});
  } 

/** 
 * @description 
 * @param
 * @return
 */ 
 function ErrorV(i) {console.log("Erreur ligne :" + ( i+ 1));
                                //  NbError ++;                           
                                //  range.format.fill.color = 'red';
 }              

/** 
 * @description 
 * @param
 * @return
 */
function Ecriture_Range() {
	console.log("Start method : Ecriture_Range");
    Excel.run(function (context) {
        var sheetName = 'Data';
        var rangeAddress = 'A1:A2000';
        var worksheet = context.workbook.worksheets.getItem(sheetName);
    
        var range = worksheet.getRange(rangeAddress);
        range.numberFormat = 'm/d/yyyy';
        range.values = '3/11/2015';
        range.load('text');
    
        return context.sync()
          .then(function () {
            console.log(range.text);
        });
    }).catch(function (error) {
        console.log('Error: ' + error);
        if (error instanceof OfficeExtension.Error) {
          console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
} 

/** 
 * @description 
 * @param
 * @return
 */
function Affiche_le_Range_Sélectionné() {
    Excel.run(function (context) {
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.load('address');
        return context.sync()
          .then(function () {
            console.log('The selected range is: ' + selectedRange.address);
        });
    }).catch(function (error) {
        console.log('error: ' + error);
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
} 
//////////////////////////////////
// Ouvrir une fenetre dialogue HTML
function MsgBox() {
    console.log("Hello");
    // document.write('Hello World!');
    Office.context.ui.displayDialogAsync('https://localhost/Test/Exemple1_validate.html', {height: 30, width: 20,  displayInIframe: true});
    // app.showNotification ("titre", "Hello");
    }    


function NewSheet(){
	Excel.run(function (context) {
		var arrayListOnglet = JSON.parse(localStorage.getItem("arrayListOngletSelected"));
		var sheets = context.workbook.worksheets;
		sheets.load("items/name");
		return context.sync()
			.then( function () {
				for (i = 0 ; i < arrayListOnglet.length ; i++){ 
					var C1 = document.getElementById("idCheck" + arrayListOnglet[i]).checked
					//var L1 = document.getElementById("Lbl" + arrayListOnglet[i]);
					//var L1 = L1.textContent;
					var L1 = arrayListOnglet[i];
					if (C1 ==true){
						var Find= false;
						for (var j in sheets.items) {
							AddSheet=sheets.items[j].name;
							if (AddSheet === L1 ){Find=true};                     
							}
						if (Find==false){ 
							console.log(L1 + " créé");
							var sheet = sheets.add(L1);
							sheet.load("name, position");
						}
					}
				}
			//    console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
			});
	})//.catch(errorHandlerFunction);
}

///////////////////////////////////////////////
function Nblx(){
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getUsedRange();
    range.load("rowCount");
    return context.sync()
        .then(function () {
          var Nbli = range.rowCount;
          localStorage.setItem("Nbli",Nbli);
        });
  })//.catch(errorHandlerFunction);
}
//////////////////////////////////////////////
function Coly(){
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getUsedRange();
    range.load("columnCount");
    return context.sync()
        .then(function () {
          var Colj = range.columnCount;
        //  console.log("Col count : " + Colj);
          localStorage.setItem("Colj",Colj)
        });
  })//.catch(errorHandlerFunction);
}

//////////////////////////////////////////////////////
//Selection Range
function SelectRange (Range){
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange(Range).select();
    return context.sync()
    //Range
        .then(function () {
        });
  })
}

////////////////////////////////////////////
// Insertion Fichier
function InsertFile(){
  SelectRange("A1");

  setTimeout(function(){ 
    var ImportF = [["Hello","coucou"],["Hello1","coucou1"]];
    console.log(ImportF);  
    Office.context.document.setSelectedDataAsync(ImportF,
                function (asyncResult) {
                    var error = asyncResult.error;
                    if (asyncResult.status === "failed") {
                    }
                    else {
                        console.log("ok");
                    }
                });
      },1000);     
        }


////////////////////////////
function run1() {

    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error.message);
      } else {
        console.log(`The selected data is "${asyncResult.value}".`);
      }
    });
  }



///////////////////////////////////////////////////////  
//Effacement Commentaire
function DelCom(d,c){
     Excel.run(function (context) {
       var Comm = (d+ "!A" + c);
      //console.log(Comm)
       return context.sync();
       context.workbook.comments.getItemByCell(Comm).delete();
   
     });
}

//////////////////////////////////////////////////////////////////

  function Verif_File(){
    console.log("Vérifivcation");
    var fileUpload = document.getElementById("fileConvert");
    
     var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.csv|.txt|.json)$/;
    if (regex.test(fileUpload.value.toLowerCase())) {
        if (typeof (FileReader) != "undefined") {
            var reader = new FileReader();
            reader.onload = function (e) {
                JsonFile=(e.target.result);
                console.log(JsonFile);

             let myNewJSON = JSON.parse(JsonFile);
            // Site=myNewJSON.Onglets[0]["Colonnes"][0].Nom
            var Contrat =[];
            for (i = 0 ; i < 3 ;i++){ 
             Contrat[i] = myNewJSON.Onglets[i]['Titre'];   
            }
            console.log(Contrat);
            }
            reader.readAsText(fileUpload.files[0]);
        } else {
            console.log("This browser does not support HTML5.");
        }
    } else {
        console.log("Please upload a valid CSV file.");
    }
}

///////////////////////////////////////////////

function GetJson(Colj){
  var JsonFile= localStorage.getItem("JsonFile");
  var JJ =JSON.parse(JsonFile);
  return [Name_Colonne = JJ.Onglets[0].Colonnes[Colj].Nom,
         Constraints = JJ.Onglets[0].Colonnes[Colj].Constraints
        ]
}

///////////////////////////////////////
// Désactivé avec le validatejs (Pour test)
/* validate.extend(validate.validators.datetime, {
  // The value is guaranteed not to be null or undefined but otherwise it
  // could be anything.
  parse: function(value, options) {
    return +moment.utc(value);
  },
  // Input is a unix timestamp
  format: function(value, options) {
    var format = options.dateOnly ? "DD/MM/YYYY" : "DD/MM/YYYY hh:mm:ss";
    return moment.utc(value).format(format);
  }
});*/

/////////////////////////////
function VerifOng(){
  console.log("Vérif");

  Excel.run(function (context) {
    var sheetName1 = 'Contrat1';
    var rangeAddress = 'A1';
    var worksheet = context.workbook.worksheets.getItem(sheetName1);

    var range = worksheet.getRange(rangeAddress).value;
    return context.sync()
    console.log(range);
  })
}

///////////////////////////////////////////
function SelectSheet(Sheetname){
Excel.run(function (context) {
  var sheet = context.workbook.worksheets.getItem(Sheetname);
  sheet.activate();
  sheet.load("name");

  return context.sync()
      .then(function () {
          console.log(`The active worksheet is "${sheet.name}"`);
      });
})
}
////////////

function CompareCol(){

  console.log("RechercheV ...");
  ListeRech();
  setTimeout(function(){ 
  
  //Nbli=localStorage.getItem("Nbli");
  var Nbli = 55; //Nblx();
  var C = 0;
// SelectSheet("C1");
  Excel.run(function (context) {

    var RangeT = ("A1:A" + Nbli);
    var sheet = context.workbook.worksheets.getItem("C1");
    sheet.activate();
    sheet.load("name");

    var _range = context.workbook.worksheets.getActiveWorksheet().getRange(RangeT).load("values,address");
   
    var RangeC = localStorage.getItem("RangeC");
    RangeC=(RangeC.split(","));

  return context.sync()
      .then(function () {
        for (var i = 0 ; i<=Nbli ; i++ ) {
          var Cellv=_range.values[i][C];
          if (RangeC.indexOf(Cellv) == "-1"){
            console.log("Ligne : " + (i+1) + " Cellv : " + Cellv + " Trouvé : " +  (RangeC.indexOf(Cellv) +1));
            var range = context.workbook.worksheets.getActiveWorksheet().getCell((i+1),(C+1));
            range.format.fill.color = 'red';
          }
        }
      });
  })
  }, 1000);
}

////////////////////////////////////////
function ListeRech(){
  var Nbli = 59 //Nblx();
  var C = 0;
  Sheetname="C2";
  SelectSheet(Sheetname);
  setTimeout(function(){ 
  Excel.run(function (context) {
     var RangeT2 = ("A1:A" + Nbli);
     var worksheet = context.workbook.worksheets.getItem(Sheetname);
     var range2 = context.workbook.worksheets.getActiveWorksheet().getRange(RangeT2).load("values,address");
  return context.sync()
      .then(function () {
          var RangeC= range2.values;
        localStorage.setItem("RangeC",RangeC);
      });
  })
}, 1000);
}

////////////////////////////////////////
// Fonction RechercheV + coloration des cellules non trouvées
function RechV_old(Cellv,i){
  Excel.run(function (context) {
    var Range = context.workbook.worksheets.getItem("C2").getRange("A1:A2000");
    var unitSoldInNov = context.workbook.functions.vlookup(Cellv, Range, 1, false);
    unitSoldInNov.load('value');

    return context.sync()
        .then(function () {
          if (unitSoldInNov.value == null){
             console.log('Non Trouvé  Ligne : ' + i + '  ' + Cellv + '  ' + unitSoldInNov.value);            
             var CC = context.workbook.worksheets.getActiveWorksheet().getCell(i,0);
             CC.format.fill.color = 'red';
           }
        });
  })
}

///////////////////////////////////////////////
function Comp_WorkB(){
    Excel.run(function (context) {
     var Wb= (context.Workbook.Open.FileName="Excel2SF.xlsx", ReadOnly=True);
      var sheet =Wd.workbook.worksheets.getItem("C3");
    var range = sheet.getUsedRange();
    range.load("rowCount");
    return context.sync()
        .then(function () {
          var Nbli = range.rowCount;
         console.log("Ligne count : " + Nbli);
        });
  })
}
///////////////////////////////////////////////
function Add_Id(){
  console.log("Start vérif Id");
  Nbli=localStorage.getItem("Nbli");
  Colj=localStorage.getItem("Colj");

  Nbli="10";
  var RangeT = ("A1:B" + Nbli);

  console.log(" Test Nbl : " + Nbli);
  Excel.run(function (context) {
	  var _range = context.workbook.worksheets.getActiveWorksheet().getRange(RangeT).load("values,address");
	  
	  return context.sync().then(function () { 
		for (var i = 1 ; i<=Nbli ; i++ ) {
			var Cellv=_range.values[i][0];
		   // _range[i][1].values="ID_1";
			console.log(Cellv);
		};
	  });
  });
}