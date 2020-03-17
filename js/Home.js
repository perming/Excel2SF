////////////////////////////////////////////

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

/////////////////////////////////////////////////////////
//Effacement des cellules coloriées
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
/////////////////////////////////////////////
//Convertion ligne colonne en Adresse Range
function RangeTrait(NbLigne,NumCol){
  var TextCol = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","R","S","T","U","V","W","X","Y","Z"];
  var Colonne = TextCol[NumCol];
  var RangeT = Colonne + "1:" + Colonne + NbLigne;
  return RangeT;
}

/////////////////////////////////////////////////

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

 /////////////////////////
 
 function ErrorV(i) {console.log("Erreur ligne :" + ( i+ 1));
                                //  NbError ++;                           
                                //  range.format.fill.color = 'red';
 }              

////////////////////

function Ecriture_Range() {
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

////////////////////////////////////////////////
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

///////////////////////////////
function LoadJs(){
    console.log("Load");
    IdL=document.getElementById("ListeContrat").options.selectedIndex
    var Lc= localStorage.getItem("Liste_Contrat");
    var Lc =JSON.parse(Lc);
    var JsonFile=Lc.objects[IdL].json

  $.get(JsonFile, function(data){
    localStorage.setItem("JsonFile",JSON.stringify(data));
  });
    $.get(JsonFile, function(data){
        for (i=0; i<4;i++){
        var c1=data.Onglets[i].Titre
        var L1 = document.getElementById("Lbl" + i);
        L1.textContent=c1;
        }
    
    });    
}
/////////////////////////////

function NewSheet(){
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");
    return context.sync()
        .then(function () {
            for (i = 0 ; i < 4 ; i++){ 
                var C1 = document.getElementById("idCheck" + i).checked
                var L1 = document.getElementById("Lbl" + i);
                var L1 = L1.textContent;
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
        //  console.log("Ligne count : " + Nbli);
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

////////////////////////////////////

function Start1(){
      console.log("Start FHA ...")
      $.get("./json/00_list_type_import.json", function(data){
        localStorage.setItem("Liste_Contrat",JSON.stringify(data));
      });
      var Liste_Contrat= localStorage.getItem("Liste_Contrat");
      var Liste_Contrat =JSON.parse(Liste_Contrat);
      var ListeC = document.getElementById('ListeContrat');

      for (var i = 0 ; i < Liste_Contrat.objects.length ; i++) {
        var Contrat = Liste_Contrat.objects[i].name;
        var opt = document.createElement('option');
        opt.appendChild(document.createTextNode(Contrat) );
        opt.value = 'option value'; 
        ListeC.appendChild(opt); 
      }
      LoadJs();
      $.get("./json/Contrat_FM.json", function(data){
        localStorage.setItem("JsonFile",JSON.stringify(data));
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
validate.extend(validate.validators.datetime, {
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
});

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

function CompareCol(){

  console.log("RechercheV ...");
//setTimeout(function(){ 

 //Nbli=localStorage.getItem("Nbli");
  var Nbli = 50 //Nblx();
  var C = 0;
  Excel.run(function (context) {
    var RangeT = ("A1:A" + Nbli);
    var _range = context.workbook.worksheets.getActiveWorksheet().getRange(RangeT).load("values,address");
  return context.sync()
      .then(function () {
        console.log(Nbli);
        for (var i = 0 ; i<=Nbli ; i++ ) {
          var Cellv=_range.values[i][C];
          RechV(Cellv,i);

        }
        if (i==Nbli){console.log("Finish");}  
      });

  })
//}, 500);
}

////////////////////////////////////////
// Fonction RechercheV + coloration des cellules non trouvées
function RechV(Cellv,i){
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