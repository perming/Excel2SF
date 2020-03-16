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

    Clear(Nbli,0);

    Verif(Nbli,0);

     }, 2000);
}

/////////////////////////////////////////////////////////
//Effacement des cellules coloriées
function Clear(Nbli,Colj){
    console.log("Start Clear");
    var NbLigne=Nbli;
    var NumCol = 0;
    var RangeT = RangeTrait(NbLigne,NumCol);

    Excel.run(function (context) {

      console.log("Clear : " + RangeT);
        var Range2 = context.workbook.worksheets.getActiveWorksheet().getRange(RangeT).load("values,address"); 
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        // context.workbook.comments.getItemByCell("Data!A2").delete();
        //Range.comments.delete();
        return context.sync().then(function () { 
          console.log(RangeT);
          Range2.format.fill.color="white";
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

    var NbLigne= 30;// Nbli;
    var NumCol = 0;
    var NbError = 0;
    var RangeT = RangeTrait(NbLigne,NumCol);
    var _range = context.workbook.worksheets.getActiveWorksheet().getRange(RangeT).load("values,address");
    GetJson(Colj);
    console.log("Nom Colo : " + Name_Ong);
    console.log( Constraints);

      return context.sync().then(function () { 
        console.log(_range.values[0][NumCol]);
        
        for (var i=0; i<NbLigne;i++){
  
          var Cellv=_range.values[i][0];
          var range = context.workbook.worksheets.getActiveWorksheet().getCell(i,NumCol);  
          var Result = validate({password: Cellv}, Constraints);
          if (typeof Result !== 'undefined'){
                    console.log("Erreur ligne :" + ( i+ 1)); 
                     range.format.fill.color = 'red'
                     NbError ++; 
                     // var comments = context.workbook.comments;
                     // comments.add(range, "Erreur 99");
                     var Error1 = Result.password[0];
                     console.log(Error1);
          }
        } 
        console.log("Vérif Terminé : " + NbError + " Erreurs");
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
// Ouvrir une fenetre dialgue HTML
function MsgBox() {
    console.log("Hello");
    // document.write('Hello World!');
    Office.context.ui.displayDialogAsync('https://localhost/Test/Exemple1_validate.html', {height: 30, width: 20,  displayInIframe: true});
    // app.showNotification ("titre", "Hello");
    }    

///////////////////////////////
function LoadJs(){
    console.log("Load");
    $.get("Contrat.json", function(data){
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

////////////////////////////////////////////
// Reads data from current document selection and displays a notification
        function writeText() {
            Office.context.document.setSelectedDataAsync("Data here -->",
                function (asyncResult) {
                    var error = asyncResult.error;
                    if (asyncResult.status === "failed") {
                        //show error. Upcoming displayDialog API will help here.
                    }
                    else {
                        //show success.Upcoming displayDialog API will help here.
                    }
                });
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

      $.get("Contrat.json", function(data){
        //console.log(data);
        
        localStorage.setItem("JsonFile",JSON.stringify(data));
        });
       var JJ= localStorage.getItem("JsonFile");
       console.log(JSON.parse(JJ));
}


//////////////////////////////////////
function __Nblx(Nbli){
    return Nbly = 99;  
}

///////////////////////////////////////////////////////  
//Effacement Commentaire
function DelCom(d,c){
    // var sheet = context.workbook.worksheets.getActiveWorksheet();
     Excel.run(function (context) {
       var Comm = (d+ "!A" + c);
      //console.log(Comm)
       return context.sync();
       context.workbook.comments.getItemByCell(Comm).delete();
   
     });
}
//////////////////////////////////////////////////////////

  //     async function ff() {
  //       let promise = new Promise((resolve, reject) => {
  //         setTimeout(() => resolve(Nbl()), 5000)
  //       });
  //       let result = await promise; // wait until the promise resolves (*)
  //     console.log(result);
  //      // alert(result); // "done!"
  //     }
  // ff();


  ///////////////////////////////////////////////////////////

  function Verif_File(){
    console.log("Vérifivcation");
    var fileUpload = document.getElementById("fileConvert");
    
   //var fileUpload.value="C:\\Users\\f.hamid\\Desktop\\Contrat Simple.json";

    //C:\Users\f.hamid\Desktop\Contrat Simple.json
    //console.log(fileUpload);
    //  if (document.getElementById("fileConvert").value =""){
    //      console.log ("KO");
    //  };
    
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

                //console.log(JSON.stringify(result));
               // return JSON.stringify(result); //JSON
              
                
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
  return [Name_Ong = JJ.Onglets[0].Colonnes[Colj].Nom,
         Constraints = JJ.Onglets[0].Colonnes[Colj].Constraints
]

}