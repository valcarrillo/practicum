//https://levelup.gitconnected.com/how-to-convert-excel-file-into-json-object-by-using-javascript-9e95532d47c5
//convertir en EXCEL https://codepedia.info/javascript-export-html-table-data-to-excel
let selectedFile=[];
let selectedFile2;
let numarchivos=0;
console.log(window.XLSX);
var docchubb= document.getElementById('inputA').addEventListener("change", (event) => { //Lee estado de cuenta
    const files = event.target.files;
    for (numarchivos=0; numarchivos < files.length; numarchivos++) {
        selectedFile[numarchivos] = event.target.files[numarchivos];
     }
})

document.getElementById('inputSicas').addEventListener("change", (event) => {// Lee SICAS
    selectedFile2 = event.target.files[0];
}
)

let objetoSICAS; //Array de objetos en el que se va a guarar SICAS
let objetoAserta=[]; //Array de objetos en el que se va a guardar CHUBB
var fechamax = new Date("2000-01-02"); // (YYYY-MM-DD)
var fechamin = new Date("2100-01-01"); // (YYYY-MM-DD)
const month = ["Nada","ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"];

document.getElementById('button').addEventListener("click", () => {
    var num;
    if(selectedFile){ //Función para convertir Edo de Cuenta en array de objetos
        for(i=0; i<numarchivos; i++){
            let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile[i]);
        fileReader.onload = (event)=>{
         let data1 = event.target.result;
         let workbook1 = XLSX.read(data1,{type:"binary"});       
         workbook1.SheetNames.forEach(sheet => {
        // EL RANGO ES LO GRANDE DEL ENCABEZADO                                         
        objeto = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet], {range:1});
                //objetoCHUBB=Object.assign(objetoCHUBB,objeto);
                //objeto={};
                for(var j=0; j<objeto.length; j++){
                    objetoAserta.push(objeto[j]);
                }
         });
        }
        }
        console.log("objeto Aserta:");
         console.log(objetoAserta);
         if(selectedFile2){ //Función que convierte SICAS en array de objetos
            let fileReader = new FileReader();
            fileReader.readAsBinaryString(selectedFile2);
            fileReader.onload = (event)=>{
             let data2 = event.target.result;
             let workbook2 = XLSX.read(data2,{type:"binary"});
             workbook2.SheetNames.forEach(sheet => {
                  objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //Nombre del array
                  console.log(objetoSICAS);
                   
            }
            );
             
            }
        } else{
             document.getElementById("jsondata").innerHTML = "No se adjuntó nada en SICAS";
        }
       
    }else{
        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de CHUBB";
    }
    
});
//fn: filename
function ExportToExcel(type, fn, dl) {// función que convierte a excel
    var elt = document.getElementById('CHUBBFianzas');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    var nombre ='CONCILIACIÓN CHUBB DEL '+fechamin.getDate()+" "+month[+fechamin.getMonth()+1]+" "+fechamin.getFullYear()+" AL "+fechamax.getDate()+" "+month[+fechamax.getMonth()+1]+" "+fechamax.getFullYear()+".";
    return dl ? //It will attempt to force a client-side download.
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || (nombre + (type || 'xlsx')));
}
