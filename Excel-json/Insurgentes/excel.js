//https://levelup.gitconnected.com/how-to-convert-excel-file-into-json-object-by-using-javascript-9e95532d47c5
//convertir en EXCEL https://codepedia.info/javascript-export-html-table-data-to-excel
let selectedFile;
let selectedFile2;
console.log(window.XLSX);
document.getElementById('inputA').addEventListener("change", (event) => { //Lee estado de cuenta
    selectedFile = event.target.files[0];
})
document.getElementById('inputSicas').addEventListener("change", (event) => {// Lee SICAS
    selectedFile2 = event.target.files[0];
}
)

let objetoSICAS; //Array de objetos en el que se va a guarar SICAS
let objetoAserta; //Array de objetos en el que se va a guardar Berkley

document.getElementById('button').addEventListener("click", () => {
    var num, num2;
    if(selectedFile){ //Función para convertir Edo de Cuenta en array de objetos
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);
        fileReader.onload = (event)=>{
         let data1 = event.target.result;
         let workbook1 = XLSX.read(data1,{type:"binary"});
         console.log(workbook1);        
         workbook1.SheetNames.forEach(sheet => {
            //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                objetoAserta = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet], {range:1}); //Nombre del array
             console.log("ASERTA")
                console.log(objetoAserta);
         });

         if(selectedFile2){ //Función que convierte SICAS en array de objetos
            let fileReader = new FileReader();
            fileReader.readAsBinaryString(selectedFile2);
            fileReader.onload = (event)=>{
             let data2 = event.target.result;
             let workbook2 = XLSX.read(data2,{type:"binary"});
             workbook2.SheetNames.forEach(sheet => {
                  objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //Nombre del array

                    //La tabla tiene atributo HIDDEN para que no se vea, pero ahí está.
                  let tabla ="<table id='Aserta' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Total Comisión</th><th>Diferencia comisión</th></tr>";
                  let resultObject;
                  let sicas;
                  var encontrar;           

                  //Encontrar un valor ahí adentro
                  search = (key, ArreyAserta) => {
                    
                      for (let i=0; i < ArreyAserta.length; i++) {
                        var polizaAserta = String(ArreyAserta[i]["No Fianza/"])
                          if (polizaAserta == key) {
                            encontrar++;
                            if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                console.log("La diferencia es de"+diferencia);
                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td></tr>";
                            }
                          }

                      }
                      if(encontrar==0){
                        tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Importe"]+"</td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                        }
                        encontrar=0;
                    }
                    for(var j=0; j<objetoSICAS.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                        var poliza =  String (objetoSICAS[j].Poliza)
                        sicas=objetoSICAS[j];
                      resultObject = search(poliza, objetoAserta)
                      console.log("Número de registros en sicas: "+j);
                     document.getElementById("jsondata").innerHTML = tabla+"</table>"; //Se manda la tabla pero no se va a ver porque tiene HIDDEN
                    }
                    ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente.
                    if(resultObject==0){
                        document.getElementById("jsondata").innerHTML = "No se encontró ninguna fianza";

                    }
                            }
                );
             
            }
        } else{
             document.getElementById("jsondata").innerHTML = "No se adjuntó nada en SICAS";
        }
       
        }
    }else{
        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Berkley";
    }
    
});

function ExportToExcel(type, fn, dl) {// función que convierte a excel
    var elt = document.getElementById('Aserta');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    return dl ?
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || ('Aseguradora Aserta.' + (type || 'xlsx')));
 }