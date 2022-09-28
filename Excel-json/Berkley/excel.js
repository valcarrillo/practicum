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
let objetoBerkley; //Array de objetos en el que se va a guardar Berkley

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
                objetoBerkley = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet], {range:3}); //Nombre del array
             console.log(objetoBerkley);
            // document.getElementById("jsondata").innerHTML = JSON.stringify(objetoBerkley,undefined,4)
         });

         if(selectedFile2){ //Función que convierte SICAS en array de objetos
            let fileReader = new FileReader();
            fileReader.readAsBinaryString(selectedFile2);
            fileReader.onload = (event)=>{
             let data2 = event.target.result;
             let workbook2 = XLSX.read(data2,{type:"binary"});
             //console.log(workbook2);
             workbook2.SheetNames.forEach(sheet => {
                  objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //Nombre del array

                    //La tabla tiene atributo HIDDEN para que no se vea, pero ahí está.
                  let tabla ="<table id='BerkleyFianzas' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' hidden> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Total Comisión</th><th>Diferencia comisión</th></tr>";
                  let resultObject;
                  let sicas;
                  var encontrar;                 

                  //Encontrar un valor ahí adentro
                  search = (key, inclu, ArrayBerkley) => {
                      for (let i=0; i < ArrayBerkley.length; i++) {
                          if (ArrayBerkley[i].FIANZA == key) {
                            if (ArrayBerkley[i].INCLUSION == inclu) {
                            encontrar=1;
                            //compara las primas netas y si son diferentes las mete en la tabla.
                            if(ArrayBerkley[i]["PRIMA NETA"] != sicas["PrimaNeta"]){
                                var diferencia= Math.round((ArrayBerkley[i]["PRIMA NETA"] -sicas["PrimaNeta"])*100)/100;
                                console.log("La diferencia es de"+diferencia);
                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArrayBerkley[i].FIANZA+"-"+ArrayBerkley[i].INCLUSION+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td></tr>";
                            }
                              return ArrayBerkley[i];
                            }
                          }
                      }
                      if(encontrar==0){
                      tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"-"+inclu+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Importe"]+"</td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                      }
                      encontrar=0; 
                    }
                    for(var j=0; j<objetoSICAS.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                        var pol = objetoSICAS[j].Poliza.split('-'),
                        poliza = pol[2];
                        inclusion=pol[3];
                        if(typeof inclusion === 'undefined'){
                            inclusion=0;
                        }
                        num = +poliza;
                        sicas=objetoSICAS[j];
                      resultObject = search(num, inclusion, objetoBerkley);
                      console.log(resultObject);
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

        }
        // document.getElementById("jsondata2").innerHTML = "No se adjuntó nada";
        }
    }else{
        //document.getElementById("jsondata").innerHTML = "No se adjuntó nada";
    }
    
    
});

function ExportToExcel(type, fn, dl) {// función que convierte a excel
    var elt = document.getElementById('BerkleyFianzas');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    return dl ?
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || ('Berkley_Fianzas_Comparacion_Mayo.' + (type || 'xlsx')));
 }