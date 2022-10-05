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
var messicas=0, aniosicas=0;
const month = ["Nada","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Augosto","Septiembre","Octubre","Noviembre","Diciembre"];

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
         });

         if(selectedFile2){ //Función que convierte SICAS en array de objetos
            let fileReader = new FileReader();
            fileReader.readAsBinaryString(selectedFile2);
            fileReader.onload = (event)=>{
             let data2 = event.target.result;
             let workbook2 = XLSX.read(data2,{type:"binary"});
             workbook2.SheetNames.forEach(sheet => {
                  objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //Nombre del array
                  console.log(objetoSICAS);
                    //La tabla tiene atributo HIDDEN para que no se vea, pero ahí está.
                  let tabla ="<table id='BerkleyFianzas' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Tipo Comisión</th><th>Total Comisión</th><th>Diferencia comisión</th><th>Diferencia en:</th></tr>";
                  let resultObject;
                  let sicas;
                  var encontrar;
                  let ArraydeEndosos;    
                  var err="NO SE ENCONTRÓ LA PÓLIZA";

                  //Encontrar un valor ahí adentro
                  search = (key, inclu, endo, ArrayBerkley) => {
                      for (let i=0; i < ArrayBerkley.length; i++) {
                        encontrar=0;
                          if (ArrayBerkley[i].FIANZA == key) {
                            if (ArrayBerkley[i].INCLUSION == inclu) {
                                if (ArrayBerkley[i].MOVIMIENTO == endo) {
                                //compara las primas netas y si son diferentes las mete en la tabla.
                                if(ArrayBerkley[i]["COMISIONES"] != sicas["Importe"]){
                                    encontrar=1;
                                    var diferencia= Math.abs(Math.round((ArrayBerkley[i]["COMISIONES"] -sicas["Importe"])*100)/100);
                                    var tipodif;
                                    if(ArrayBerkley[i]["PRIMA NETA"] !=sicas["PrimaNeta"] && ArrayBerkley[i]["% COMISION"] !=sicas["% Participacion"]){
                                        tipodif="Prima Neta y % Comisión";
                                    }else if(ArrayBerkley[i]["PRIMA NETA"] !=sicas["PrimaNeta"]){
                                        tipodif="Prima Neta";
                                    }else if(ArrayBerkley[i]["% COMISION"] !=sicas["% Participacion"]){
                                        tipodif="% Comisión";
                                    }else{
                                        tipodif=" ";
                                    }
                                    console.log("La diferencia es de"+diferencia);
                                    tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArrayBerkley[i].FIANZA+"-"+(ArrayBerkley[i].INCLUSION)+"-"+(endo-1)+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>"+tipodif+"</td></tr>"
                            }
                              return ArrayBerkley[i];
                                }else{
                                    err="NO SE ENCONTRÓ LA PÓLIZA";
                                }
                            }
                          }
                      }
                      if(encontrar==0){
                      tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"-"+inclu+"-"+(endo-1)+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                      }
                      encontrar=0; 
                    }
                    
                    if(typeof objetoSICAS[0].Poliza === 'undefined' || objetoBerkley[0].FIANZA==='undefined'){
                        document.getElementById("jsondata").innerHTML = "No se pudo leer el documento. Revise haber adjuntado el correcto.";
                    }else{
                        for(var j=0; j<objetoSICAS.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                            var fechas =objetoSICAS[j]["Fecha Pago Recibo"].split('/');
                            anio=fechas[2];
                            mes=fechas[1];
                            console.log("Mes: "+mes+" Año: "+anio);
                            var pol = objetoSICAS[j].Poliza.split('-');
                            if(anio>aniosicas){
                                aniosicas= +anio;
                                messicas= +mes;
                            }else if(anio==aniosicas){
                                if(mes>messicas){
                                        messicas= +mes;
                                    } 
                            }
                            var pol = objetoSICAS[j].Poliza.split('-'),
                            poliza = pol[2];
                            inclusion= pol[3];
                            if(typeof inclusion === 'undefined'){
                                inclusion=0;
                            }else{
                                inclusion= +inclusion;
                            }
                            num = +poliza;
                            var endo=objetoSICAS[j].Endoso;
                            if(typeof endo === 'undefined'){
                                endo= +1; //Si el endoso no está en SICAS entonces se registra como 1
                            }else{
                                endo= +endo +1;
                            }
                            sicas=objetoSICAS[j];
                        resultObject = search(num, inclusion, endo, objetoBerkley);
                        console.log(resultObject);
                        console.log("Número de registros en sicas: "+j);
                        document.getElementById("jsondata").innerHTML = tabla+"</table>   Mes: "+month[messicas]+" Año: "+aniosicas; //Se manda la tabla pero no se va a ver porque tiene HIDDEN
                        }
                    // ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente.
                        if(resultObject==0){
                            document.getElementById("jsondata").innerHTML = "No se encontró ninguna fianza";

                        }
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
    var elt = document.getElementById('BerkleyFianzas');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    return dl ?
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || ('Berkley_Fianzas_Comparacion_Mayo.' + (type || 'xlsx')));
 }