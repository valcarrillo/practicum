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
var fechamax = new Date("2000-01-02"); // (YYYY-MM-DD)
var fechamin = new Date("2100-01-01"); // (YYYY-MM-DD)
const month = ["Nada","ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"];

document.getElementById('button').addEventListener("click", () => {
    var num;
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
                  let tabla ="<table id='BerkleyFianzas' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' HIDDENa> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Tipo Comisión</th><th>Total Comisión</th><th>Diferencia de Comisión</th><th>No coincide:</th></tr>";
                  let resultObject;
                  let sicas;
                  var encontrar;
                  let tablanoencontrados="";    //Hay dos tablas, la de error y no encontrados. Se unen al final para tener más orden.
                  var err="NO SE ENCONTRÓ LA PÓLIZA";

                  //Encontrar un valor ahí adentro
                  search = (key, inclu, endo, ArrayBerkley) => {
                      for (let i=0; i < ArrayBerkley.length; i++) {
                        encontrar=0;
                          if (ArrayBerkley[i].FIANZA == key) {
                            if (ArrayBerkley[i].INCLUSION == inclu) {
                                if (ArrayBerkley[i].MOVIMIENTO == endo) {
                                    encontrar=1;
                                //compara las primas netas y si son diferentes las mete en la tabla.
                                    if(ArrayBerkley[i]["COMISIONES"] != sicas["Importe"]){
                                    var diferencia= Math.abs(Math.round((ArrayBerkley[i]["COMISIONES"] -sicas["Importe"])*10000)/10000);
                                    var tipodif;
                                    if(ArrayBerkley[i]["PRIMA NETA"] !=sicas["PrimaNeta"] && ArrayBerkley[i]["% COMISION"] !=sicas["% Participacion"]){
                                        tipodif="Prima Neta y % Comisión";
                                    }else if(ArrayBerkley[i]["PRIMA NETA"] !=sicas["PrimaNeta"]){
                                        tipodif="Prima Neta";
                                    }else if(ArrayBerkley[i]["% COMISION"] !=sicas["% Participacion"]){
                                        tipodif="% Comisión";
                                    }else{
                                        tipodif="Total Comisión";
                                    }
                                    console.log("La diferencia es de"+diferencia);
                                    tabla=tabla+"<tr><td style='background-color:#8495cb'>"+ArrayBerkley[i].FIANZA+"-"+(ArrayBerkley[i].INCLUSION)+"-"+(endo-1)+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:#9c0b0be7'>"+diferencia+"</td><td>"+tipodif+"</td></tr>"
                                }
                              return ArrayBerkley[i];
                                }else{
                                    err="NO SE ENCONTRÓ LA PÓLIZA";
                                }
                            }
                          }
                      }
                      if(encontrar==0){ // Encontrar es una bandera. Si no se encuentra, se incluye lo de abajo
                      tablanoencontrados= tablanoencontrados+"<tr><td style='background-color:#8495cb'>"+key+"-"+inclu+"-"+(endo-1)+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td></td><td style='background-color:#ff00007e'>NO SE ENCONTRÓ</td></tr>";
                        return ArrayBerkley;
                        }
                      encontrar=0; 
                    }
                    
                    if(typeof objetoSICAS[0].Poliza === 'undefined' || objetoBerkley[0].FIANZA==='undefined'){
                        document.getElementById("jsondata").innerHTML = "No se pudo leer el documento. Revise haber adjuntado el correcto.";
                    }else{
                        for(var j=0; j<objetoSICAS.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                            //Busca la fecha más antigua y la más reciente en SICAS para el nombre del xlsx.
                            var fechas =objetoSICAS[j]["Fecha Pago Recibo"]; 
                            const [dia, mes, anio] = fechas.split('/');
                            const fecha1 = new Date(+anio, +mes - 1, +dia);
                            if(fecha1>fechamax){
                                fechamax=fecha1;
                            }
                            if(fecha1<fechamin){
                                fechamin=fecha1;
                            }
                            //Divide la póliza de SICAS por '-' . La posición 2 es la fianza y la 3 es la inclusión
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
                            //Manda a llamar a la función de búsqueda
                        resultObject = search(num, inclusion, endo, objetoBerkley);
                        console.log(resultObject);
                        console.log("Número de registros en sicas: "+j);
                        document.getElementById("jsondata").innerHTML = tabla+tablanoencontrados+"</table>"; // DEL "+fechamin.getDate()+" "+month[+fechamin.getMonth()+1]+" "+fechamin.getFullYear()+" AL "+fechamax.getDate()+" "+month[+fechamax.getMonth()+1]+" "+fechamax.getFullYear();;//+month[messicas]+" Año: "+aniosicas; //Se manda la tabla pero no se va a ver porque tiene HIDDEN
                    }
                        ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente.
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
//fn: filename
function ExportToExcel(type, fn, dl) {// función que convierte a excel
    var elt = document.getElementById('BerkleyFianzas');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    var nombre ='CONCILIACIÓN BERKLEY FIANZAS DEL '+fechamin.getDate()+" "+month[+fechamin.getMonth()+1]+" "+fechamin.getFullYear()+" AL "+fechamax.getDate()+" "+month[+fechamax.getMonth()+1]+" "+fechamax.getFullYear()+".";
    return dl ? //It will attempt to force a client-side download.
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || (nombre + (type || 'xlsx')));
}
