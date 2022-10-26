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

                  let tabladiferencias ="<table id='BerkleyFianzas' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Póliza</th><th>Endoso</th><th>Prima Neta</th><th>% Comisión</th><th>Tipo Comisión</th><th>Comisiones</th><th>Diferencia de Comisión</th><th>No coincide:</th></tr>";

                  let tabla ="<table id='BerkleyFianzas' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' HIDDEN> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Tipo Comisión</th><th>Total Comisión</th><th>Diferencia de Comisión</th><th>No coincide:</th></tr>";

                  let resultObject;
                  let berkley;
                  var encontrar;
                  let tablanoencontrados="";    //Hay dos tablas, la de error y no encontrados. Se unen al final para tener más orden.
                  let tablaiguales=""; //Aquí se almacenan los que no tienen diferencias. Solo es por estilo.
                  var err="NO SE ENCONTRÓ LA PÓLIZA";

                  //Encontrar un valor ahí adentro
                  search = (key, inclu, endo, ArraySICAS) => {
                      for (let i=0; i < ArraySICAS.length; i++) {
                        encontrar=0;
                        //Divide la póliza de SICAS por '-' . La posición 2 es la fianza y la 3 es la inclusión
                        var pol = ArraySICAS[i].Poliza.split('-'),
                        SICASpoliza = pol[2];
                        SICASinclusion= pol[3];
                        if(typeof  SICASinclusion === 'undefined'){
                            SICASinclusion=0;
                        }else{
                            SICASinclusion= + SICASinclusion;
                        }
                        numpoliza = +SICASpoliza;
                        var SICASendoso=ArraySICAS[i].Endoso;
                        if(typeof SICASendoso === 'undefined'){
                            SICASendoso= +1; //Si el endoso no está en SICAS entonces se registra como 1
                        }else{
                            SICASendoso= +SICASendoso +1;
                        }
                        //Busqueda
                          if (numpoliza == key) {
                            if (SICASinclusion == inclu) {
                                if (SICASendoso == endo) {
                                    encontrar=1;

                                //compara las primas netas y si son diferentes las mete en la tabla.
                                //FUNCIÓN QUE HACE EL TIPO DE CAMBIO
                                importeSicas=Math.round(ArraySICAS[i]["Importe"]*ArraySICAS[i].TC*100)/100;// Esto es para que solo cuente los primeros dos decimales
                                importeBerkley=Math.round(berkley["COMISIONES"]*berkley["TIPO CAMBIO"]*100)/100;
                                console.log("IMPORTE SICAS: "+importeSicas);
                                console.log("IMPORTE BERKLEY: "+importeBerkley);
                                    
                                    var diferencia= Math.round((importeBerkley -importeSicas)*100)/100;
                                    var tipodif;
                                    if(berkley["PRIMA NETA"] !=ArraySICAS[i]["PrimaNeta"] && berkley["% COMISION"] !=ArraySICAS[i]["% Participacion"]){
                                        tipodif="Prima Neta y % Comisión";
                                    }else if(berkley["PRIMA NETA"] !=ArraySICAS[i]["PrimaNeta"]){
                                        tipodif="Prima Neta";
                                    }else if(berkley["% COMISION"] !=ArraySICAS[i]["% Participacion"]){
                                        tipodif="% Comisión";
                                    }else if(berkley["TIPO CAMBIO"] !=ArraySICAS[i].TC){
                                            tipodif="Tipo de Cambio";
                                    }else{
                                        tipodif="Total Comisión";
                                    }
                                    console.log("La diferencia es de"+diferencia);
                                if(importeBerkley != importeSicas){
                                    tabladiferencias=tabladiferencias+"<tr><td style='background-color:#8495cb'>"+berkley.FIANZA+"-"+berkley.INCLUSION+"</td><td>"+berkley.MOVIMIENTO+"</td><td>"+berkley["PRIMA NETA"]+"</td><td>"+berkley["% COMISION"]+"</td><td>"+berkley["TIPO COMISION"]+"</td><td>"+berkley.COMISIONES+"</td><td style='color:#9c0b0be7'>"+diferencia+"</td><td>"+tipodif+"</td></tr>";
                                }else{
                                    tablaiguales=tablaiguales+"<tr><td style='background-color:#8495cb'>"+berkley.FIANZA+"-"+berkley.INCLUSION+"</td><td>"+berkley.MOVIMIENTO+"</td><td>"+berkley["PRIMA NETA"]+"</td><td>"+berkley["% COMISION"]+"</td><td>"+berkley["TIPO COMISION"]+"</td><td>"+berkley.COMISIONES+"</td><td>"+diferencia+"</td><td></td></tr>";
                                }
                              return ArraySICAS[i];
                                }else{
                                    err="NO SE ENCONTRÓ LA PÓLIZA";
                                }
                            }
                          }
                      }
                      if(encontrar==0){ // Encontrar es una bandera. Si no se encuentra, se incluye lo de abajo
                      tablanoencontrados= tablanoencontrados+"<tr><td style='background-color:#8495cb'>"+berkley.FIANZA+"-"+berkley.INCLUSION+"</td><td>"+berkley.MOVIMIENTO+"</td><td>"+berkley["PRIMA NETA"]+"</td><td>"+berkley["% COMISION"]+"</td><td>"+berkley["TIPO COMISION"]+"</td><td>"+berkley.COMISIONES+"</td><td></td><td>NO SE ENCONTRÓ</td></tr>";
                        return ArraySICAS;
                        }
                      encontrar=0; 
                    }
                    

                    //###########FUNCIÓN QUE MANDA A LLAMAR LA BÚSQUEDA################
                    if(typeof objetoSICAS[0].Poliza === 'undefined' || objetoBerkley[0].FIANZA==='undefined'){
                        document.getElementById("jsondata").innerHTML = "No se pudo leer el documento. Revise haber adjuntado el correcto.";
                    }else{
                        for(var j=0; j<objetoBerkley.length-1; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                            //Busca la fecha más antigua y la más reciente en SICAS para el nombre del xlsx.
                            
                                var fechas =objetoBerkley[j]["FECHA APLICACION"]; 
                                console.log(fechas);
                                const [dia, mes, anio] = fechas.split('/');
                                const fecha1 = new Date(+anio, +mes - 1, +dia);
                                if(fecha1>fechamax){
                                    fechamax=fecha1;
                                }
                                if(fecha1<fechamin){
                                    fechamin=fecha1;
                                }
                            
                            berkley=objetoBerkley[j];
                            poliza=objetoBerkley[j].FIANZA
                            inclusion=objetoBerkley[j].INCLUSION
                            movimiento=objetoBerkley[j].MOVIMIENTO
                            //Manda a llamar a la función de búsqueda
                        resultObject = search(poliza, inclusion, movimiento, objetoSICAS);
                        console.log(resultObject);
                        console.log("Número de registros en berkley: "+j);
                        document.getElementById("jsondata").innerHTML = tabladiferencias+tablanoencontrados+tablaiguales+"<tr><td>DEL</td><td>"+fechamin.getDate()+" "+month[+fechamin.getMonth()+1]+" "+fechamin.getFullYear()+"</td><td>AL</td><td>"+fechamax.getDate()+" "+month[+fechamax.getMonth()+1]+" "+fechamax.getFullYear()+"</td><td># Registros</td><td>"+j+"</td><td></td><td></td></tr></table>"; // DEL "+fechamin.getDate()+" "+month[+fechamin.getMonth()+1]+" "+fechamin.getFullYear()+" AL "+fechamax.getDate()+" "+month[+fechamax.getMonth()+1]+" "+fechamax.getFullYear();;//+month[messicas]+" Año: "+aniosicas; //Se manda la tabla pero no se va a ver porque tiene HIDDEN
                    }
                        //ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente.
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
