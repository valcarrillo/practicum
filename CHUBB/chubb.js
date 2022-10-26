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
let objetoCHUBB; //Array de objetos en el que se va a guardar CHUBB
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
                objetoCHUBB = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet]); //Nombre del array
             console.log(objetoCHUBB);
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
                  let tabladiferencias ="<table id='CHUBBFianzas' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Póliza</th><th>Endoso</th><th>Prima Neta</th><th>% Comisión</th><th>Tipo Comisión</th><th>Comisiones</th><th>Diferencia de Comisión</th><th>No coincide:</th></tr>";
                  let resultObject;
                  let CHUBB;
                  let noencontrados=[];
                  var encontrar;
                  let tablanoencontrados="";    //Hay dos tablas, la de error y no encontrados. Se unen al final para tener más orden.
                  let tablaiguales=""; //Aquí se almacenan los que no tienen diferencias. Solo es por estilo.
                  var err="NO SE ENCONTRÓ LA PÓLIZA";

                  //Encontrar un valor ahí adentro
                  search = (poliza, CHUBB, ArraySICAS) => {
                    for (let i=0; i < ArraySICAS.length; i++) {
                        encontrar=0;
                        var SICASendoso=ArraySICAS[i].Endoso;
                        if (ArraySICAS[i].Poliza == poliza) {
                            console.log("Si coincide las polizas");
                            if (SICASendoso == CHUBB.Endoso) {
                               // if (SICASendoso == endo) {
                                    encontrar=1;

                                //compara las primas netas y si son diferentes las mete en la tabla.
                                //FUNCIÓN QUE HACE EL TIPO DE CAMBIO
                                importeSicas=Math.round(ArraySICAS[i]["Importe"]*ArraySICAS[i].TC*100)/100;// Esto es para que solo cuente los primeros dos decimales
                                importeCHUBB=Math.round(CHUBB.ComisionMto*100)/100;
                                console.log("IMPORTE SICAS: "+importeSicas);
                                console.log("IMPORTE CHUBB: "+importeCHUBB);
                                    
                                    var diferencia= Math.round((importeCHUBB -importeSicas)*100)/100;
                                    var tipodif;
                                    if(CHUBB.PNetaMto !=ArraySICAS[i]["PrimaNeta"] && CHUBB["% Comision"] !=ArraySICAS[i]["% Participacion"]){
                                        tipodif="Prima Neta y % Comisión";
                                    }else if(CHUBB.PNetaMto !=ArraySICAS[i]["PrimaNeta"]){
                                        tipodif="Prima Neta";
                                    }else if(CHUBB["% Comision"] !=ArraySICAS[i]["% Participacion"]){
                                        tipodif="% Comisión";
                                    //}else if(CHUBB["TIPO CAMBIO"] !=ArraySICAS[i].TC){
                                            //tipodif="Tipo de Cambio";
                                    }else{
                                        tipodif="Total Comisión";
                                    }
                                    console.log("La diferencia es de"+diferencia);
                                if(importeCHUBB != importeSicas){
                                    tabladiferencias=tabladiferencias+"<tr><td style='background-color:#8495cb'>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td>"+CHUBB.PNetaMto+"</td><td>"+CHUBB["% Comision"]+"</td><td>Comisión base</td><td>"+CHUBB.ComisionMto+"</td><td style='color:#9c0b0be7'>"+diferencia+"</td><td>"+tipodif+"</td></tr>";
                                }else{
                                    tablaiguales=tablaiguales+"<tr><td style='background-color:#8495cb'>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td>"+CHUBB.PNetaMto+"</td><td>"+CHUBB["% Comision"]+"</td><td>Comisión Base</td><td>"+CHUBB.ComisionMto+"</td><td>"+diferencia+"</td><td></td></tr>";
                                }
                              return ArraySICAS;
                                //}else{
                                   // err="NO SE ENCONTRÓ LA PÓLIZA";
                               // }
                           }
                          }
                    }
                      if(encontrar==0){ // Encontrar es una bandera. Si no se encuentra, se incluye lo de abajo
                        CHUBB.Endoso=0;
                        noencontrados.push(CHUBB);
                      //tablanoencontrados= tablanoencontrados+"<tr><td style='background-color:#8495cb'>"+poliza+"</td><td>"+endoso+"</td><td>"+pnetamto+"</td><td>"+porcentaje+"</td><td>Comisión Base</td><td>"+comisionmto+"</td><td></td><td>NO SE ENCONTRÓ</td></tr>";
                        return CHUBB;
                        }
                      encontrar=0; 
                    }
                    
                    var pnetamto =0;
                    var comisionmto=0, comisionrecargo=0;
                    var claveanterior="", polizaanterior="", endosoanterior="";
                    var repeticion=1;
                    let objetoNuevochubb=[];
                    let jsonObj;
                    let Sicassinendoso=[];
                    let Sicasconendoso=[];
                    //###########FUNCIÓN QUE MANDA A LLAMAR LA BÚSQUEDA################
                    if(typeof objetoSICAS[0].Poliza === 'undefined' || objetoCHUBB[0].PolizaId==='undefined'){
                        document.getElementById("jsondata").innerHTML = "No se pudo leer el documento. Revise haber adjuntado el correcto.";
                    }else{
                        for(var j=0; j<objetoSICAS.length; j++){
                            if(objetoSICAS[j].Endoso===''){
                                objetoSICAS[j].Endoso='0';
                                Sicassinendoso.push(objetoSICAS[j]);
                                
                            }else{
                                Sicasconendoso.push(objetoSICAS[j]);
                                
                            }

                        }
                        console.log("Sicassinendoso");
                        console.log(Sicassinendoso);
                        console.log("Sicasconendoso");
                        console.log(Sicasconendoso);
                        for(var j=0; j<objetoCHUBB.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en CHUBB
                            
                            if(objetoCHUBB[j].hasOwnProperty('PolizaId')){
                                //Busca la fecha más antigua y la más reciente en SICAS para el nombre del xlsx.                        
                                var fechas =objetoCHUBB[j].Fecha.split(' '); 
                                const [mes, dia, anio] = fechas[0].split('/');
                                const fecha1 = new Date(+anio, +mes - 1, +dia);
                                if(fecha1>fechamax){
                                    fechamax=fecha1;
                                }
                                if(fecha1<fechamin){
                                    fechamin=fecha1;
                                }


                            //console.log(objetoCHUBB[j]);
                            CHUBB=objetoCHUBB[j];
                            clave=objetoCHUBB[j].ClaveId;
                            if(clave != null){
                                clave = clave.replace(/\s/g, '');
                            }
                            poliza=objetoCHUBB[j].PolizaId;
                            poliza = poliza.replace(/\s/g, '');
                            endoso=objetoCHUBB[j].Endoso;
                            recibo=objetoCHUBB[j].Recibo;

                            //SE VA A CREAR UN NUEVO ARRAY DE OBJETOS JSON SUMANDO LAS PÓLIZAS QUE SEAN IGUALES
                            if(clave==claveanterior && poliza==polizaanterior && endoso==endosoanterior){
                                pnetamto=pnetamto+objetoCHUBB[j].PNetaMto;
                                comisionmto=comisionmto+objetoCHUBB[j].ComisionMto;
                               // console.log(objetoCHUBB[j].ComisionSobreRecargoMto);
                                comisionrecargo=comisionrecargo+objetoCHUBB[j].ComisionSobreRecargoMto;
                                repeticion++;
                            }else{
                                porcentajecomision=Math.round((comisionmto/pnetamto)*100);
                                jsonObj={"ClaveId": claveanterior, "PolizaId": polizaanterior,"Endoso": endosoanterior, "Recibo": objetoCHUBB[j-1].Recibo,"PNetaMto": Math.round(pnetamto*100)/100, "ComisionMto": Math.round(comisionmto*100)/100, "ComisionSobreRecargoMto": Math.round(comisionrecargo*100)/100, "% Comision": porcentajecomision};
                                //console.log(jsonObj);
                                objetoNuevochubb.push(jsonObj);
                               // console.log(objetoNuevochubb);
                                jsonObj={};
                                pnetamto=objetoCHUBB[j].PNetaMto;
                                comisionmto=objetoCHUBB[j].ComisionMto;
                                repeticion=1;
                            }
                            polizaanterior=poliza;
                            endosoanterior=endoso;
                            claveanterior=clave;
                            
                            }
                        }
                     for(var j=1; j<objetoNuevochubb.length; j++){
                            poliza=objetoNuevochubb[j].ClaveId+" "+objetoNuevochubb[j].PolizaId;
                            console.log(objetoNuevochubb[j]);
                            resultObject = search(poliza, objetoNuevochubb[j], Sicasconendoso);
                     }
                     for(var j=0; j<noencontrados.length; j++){
                        search(poliza, 0,  noencontrados[j], Sicassinendoso);
                     }
                       
                        document.getElementById("jsondata").innerHTML = tabladiferencias+tablanoencontrados+tablaiguales+"<tr><td>DEL</td><td>"+fechamin.getDate()+" "+month[+fechamin.getMonth()+1]+" "+fechamin.getFullYear()+"</td><td>AL</td><td>"+fechamax.getDate()+" "+month[+fechamax.getMonth()+1]+" "+fechamax.getFullYear()+"</td><td># Registros</td><td>"+(j-1)+"</td><td></td><td></td></tr></table>"; // DEL "+fechamin.getDate()+" "+month[+fechamin.getMonth()+1]+" "+fechamin.getFullYear()+" AL "+fechamax.getDate()+" "+month[+fechamax.getMonth()+1]+" "+fechamax.getFullYear();;//+month[messicas]+" Año: "+aniosicas; //Se manda la tabla pero no se va a ver porque tiene HIDDEN
                    }
                        //ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente.
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
