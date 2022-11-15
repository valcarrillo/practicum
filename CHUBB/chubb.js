//https://levelup.gitconnected.com/how-to-convert-excel-file-into-json-object-by-using-javascript-9e95532d47c5
//convertir en EXCEL https://codepedia.info/javascript-export-html-table-data-to-excel
let selectedFile=[];
let selectedFile2;
let numarchivos=0;
console.log(window.XLSX);
document.getElementById('inputA').addEventListener("change", (event) => { //Lee estado de cuenta
    const files = event.target.files;
    for (numarchivos=0; numarchivos < files.length; numarchivos++) {
        selectedFile[numarchivos] = event.target.files[numarchivos]; //selectedFile[0]=file[0]
     }
})

document.getElementById('inputSicas').addEventListener("change", (event) => {// Lee SICAS
    selectedFile2 = event.target.files[0];
}
)

let objetoSICAS; //Array de objetos en el que se va a guarar SICAS
let objetoCHUBB=[]; //Array de objetos en el que se va a guardar CHUBB
var fechamax = new Date("2000-01-02"); // (YYYY-MM-DD)
var fechamin = new Date("2100-01-01"); // (YYYY-MM-DD)
const month = ["Nada","ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"];

var reng_EC=0;  

document.getElementById('button').addEventListener("click", () => {
    jsonObj={"Asegurado": "NO SIRVE SOLO PARA PRUEBA", "ClaveId": "AAA", "PolizaId": "00000000","Endoso": "00000", "Recibo": "0", "TotalRecibo": "0","PNetaMto": "000", "ComisionMto": "00000", "ComisionSobreRecargoMto": "000", "TipoMov": "0"};
    objetoCHUBB.push(jsonObj);
    jsonObj={};
    if(selectedFile){ //Función para convertir Edo de Cuenta en array de objetos
        for(i=0; i<numarchivos; i++){ //ciclo que lee cada selected file
            let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile[i]);
        fileReader.onload = (event)=>{
         let data1 = event.target.result;
         let workbook1 = XLSX.read(data1,{type:"binary"});       
         workbook1.SheetNames.forEach(sheet => {
                                                           // EL RANGO ES LO GRANDE DEL ENCABEZADO                                         
                objeto = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet]); //Nombre del array
                for(var j=1; j<objeto.length; j++){
                    objetoCHUBB.push(objeto[j]);
                }
         });
        }
        }
        console.log("objetoCHUBB:");
         console.log(objetoCHUBB);
         if(selectedFile2){ //Función que convierte SICAS en array de objetos
            let fileReader = new FileReader();
            fileReader.readAsBinaryString(selectedFile2);
            fileReader.onload = (event)=>{
             let data2 = event.target.result;
             let workbook2 = XLSX.read(data2,{type:"binary"});
             workbook2.SheetNames.forEach(sheet => {
                  objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //Nombre del array
                  console.log(objetoSICAS);
                  var reng_SICAS=objetoSICAS.length;
                    //La tabla tiene atributo HIDDEN para que no se vea, pero ahí está.
                  let tabladiferencias ="<table id='CHUBBFianzas' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' ><tr><th>SICAS</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th>CHUBB</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>";
                  tabladiferencias=tabladiferencias+"<tr><td>Nombre Asegurado o Fiado</td><td>Poliza</td><td>Endoso</td><td>Moneda</td><td>Serie Recibo</td><td>Tipo Cambio</td><td>PrimaNeta</td><td>Tipo Comision</td><td>Importe</td><td>% Participacion</td><td></td><td>Nombre Asegurado o Fiado</td><td>Poliza</td><td>Endoso</td><td>Moneda</td><td>Serie Recibo</td><td>% Comisión</td><td>Comisión</td><td>Tipo Cambio</td><td>Diferencia</td><td>Incidencia</td></tr>";
                  let resultObject;
                  let noencontrados=[];
                  var encontrar;
                  let tablanoencontrados="";    //Hay dos tablas, la de error y no encontrados. Se unen al final para tener más orden.
                  let tablaiguales=""; //Aquí se almacenan los que no tienen diferencias. Solo es por estilo.

                  //Encontrar un valor ahí adentro
                  search = (poliza, CHUBB, ArraySICAS) => {
                    for (let i=0; i < ArraySICAS.length; i++) {
                        encontrar=0;
                        var SICASendoso=ArraySICAS[i].Endoso;
                        if (ArraySICAS[i].Poliza == poliza) {
                            if (SICASendoso == CHUBB.Endoso) {
                               if (ArraySICAS[i].Serie == CHUBB.Serie) {
                                    if (ArraySICAS[i]["Tipo Comision"] == CHUBB["Tipo Comisión"]) {
                                       
                                    
                                //compara las primas netas y si son diferentes las mete en la tabla.
                                //FUNCIÓN QUE HACE EL TIPO DE CAMBIO
                                importeSicas=Math.round(ArraySICAS[i]["Importe"]*ArraySICAS[i].TC*100)/100;// Esto es para que solo cuente los primeros dos decimales
                              
                                    
                                    var diferencia= Math.round((CHUBB.Importe -importeSicas)*100)/100;
                                    var tipodif;
                                    if(CHUBB.PNetaMto !=ArraySICAS[i]["PrimaNeta"] && CHUBB["% Comision"] !=ArraySICAS[i]["% Participacion"]){
                                        tipodif="Prima Neta y % Comisión";
                                    }else if(CHUBB.PNetaMto !=ArraySICAS[i]["PrimaNeta"]){
                                        tipodif="Prima Neta";
                                    }else if(CHUBB["% Comision"] !=ArraySICAS[i]["% Participacion"]){
                                        tipodif="% Comisión";
                                    }else{
                                        tipodif="Total Comisión";
                                    }

                                if(CHUBB.Importe!= importeSicas){
                                    tabladiferencias=tabladiferencias+"<tr><td>"+CHUBB.Asegurado+"</td><td>"+ArraySICAS[i].Poliza+"</td><td>"+ArraySICAS[i].Endoso+"</td><td>"+ArraySICAS[i].Moneda+"</td><td>'"+ArraySICAS[i].Serie+"'</td><td>"+ArraySICAS[i].TC+"</td><td>"+ArraySICAS[i].PrimaNeta+"</td><td>"+ArraySICAS[i]["Tipo Comision"]+"</td><td>"+ArraySICAS[i].Importe+"</td><td>"+ArraySICAS[i]["% Participacion"]+"</td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td>"+diferencia+"</td><td>"+tipodif+"</td></tr>";
                                }else{
                                    tablaiguales=tablaiguales+"<tr><td>"+CHUBB.Asegurado+"</td><td>"+ArraySICAS[i].Poliza+"</td><td>"+ArraySICAS[i].Endoso+"</td><td>"+ArraySICAS[i].Moneda+"</td><td>'"+ArraySICAS[i].Serie+"'</td><td>"+ArraySICAS[i].TC+"</td><td>"+ArraySICAS[i].PrimaNeta+"</td><td>"+ArraySICAS[i]["Tipo Comision"]+"</td><td>"+ArraySICAS[i].Importe+"</td><td>"+ArraySICAS[i]["% Participacion"]+"</td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td>"+diferencia+"</td><td></td></tr>";
                                }
                              return ArraySICAS;
                            }
                            
                            }
                           }
                          }
                    }
                      if(encontrar==0){ // Encontrar es una bandera. Si no se encuentra, se incluye lo de abajo
                        noencontrados.push(CHUBB);
                      //tablanoencontrados= tablanoencontrados+"<tr><td style='background-color:#8495cb'>"+poliza+"</td><td>"+endoso+"</td><td>"+pnetamto+"</td><td>"+porcentaje+"</td><td>Comisión Base</td><td>"+comisionmto+"</td><td></td><td>NO SE ENCONTRÓ</td></tr>";
                        return CHUBB;
                        }
                      encontrar=0; 
                    }

                    //////////////////
                    //Encontrar un valor ahí adentro
                  search2 = (poliza, CHUBB, ArraySICAS) => {
                    for (let i=0; i < ArraySICAS.length; i++) {
                        encontrar=0;
                        if (ArraySICAS[i].Poliza == poliza) {
                            if (ArraySICAS[i].Serie == CHUBB.Serie) {
                                if (ArraySICAS[i]["Tipo Comision"] == CHUBB["Tipo Comisión"]) {
                                encontrar=1;

                            //compara las primas netas y si son diferentes las mete en la tabla.
                            //FUNCIÓN QUE HACE EL TIPO DE CAMBIO
                            importeSicas=Math.round(ArraySICAS[i]["Importe"]*ArraySICAS[i].TC*100)/100;// Esto es para que solo cuente los primeros dos decimales
                          
                                
                                var diferencia= Math.round((CHUBB.Importe -importeSicas)*100)/100;
                                var tipodif;
                                if(CHUBB.PNetaMto !=ArraySICAS[i]["PrimaNeta"] && CHUBB["% Comision"] !=ArraySICAS[i]["% Participacion"]){
                                    tipodif="Prima Neta y % Comisión";
                                }else if(CHUBB.PNetaMto !=ArraySICAS[i]["PrimaNeta"]){
                                    tipodif="Prima Neta";
                                }else if(CHUBB["% Comision"] !=ArraySICAS[i]["% Participacion"]){
                                    tipodif="% Comisión";
                                }else{
                                    tipodif="Total Comisión";
                                }
                             
                            if(CHUBB.Importe != importeSicas){
                                tabladiferencias=tabladiferencias+"<tr><td>"+CHUBB.Asegurado+"</td><td>"+ArraySICAS[i].Poliza+"</td><td>"+ArraySICAS[i].Endoso+"</td><td>"+ArraySICAS[i].Moneda+"</td><td>'"+ArraySICAS[i].Serie+"'</td><td>"+ArraySICAS[i].TC+"</td><td>"+ArraySICAS[i].PrimaNeta+"</td><td>"+ArraySICAS[i]["Tipo Comision"]+"</td><td>"+ArraySICAS[i].Importe+"</td><td>"+ArraySICAS[i]["% Participacion"]+"</td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td>"+diferencia+"</td><td>"+tipodif+"</td></tr>";
                            }else{
                                tablaiguales=tablaiguales+"<tr><td>"+CHUBB.Asegurado+"</td><td>"+ArraySICAS[i].Poliza+"</td><td>"+ArraySICAS[i].Endoso+"</td><td>"+ArraySICAS[i].Moneda+"</td><td>'"+ArraySICAS[i].Serie+"'</td><td>"+ArraySICAS[i].TC+"</td><td>"+ArraySICAS[i].PrimaNeta+"</td><td>"+ArraySICAS[i]["Tipo Comision"]+"</td><td>"+ArraySICAS[i].Importe+"</td><td>"+ArraySICAS[i]["% Participacion"]+"</td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+ CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td>"+diferencia+"</td><td></td></tr>";
                            }
                          return ArraySICAS;
            
                                }
                        }
                          }
                    }
                      if(encontrar==0){ // Encontrar es una bandera. Si no se encuentra, se incluye lo de abajo
                        tablanoencontrados= tablanoencontrados+"<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>"+CHUBB["Tipo Comisión"] +"</td><td></td><td></td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+ CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td></td><td>NO SE ENCONTRÓ</td></tr>";
                        return CHUBB;
                        }
                      encontrar=0; 
                    }

                    ////////////////
                    var pnetamto =0;
                    var comisionmto=0, comisionrecargo=0,  comisionespecial=0;
                    var claveanterior="", polizaanterior="", endosoanterior="";
                    var registers=0;
                    let objetoNuevochubb=[];
                    let jsonObj;
                    let Sicassinendoso=[];
                    let Sicasconendoso=[];
                    //###########FUNCIÓN QUE MANDA A LLAMAR LA BÚSQUEDA################
                    console.log(objetoCHUBB[0]);
                    console.log(objetoCHUBB[0].PolizaId);
                
                try {
                    if(typeof objetoSICAS[0].Poliza === 'undefined' || !(objetoCHUBB[0].hasOwnProperty('PolizaId'))){
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
                        jsonObj={"Asegurado": "NO SIRVE SOLO PARA PRUEBA", "ClaveId": "AAA", "PolizaId": "00000000","Endoso": "00000", "Recibo": "0", "TotalRecibo": "0","PNetaMto": "000", "ComisionMto": "00000", "ComisionSobreRecargoMto": "000", "TipoMov": "0"};
                        objetoCHUBB.push(jsonObj);
                        jsonObj={};
                        console.log(objetoCHUBB);
                        for(var j=1; j<objetoCHUBB.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en CHUBB
                            
                            if(objetoCHUBB[j].hasOwnProperty('PolizaId')){
                                //Busca la fecha más antigua y la más reciente en SICAS para el nombre del xlsx. 
                                if(objetoCHUBB[j].hasOwnProperty('Fecha')){                       
                                    var fechas =objetoCHUBB[j].Fecha.split(' '); 
                                    const [mes, dia, anio] = fechas[0].split('/');
                                    const fecha1 = new Date(+anio, +mes - 1, +dia);
                                    if(fecha1>fechamax){
                                        fechamax=fecha1;
                                    }
                                    if(fecha1<fechamin){
                                        fechamin=fecha1;
                                    }
                                }

                            //console.log(objetoCHUBB[j]);
                            CHUBB=objetoCHUBB[j];
                            clave=objetoCHUBB[j].ClaveId;
                            if(clave != null){
                                clave = clave.toString().replace(/\s/g, '');
                            }else{
                                clave='';
                                objetoCHUBB[j].ClaveId=clave;
                            }
                            if(!(objetoCHUBB[j].hasOwnProperty('Asegurado'))){
                                objetoCHUBB[j].Asegurado=' ';
                            }
                            poliza=objetoCHUBB[j].PolizaId;
                            poliza = poliza.toString().replace(/\s/g, '');
                            endoso=objetoCHUBB[j].Endoso;
                            recibo=objetoCHUBB[j].Recibo;

                            //SE VA A CREAR UN NUEVO ARRAY DE OBJETOS JSON SUMANDO LAS PÓLIZAS QUE SEAN IGUALES
                            if(clave==claveanterior && poliza==polizaanterior && endoso==endosoanterior){
                                pnetamto=pnetamto+objetoCHUBB[j].PNetaMto;
                                comisionmto=comisionmto+objetoCHUBB[j].ComisionMto;
                                comisionrecargo=comisionrecargo+objetoCHUBB[j].ComisionSobreRecargoMto;
                                comisionespecial=comisionespecial+objetoCHUBB[j].Comision2
                            }else{
                                if(pnetamto!=0){
                                    porcentajecomision=Math.round((comisionmto/pnetamto)*100);
                                }else{
                                    porcentajecomision="";
                                }
                                if(objetoCHUBB[j-1].hasOwnProperty('Recibo') && objetoCHUBB[j-1].hasOwnProperty('TotalRecibo')){ 
                                    serie=objetoCHUBB[j-1].Recibo.toString().padStart(3, "0")+"/"+objetoCHUBB[j-1].TotalRecibo.toString().padStart(3, "0");
                                }else{
                                   serie="000/000";
                                }
                                if(objetoCHUBB[j-1].TipoMov==6){
                                    tipocomision="Bono Internet/ Comisión por derechos";
                                    jsonObj={"Asegurado": objetoCHUBB[j-1].Asegurado, "ClaveId": claveanterior, "PolizaId": polizaanterior,"Endoso": endosoanterior, "Serie": serie,"PNetaMto": Math.round(pnetamto*100)/100, "Importe": Math.round(comisionmto*100)/100, "% Comision": porcentajecomision, "Tipo Comisión": tipocomision};
                                    objetoNuevochubb.push(jsonObj);
                                    jsonObj={};
                                }else{
                                    if(comisionrecargo!=0 || comisionrecargo!='NaN'){
                                        tipocomision="Comisión de Recargos";
                                        jsonObj={"Asegurado": objetoCHUBB[j-1].Asegurado, "ClaveId": claveanterior, "PolizaId": polizaanterior,"Endoso": endosoanterior, "Serie": serie,"PNetaMto": Math.round(pnetamto*100)/100, "Importe": Math.round(comisionrecargo*100)/100, "% Comision": porcentajecomision, "Tipo Comisión": tipocomision};
                                        objetoNuevochubb.push(jsonObj);
                                        jsonObj={};
                                    }
                                    if(comisionespecial!=0 || !(isNaN(comisionespecial))){
                                        tipocomision="Comisión Especial";
                                        jsonObj={"Asegurado": objetoCHUBB[j-1].Asegurado, "ClaveId": claveanterior, "PolizaId": polizaanterior,"Endoso": endosoanterior, "Serie": serie,"PNetaMto": Math.round(pnetamto*100)/100, "Importe": Math.round(comisionespecial*100)/100, "% Comision": porcentajecomision, "Tipo Comisión": tipocomision};
                                        objetoNuevochubb.push(jsonObj);
                                        jsonObj={};
                                    }
                                    tipocomision="Comisión Base o de Neta";
                                    jsonObj={"Asegurado": objetoCHUBB[j-1].Asegurado, "ClaveId": claveanterior, "PolizaId": polizaanterior,"Endoso": endosoanterior, "Serie": serie,"PNetaMto": Math.round(pnetamto*100)/100, "Importe": Math.round(comisionmto*100)/100, "% Comision": porcentajecomision, "Tipo Comisión": tipocomision}; 
                                   
                                    //console.log(jsonObj);
                                    objetoNuevochubb.push(jsonObj);
                                    jsonObj={};
                                    pnetamto=objetoCHUBB[j].PNetaMto;
                                    comisionmto=objetoCHUBB[j].ComisionMto;
                                    comisionrecargo=objetoCHUBB[j].ComisionSobreRecargoMto;
                                    comisionespecial=objetoCHUBB[j].Comision2
                                }
                           
                            }
                            polizaanterior=poliza;
                            endosoanterior=endoso;
                            claveanterior=clave;
                            
                            }
                        }
                        console.log(objetoNuevochubb);
                        var reng_EC=objetoNuevochubb.length;  
                     for(var j=1; j<objetoNuevochubb.length; j++){
                            poliza=objetoNuevochubb[j].ClaveId+" "+objetoNuevochubb[j].PolizaId;
                            //console.log(objetoNuevochubb[j]);
                            resultObject = search(poliza, objetoNuevochubb[j], Sicasconendoso);
                            registers++;
                     }
                    for(var j=0; j<noencontrados.length; j++){
                        poliza=noencontrados[j].ClaveId+" "+noencontrados[j].PolizaId;
                        //console.log(noencontrados[j]);
                        search2(poliza, noencontrados[j], Sicassinendoso);
                    }
                    document.getElementById("numregistros").innerHTML = "Renglones Estado de Cuenta: "+reng_EC+"\nRenglones SICAS: "+reng_SICAS+"\n";
                        document.getElementById("jsondata").innerHTML = tabladiferencias+tablaiguales+tablanoencontrados+"<tr><td>DEL</td><td>"+fechamin.getDate()+" "+month[+fechamin.getMonth()+1]+" "+fechamin.getFullYear()+"</td><td>AL</td><td>"+fechamax.getDate()+" "+month[+fechamax.getMonth()+1]+" "+fechamax.getFullYear()+"</td><td># Registros</td><td>"+(registers)+"</td><td></td><td></td></tr></table>"; // DEL "+fechamin.getDate()+" "+month[+fechamin.getMonth()+1]+" "+fechamin.getFullYear()+" AL "+fechamax.getDate()+" "+month[+fechamax.getMonth()+1]+" "+fechamax.getFullYear();;//+month[messicas]+" Año: "+aniosicas; //Se manda la tabla pero no se va a ver porque tiene HIDDEN
                    }
                        //ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente.
                        objetoCHUBB=[];
                        if(resultObject==0){
                            document.getElementById("jsondata").innerHTML = "No se encontró ninguna fianza";

                        }
                    } catch (error) {
                        document.getElementById("jsondata").innerHTML = "Algo salió mal al leer el documento. Revise que el encabezado tenga el formato correcto. Error: "+error;
                        // expected output: ReferenceError: nonExistentFunction is not defined
                        // Note - error messages will vary depending on browser
                      }
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
