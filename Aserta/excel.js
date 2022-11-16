//https://levelup.gitconnected.com/how-to-convert-excel-file-into-json-object-by-using-javascript-9e95532d47c5
//convertir en EXCEL https://codepedia.info/javascript-export-html-table-data-to-excel
let selectedFile=[];
let selectedFile2;
let numarchivos=0;
console.log(window.XLSX);
var docAserta= document.getElementById('inputA').addEventListener("change", (event) => { //Lee estado de cuenta
    const files = event.target.files;
    for (numarchivos=0; numarchivos < files.length; numarchivos++) {
        selectedFile[numarchivos] = event.target.files[numarchivos];
     }
})
//BD SICAS
document.getElementById('inputSicas').addEventListener("change", (event) => {// Lee SICAS
    selectedFile2 = event.target.files[0];
}
)

let objetoSICAS; //Array de objetos en el que se va a guarar SICAS
let objetoAserta; //Array de objetos en el que se va a guardar Berkley
var fechamax = new Date("2000-01-02"); // (YYYY-MM-DD)
var fechamin = new Date("2100-01-01"); // (YYYY-MM-DD)
const month = ["Nada","ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"];

document.getElementById('button').addEventListener("click", () => {
    var reng_SICAS=0;
    var reng_EC=0;  
    num_ECT = document.getElementById("select").value
    //ESTRUCTURA TABLA
    let tabla ="<table id='Aserta' width='90%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' HIDDEN ><tr><td colspan='10'>SICAS</td><td>--</td><td colspan ='11'>ESTADOS DE CUENTA</td></tr><tr><th>Nombre Asegurado o Fiado</th><th>Póliza</th><th>Endoso</th><th>Moneda</th><th>Serie Recibo</th><th>Tipo Cambio</th><th>Prima Neta</th><th>Tipo Comisión</th><th>Importe</th><th>% Participación</th><th>--</th><th>Nombre Asegurado o Fiado</th><th>Póliza</th><th>Folio Factura</th><th>Endoso</th><th>Moneda</th><th>Serie Recibo</th><th>% Comisión</th><th>Comisión</th><th>Tipo Cambio</th><th>Diferencia</th><th>Incidencia</th></tr>";
    let tablaNA ="";
///Sacar datos de varios archivos
                    if(selectedFile){ //Función para convertir Edo de Cuenta en array de objetos
                        let fileReader = new FileReader();
                        fileReader.readAsBinaryString(selectedFile);
                        fileReader.onload = (event)=>{
                         let data1 = event.target.result;
                         let workbook1 = XLSX.read(data1,{type:"binary"});
                         console.log(workbook1);        
                         workbook1.SheetNames.forEach(sheet => {
                                objetoAserta = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet], {range:1}); //Nombre del array
                                console.log(objetoAserta)
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
                                  let resultObject;
                                  let aserta;
                                  var encontrar = 0;   
                                  //Encontrar un valor ahí adentro
                                  search = (key, ArreySICAS) => {
                                    reng_SICAS = ArreySICAS.length;
                                    //Recorremos arreglos de SICAS
                                      for (let i=0; i < ArreySICAS.length; i++) {
                                        var polizaSicas = String(ArreySICAS[i].Poliza)
                                        var comision = String(ArreySICAS[i]["Tipo Comision"])
                                        var errores="";
                                        var diferencias ="";
                                        var porcentaje ="";
                                        var comision_I = "";
                                        var importe_mxn ="";
                                        var TC=String(ArreySICAS[i]["TC"])
                                        //Obtener TC del Estado de Cuenta
                                        if(aserta['Moneda'] != 'MXN'){
                                            var TC_EstadoCuenta= Number(aserta["Comisión"])/ (Number(aserta["% de Comisión"])*Number(aserta["Prima Neta"]))
                                            TC_EstadoCuenta= Math.trunc(TC_EstadoCuenta*1000000)/10000
                                        }
                                        else{
                                            var TC_EstadoCuenta= 1
                                        }
                                          
                                          //Obtener importe mxn
                                          importe_mxn= (ArreySICAS[i]["Importe"]*ArreySICAS[i]["TC"])
                                          if (polizaSicas == key ) { 
                                            encontrar++;
                                            if(comision == 'Comisión Base o de Neta'){
                                                porcentaje = String(aserta["% de Comisión"])
                                                comision_I = String(aserta["Comisión"])
                                            }
                                            else if(comision == 'Comisión de Derechos'){
                                                porcentaje = String(aserta["% Maquila"])
                                                comision_I = String(aserta["Comisión Gtos. Exp."])
                                            }
                                            else if(comision == 'Comisión Especial'){
                                                porcentaje = String(aserta["%Incentivo Prod-Renov"])
                                                comision_I = String(aserta["Incentivo Prod-Renov"])
                                            }
                                            //Prima Neta
                                            if(aserta["Prima Neta"] != ArreySICAS[i]["PrimaNeta"]){
                                               //var diferencia= Math.round((aserta["Prima Neta"] - ArreySICAS[i]["PrimaNeta"] )*100)/100;
                                                                errores =  errores + "- Prima Neta\n"
                                                                //diferencias = String (diferencia) + "\n"
                                            }
                                            //TC = MXN
                                            if (TC ==1){
                                                //Comisión Base
                                            if(ArreySICAS[i]["Importe"] != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                                diferencias = diferencias + String(diferenciaCB) + "\n"
                                                errores = errores + "- Comisión \n"
                                            }
                                            else if(ArreySICAS[i]["Importe"] == aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                                diferencias = diferencias + String(diferenciaCB) + "\n"
                                            }
                                            if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                //var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                //diferencias = diferencias + String(diferenciaPC) + "\n"
                                                errores = errores + "- % Comisión \n"
                                            }
                                            //Comisión de Derechos
                                            if(ArreySICAS[i]["Importe"] != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                                diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                errores = errores + "- Comisión Gtos. Exp.\n"
                                            }
                                            else if(ArreySICAS[i]["Importe"] == aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                                diferencias = diferencias +  String(diferenciaCG) + "\n"
                                            }
                                            if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                //var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                //diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                errores = errores + "- % Comisión Maquila\n"
                                            }
                                             //Comisión Especial
                                             if(ArreySICAS[i]["Importe"] != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                                diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                errores = errores + "- Incentivo Prod-Renov\n"
                                            }
                                            else if(ArreySICAS[i]["Importe"] == aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                                diferencias = diferencias +  String(diferenciaIP) + "\n"
                                            }
                                            if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                //var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                //diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                errores = errores + "- %Incentivo Prod-Renov\n"
                                            }
                                            }
                                            //TC diferente a 1
                                            else if(TC!=1){
                                                //Comisión Base
                                                if(importe_mxn != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                    var diferenciaCB= Math.round((aserta["Comisión"] - importe_mxn)*100)/100;
                                                    diferenciaCB=Math.round(diferenciaCB/Number(TC)*100)/100
                                                    diferencias = diferencias + String(diferenciaCB) + "\n"
                                                    errores = errores + "- Comisión \n"
                                                    if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                        errores = errores + "- Tipo de Cambio\n"
                                                    }
                                                }
                                                else if(ArreySICAS[i]["Importe"] == aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                    var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                                    diferenciaCB=Math.round(diferenciaCB/Number(TC)*100)/100
                                                    diferencias = diferencias + String(diferenciaCB) + "\n"
                                                }
                                                if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                    //var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                    //diferencias = diferencias + String(diferenciaPC) + "\n"
                                                    errores = errores + "- % Comisión \n"
                                                }
                                                //Comisión de Derechos
                                                if(importe_mxn != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                    var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - importe_mxn)*100)/100;
                                                    diferenciaCG=Math.round(diferenciaCB/Number(TC)*100)/100
                                                    diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                    errores = errores + "- Comisión Gtos. Exp.\n"
                                                    if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                        errores = errores + "- Tipo de Cambio\n"
                                                    }
                                                }
                                                else if(ArreySICAS[i]["Importe"] == aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                    var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                                    diferenciaCG=Math.round(diferenciaCB/Number(TC)*100)/100
                                                    diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                }
                                                if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                    //var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                    //diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                    errores = errores + "- % Comisión Maquila\n"
                                                }
                                                 //Comisión Especial
                                                 if(importe_mxn != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                    var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] -importe_mxn)*100)/100;
                                                    diferenciaIP=Math.round(diferenciaCB/Number(TC)*100)/100
                                                    diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                    errores = errores + "- Incentivo Prod-Renov\n"
                                                    if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                        errores = errores + "- Tipo de Cambio\n"
                                                    }
                                                }
                                                else if(ArreySICAS[i]["Importe"] == aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                    var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                                    diferenciaIP=Math.round(diferenciaCB/Number(TC)*100)/100
                                                    diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                }
                                                if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                    //var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                    //diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                    errores = errores + "- %Incentivo Prod-Renov\n"
                                                }                                              
                                                
                                            }
                                            
                                                var tabla_sicas = ArreySICAS[i]['Nombre Asegurado o Fiado']+"\t<td>"+ArreySICAS[i]['Poliza']+"\t</td><td>"+ArreySICAS[i]['Endoso']+"\t</td><td>"+ArreySICAS[i]['Moneda']+"\t</td><td>'"+ArreySICAS[i]['Serie']+"'\t</td><td>"+ArreySICAS[i]['TC']+"\t</td><td>"+ArreySICAS[i]['PrimaNeta']+"\t</td><td>"+ArreySICAS[i]['Tipo Comision']+"\t</td><td>"+ArreySICAS[i]['Importe']+"\t</td><td>"+ArreySICAS[i]['% Participacion']+"\t</td>"
                                                var tabla_EC = "<td>"+aserta['Fiado/Contratante']+"</td><td>"+aserta['No Fianza/Certificado']+"</td><td>"+aserta['Folio Factura']+"<td></td><td>"+aserta['Moneda']+"</td><td></td><td>"+porcentaje+"</td><td>"+comision_I+"</td></td><td>"+TC_EstadoCuenta+"</td>"+ "</td><td style='color:var(--b0s-rojo1)'>"+diferencias+"</td>"
                                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+ "<td></td>"+tabla_EC +"<td>"+errores+"</td></tr>"  
                                        }
                                      }
                                      if (encontrar==0){
                                        console.log(key)
                                        var tabla_EC = "<td>"+aserta['Fiado/Contratante']+"</td><td>"+aserta['No Fianza/Certificado']+"</td><td>"+aserta['Folio Factura']+"</td><td></td><td>"+aserta['Moneda']+"</td><td></td><td>"+aserta['% de Comisión']+"</td><td>"+aserta['Comisión']+"</td><td>"+TC_EstadoCuenta+"</td>"+ "</td><td style='color:var(--b0s-rojo1)'></td><td>NO SE ENCONTRÓ</td>"
                                        var tabla_sicas ="Invalido<td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td>"
                                        tablaNA=tablaNA+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+"</td><td>--</td>"+tabla_EC+"</tr>"  
                                      }
                                      
                                        encontrar=0;
                                    }


                                    for(var j=0; j<objetoAserta.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                        var poliza =  String (objetoAserta[j]["No Fianza/Certificado"])
                                        aserta=objetoAserta[j];
                                        if(poliza.length> 11 && poliza.length < 16 && poliza.lastIndexOf('-') != -1){
                                            reng_EC = reng_EC + 1;
                                            resultObject = search(poliza, objetoSICAS)
                                        }
                                        console.log(reng_EC)
                                        document.getElementById("jsondata").innerHTML = "Renglones Estado de Cuenta: "+reng_EC+"\nRenglones SICAS: "+reng_SICAS+"\n";
                                     document.getElementById("jsondata1").innerHTML = tabla+tablaNA+"</table>"; //Se manda la tabla pero no se va a ver porque tiene HIDDEN

                                     
                                    }
                                    ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente.
                                    if(resultObject==0){
                                        document.getElementById("jsondata1").innerHTML = "No se encontró ninguna fianza";
                
                                    }
                                            }
                                );
                             
                            }
                        } else{
                             document.getElementById("jsondata").innerHTML = "No se adjuntó nada en SICAS";
                        }
                       
                        }
                    }else{
                        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Aserta";
                    } 
             
    });
    
    function ExportToExcel(type, fn, dl) {// función que convierte a excel
        var elt = document.getElementById('Aserta');
        var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
        return dl ?
          XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
          XLSX.writeFile(wb, fn || ('Incidencias Aserta.' + (type || 'xlsx')));
     }