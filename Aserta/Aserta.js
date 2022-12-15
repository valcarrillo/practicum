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
    var reng_SICAS=0;
    var reng_EC=0; 
    let tabla ="<table id='Aserta' width='90%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' HIDDEN ><tr><td colspan='10'>SICAS</td><td>--</td><td colspan ='11'>ESTADOS DE CUENTA</td></tr><tr><th>Nombre Asegurado o Fiado</th><th>Póliza</th><th>Endoso</th><th>Moneda</th><th>Serie Recibo</th><th>Tipo Cambio</th><th>Prima Neta</th><th>Tipo Comisión</th><th>Importe</th><th>% Participación</th><th>--</th><th>Nombre Asegurado o Fiado</th><th>Póliza</th><th>Folio Factura</th><th>Endoso</th><th>Moneda</th><th>Serie Recibo</th><th>% Comisión</th><th>Comisión</th><th>Tipo Cambio</th><th>Diferencia % Comisión</th><th>Diferencia Comisión</th><th>Incidencia</th></tr>";
    let tablaNA ="";
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
                    //La tabla tiene atributo HIDDEN para que no se vea, pero ahí está.
                                  
                    let resultObject;//se guardan los resultados de hacer el search. Solo sirve para comprobar la operación. 
                    let aserta;//En esta variable solo se guarda un objeto de Berkley, es decir, un renglón, que es aquel que se va a buscar entre los renglones de SICAS en la función search()
                    var encontrar = 0;   
                          
  
                    //Encontrar un valor ahí adentro
                    search = (key, ArreySICAS, factura) => {
                      //Recorremos arreglos de SICAS
                        for (let i=0; i < ArreySICAS.length; i++) {
                          var polizaSicas = String(ArreySICAS[i].Poliza)
                          var fechas = String(ArreySICAS[i]["Fecha Pago Recibo"]) 
                            const [dia, mes, anio] = fechas.split('/');
                            const fecha1 = new Date(+anio, +mes - 1, +dia);
                            if(fecha1>fechamax){
                                fechamax=fecha1;
                            }
                            if(fecha1<fechamin){
                                fechamin=fecha1;
                            }
                          var facturaSicas = String(ArreySICAS[i].FolioRec)
                          var comision = String(ArreySICAS[i]["Tipo Comision"])
                          var errores=""; //Guarda la columna donde se encuentran los errores de la busqueda
                          var diferencias =""; //Guarda la diferecia entre las comisiones del Edo. Cuenta y Sicas
                          var diferenciasP ="";//Guarda la diferecia entre porcentaje de comisiones del Edo. Cuenta y Sicas
                          var porcentaje =""; //Guarda el porcentaje dependiendo de la columna consultada debido al tipo de comisión
                          var comision_I = "";//Guarda la comisión dependiendo de la columna consultada debido al tipo de comisión
                          var importe_mxn ="";//Multiplica la comsión obtenida por el tipo de cambio ( se utiliza cuando el TC != 1)
                          var TC=String(ArreySICAS[i]["TC"]) //Obtiene el tipo de cambio
                          //Obtener TC del Estado de Cuenta
                          if(aserta['Moneda'] != 'MXN'){
                              var TC_EstadoCuenta= Number(aserta["Comisión"])/ (Number(aserta["% de Comisión"])*Number(aserta["Prima Neta"]))
                              TC_EstadoCuenta= Math.trunc(TC_EstadoCuenta*1000000)/10000
                          }
                          else{
                              var TC_EstadoCuenta= 1
                          }
                            if (polizaSicas && polizaSicas.lastIndexOf('-') != -1 && facturaSicas && polizaSicas.lastIndexOf('-') != -1){
                                reng_SICAS = ArreySICAS.length
                            }
                            //Obtener importe mxn
                            importe_mxn= (ArreySICAS[i]["Importe"]*ArreySICAS[i]["TC"])
                            if (polizaSicas == key && factura == facturaSicas ) { 
                                encontrar++; // Se utiliza como indicador de que si se encontró el renglon buscado
                              //Prima Neta
                              if(aserta["Prima Neta"] != ArreySICAS[i]["PrimaNeta"]){
                                 //var diferencia= Math.round((aserta["Prima Neta"] - ArreySICAS[i]["PrimaNeta"] )*100)/100;
                                                  errores =  errores + "- Prima Neta\n"
                                                  //diferencias = String (diferencia) + "\n"
                              }
                              //TC = MXN
                              if (TC ==1){
                                  // Comparación de comisiones base diferentes, se obtiene la diferencia entre ellas y se registra el error
                              if(ArreySICAS[i]["Importe"] != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                  var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                  diferencias = diferencias + String(diferenciaCB) + "\n"
                                  errores = errores + "- Comisión \n"
                              }
                              // Comparación de comisiones base iguales, se obtiene la diferencia entre ellas 
                              else if(ArreySICAS[i]["Importe"] == aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                  var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                  diferencias = diferencias + String(diferenciaCB) + "\n"
                              }
                              // Comparación de % de comisiones base diferentes, se obtiene la diferencia entre ellas y se registra el error
                              if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                  var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                  diferenciasP = diferenciasP + String(diferenciaPC) + "\n"
                                  errores = errores + "- % Comisión \n"
                              }
                              // Comparación de % de comisiones base iguales, se obtiene la diferencia entre ellas 
                              else if (aserta["% de Comisión"] == ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                diferenciasP = diferenciasP + String(diferenciaPC) + "\n"
                            }
                              // Comparación de comisiones de derechos diferentes, se obtiene la diferencia entre ellas y se registra el error
                              if(ArreySICAS[i]["Importe"] != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                  var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                  diferencias = diferencias +  String(diferenciaCG) + "\n"
                                  errores = errores + "- Comisión Gtos. Exp.\n"
                              }
                              // Comparación de comisiones de derechos iguales, se obtiene la diferencia entre ellas 
                              else if(ArreySICAS[i]["Importe"] == aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                  var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                  diferencias = diferencias +  String(diferenciaCG) + "\n"
                              }
                              // Comparación de % de comisiones de derechos diferentes, se obtiene la diferencia entre ellas y se registra el error
                              if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                  var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                  diferenciasP = diferenciasP +  String(diferenciaPM) + "\n"
                                  errores = errores + "- % Comisión Maquila\n"
                              }
                              // Comparación de % de comisiones de derechos iguales, se obtiene la diferencia entre ellas 
                              else if (aserta["% Maquila"] == ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                diferenciasP = diferenciasP +  String(diferenciaPM) + "\n"
                            }                               
                               // Comparación de comisiones Especiales diferentes, se obtiene la diferencia entre ellas y se registra el error
                               if(ArreySICAS[i]["Importe"] != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                  var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                  diferencias = diferencias +  String(diferenciaIP) + "\n"
                                  errores = errores + "- Incentivo Prod-Renov\n"
                              }
                              // Comparación de comisiones Especiales iguales, se obtiene la diferencia entre ellas 
                              else if(ArreySICAS[i]["Importe"] == aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                  var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                  diferencias = diferencias +  String(diferenciaIP) + "\n"
                              }
                              // Comparación de % de comisiones Especiales  diferentes, se obtiene la diferencia entre ellas y se registra el error
                              if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                  var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                  diferenciasP = diferenciasP +  String(diferenciaPI) + "\n"
                                  errores = errores + "- %Incentivo Prod-Renov\n"
                              }
                              // Comparación de % de comisiones Especiales  iguales, se obtiene la diferencia entre ellas 
                              else if (aserta["%Incentivo Prod-Renov"] == ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                diferenciasP = diferenciasP +  String(diferenciaPI) + "\n"
                            }
                              }
                              //TC diferente a 1
                              else if(TC!=1){
                                  // Comparación de comisiones Base  diferentes, se obtiene la diferencia entre ellas y se registra el error
                                  if(importe_mxn != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                      var diferenciaCB= Math.round((aserta["Comisión"] - importe_mxn)*100)/100;
                                      diferenciaCB=Math.round(diferenciaCB/Number(TC)*100)/100
                                      diferencias = diferencias + String(diferenciaCB) + "\n"
                                      errores = errores + "- Comisión \n"
                                      //Comparación de tipode cambio, si existe, se registra el error
                                      if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                          errores = errores + "- Tipo de Cambio\n"
                                      }
                                  }
                                  // Comparación de comisiones Base  iguales, se obtiene la diferencia entre ellas 
                                  else if(ArreySICAS[i]["Importe"] == aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                      var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                      diferenciaCB=Math.round(diferenciaCB/Number(TC)*100)/100
                                      diferencias = diferencias + String(diferenciaCB) + "\n"
                                  }
                                  // Comparación de % de comisiones Base  diferentes, se obtiene la diferencia entre ellas y se registra el error
                                  if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                      var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                      diferenciasP = diferenciasP + String(diferenciaPC) + "\n"
                                      errores = errores + "- % Comisión \n"
                                  }
                                  // Comparación de % de comisiones Base  iguales, se obtiene la diferencia entre ellas 
                                  else if (aserta["% de Comisión"] == ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                    var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                    diferenciasP = diferenciasP + String(diferenciaPC) + "\n"
                                }
                                  // Comparación de comisiones de derechos diferentes, se obtiene la diferencia entre ellas y se registra el error
                                  if(importe_mxn != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                      var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - importe_mxn)*100)/100;
                                      diferenciaCG=Math.round(diferenciaCB/Number(TC)*100)/100
                                      diferencias = diferencias +  String(diferenciaCG) + "\n"
                                      errores = errores + "- Comisión Gtos. Exp.\n"
                                      if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                          errores = errores + "- Tipo de Cambio\n"
                                      }
                                  }
                                  // Comparación de comisiones de derechosiguales, se obtiene la diferencia entre ellas 
                                  else if(ArreySICAS[i]["Importe"] == aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                      var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                      diferenciaCG=Math.round(diferenciaCB/Number(TC)*100)/100
                                      diferencias = diferencias +  String(diferenciaCG) + "\n"
                                  }
                                  // Comparación de % de comisiones de derechos diferentes, se obtiene la diferencia entre ellas y se registra el error
                                  if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                      var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                      diferenciasP = diferenciasP +  String(diferenciaPM) + "\n"
                                      errores = errores + "- % Comisión Maquila\n"
                                  } 
                                  // Comparación de % de comisiones de derechos iguales, se obtiene la diferencia entre ellas 
                                  else if (aserta["% Maquila"] == ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                    var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                    diferenciasP = diferenciasP +  String(diferenciaPM) + "\n"
                                    errores = errores + "- % Comisión Maquila\n"
                                }
                                   // Comparación de comisiones Especiales diferentes, se obtiene la diferencia entre ellas y se registra el error
                                   if(importe_mxn != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                      var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] -importe_mxn)*100)/100;
                                      diferenciaIP=Math.round(diferenciaCB/Number(TC)*100)/100
                                      diferencias = diferencias +  String(diferenciaIP) + "\n"
                                      errores = errores + "- Incentivo Prod-Renov\n"
                                      //Se comparan ambos Tipos de cambio, en caso de ser diferentes se registra el error.
                                      if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                          errores = errores + "- Tipo de Cambio\n"
                                      }
                                  }
                                  // Comparación de comisiones Especiales iguales, se obtiene la diferencia entre ellas
                                  else if(ArreySICAS[i]["Importe"] == aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                      var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                      diferenciaIP=Math.round(diferenciaCB/Number(TC)*100)/100
                                      diferencias = diferencias +  String(diferenciaIP) + "\n"
                                  }
                                  // Comparación de % de comisiones Especiales  diferentes, se obtiene la diferencia entre ellas y se registra el error
                                  if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                      var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                      diferenciasP = diferenciasP +  String(diferenciaPI) + "\n"
                                      errores = errores + "- %Incentivo Prod-Renov\n"
                                  }   
                                  // Comparación de % de comisiones Especiales  iguales, se obtiene la diferencia entre ellas 
                                  else if (aserta["%Incentivo Prod-Renov"] == ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                    var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                    diferenciasP = diferenciasP +  String(diferenciaPI) + "\n"
                                    errores = errores + "- %Incentivo Prod-Renov\n"
                                }                                            
                                  
                              }
                                console.log("Fecha max: "+fechamax+"\nFecha min:"+fechamin)
                                //datos del renglon de sicas que se esté comparando
                                  var tabla_sicas = ArreySICAS[i]['Nombre Asegurado o Fiado']+"\t<td>"+ArreySICAS[i]['Poliza']+"\t</td><td>"+ArreySICAS[i]['Endoso']+"\t</td><td>"+ArreySICAS[i]['Moneda']+"\t</td><td>'"+ArreySICAS[i]['Serie']+"'\t</td><td>"+ArreySICAS[i]['TC']+"\t</td><td>"+ArreySICAS[i]['PrimaNeta']+"\t</td><td>"+ArreySICAS[i]['Tipo Comision']+"\t</td><td>"+ArreySICAS[i]['Importe']+"\t</td><td>"+ArreySICAS[i]['% Participacion']+"\t</td>"
                                  //datos del renglo del Estado de Cuenta que se esté comparando
                                  var tabla_EC = "<td>"+aserta['Fiado/Contratante']+"</td><td>"+aserta['No Fianza/Certificado']+"</td><td>"+aserta['Folio Factura']+"<td></td><td>"+aserta['Moneda']+"</td><td></td><td>"+porcentaje+"</td><td>"+comision_I+"</td></td><td>"+TC_EstadoCuenta+"</td>"+ "<td>"+diferenciasP+"</td><td style='color:var(--b0s-rojo1)'>"+diferencias+"</td>"
                                  tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+ "<td></td>"+tabla_EC +"<td>"+errores+"</td></tr>"  
                          }
                        }
                        if (encontrar==0){
                          
                          var tabla_EC = "<td>"+aserta['Fiado/Contratante']+"</td><td>"+aserta['No Fianza/Certificado']+"</td><td>"+aserta['Folio Factura']+"</td><td></td><td>"+aserta['Moneda']+"</td><td></td><td>"+aserta['% de Comisión']+"</td><td>"+aserta['Comisión']+"</td><td>"+TC_EstadoCuenta+"</td>"+ "</td><td style='color:var(--b0s-rojo1)'></td><td></td><td>NO SE ENCONTRÓ</td>"
                          
                          var tabla_sicas ="Invalido<td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td>"
                          //Unen ambas tablas, formando una única tabla
                          tablaNA=tablaNA+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+"</td><td>--</td>"+tabla_EC+"</tr>"  
                        }
                          encontrar=0;
                    }
                    
                      for(var j=0; j<objetoAserta.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                          var poliza =  String (objetoAserta[j]["No Fianza/Certificado"])
                          var factura =  String (objetoAserta[j]["Folio Factura"]);
                          aserta=objetoAserta[j];
                          if(poliza && poliza.lastIndexOf('-') != -1 && factura){
                              reng_EC = reng_EC + 1;
                              resultObject = search(poliza, objetoSICAS,factura)
                          }                       
                      }
                      if(reng_EC == 0 || reng_SICAS == 0){
                        if (reng_EC == 0){
                            document.getElementById("jsondata").innerHTML = "Estado(s) de cuenta erroneo";
                          }
                        if (reng_SICAS == 0){
                            document.getElementById("jsondata").innerHTML = "Base de Datos Sicas erroneo";
                          }
                      }
                      if (reng_EC != 0 && reng_SICAS != 0){
                       document.getElementById("jsondata1").innerHTML = tabla+tablaNA+"</table>"; //Se manda la tabla pero no se va a ver porque tiene HIDDEN
                        ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente.
                        //document.getElementById("jsondata").innerHTML = "Renglones Estado de Cuenta: "+reng_EC+"\nRenglones SICAS: "+reng_SICAS+"\n";
                        document.getElementById("jsondata").innerHTML = "Renglones Estado de Cuenta: "+reng_EC+"\nRenglones SICAS: "+reng_SICAS+"\n";
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
    var elt = document.getElementById('Aserta');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    var nombre ='CONCILIACIÓN ASERTA DEL '+fechamin.getDate()+" "+month[+fechamin.getMonth()+1]+" "+fechamin.getFullYear()+" AL "+fechamax.getDate()+" "+month[+fechamax.getMonth()+1]+" "+fechamax.getFullYear()+".";
    return dl ?
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || (nombre + (type || 'xlsx')));
 }