//https://levelup.gitconnected.com/how-to-convert-excel-file-into-json-object-by-using-javascript-9e95532d47c5
//convertir en EXCEL https://codepedia.info/javascript-export-html-table-data-to-excel
let selectedFile;
let selectedFile1_2;
let selectedFile1_3;
let selectedFile1_4;
let selectedFile1_5;
let selectedFile2;
let num_ECT;
console.log(window.XLSX);
//EC 1
document.getElementById('inputA').addEventListener("change", (event) => { //Lee estado de cuenta
    selectedFile = event.target.files[0];
})
//EC 2
document.getElementById('inputB').addEventListener("change", (event) => {// Lee SICAS
    selectedFile1_2 = event.target.files[0];
}
)
//EC 3
document.getElementById('inputC').addEventListener("change", (event) => {// Lee SICAS
    selectedFile1_3 = event.target.files[0];
}
)
//EC 4
document.getElementById('inputD').addEventListener("change", (event) => {// Lee SICAS
    selectedFile1_4 = event.target.files[0];
}
)
//EC 5
document.getElementById('inputE').addEventListener("change", (event) => {// Lee SICAS
    selectedFile1_5 = event.target.files[0];
}
)
//BD SICAS
document.getElementById('inputSicas').addEventListener("change", (event) => {// Lee SICAS
    selectedFile2 = event.target.files[0];
}
)

let objetoSICAS; //Array de objetos en el que se va a guarar SICAS
let objetoAserta; //Array de objetos en el que se va a guardar Berkley
let objetoAserta2;
let objetoAserta3;
let objetoAserta4;
let objetoAserta5;
let objetoFinal;
let encabezado;

document.getElementById('button').addEventListener("click", () => {
    var num, num2;
    num_ECT = document.getElementById("select").value
    //ESTRUCTURA TABLA
    let tabla ="<table id='Aserta' width='90%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' HIDDEN><tr><th>Nombre Asegurado o Fiado</th><th>Póliza</th><th>Endoso</th><th>Moneda</th><th>Serie Recibo</th><th>Tipo Cambio</th><th>Prima Neta</th><th>Tipo Comisión</th><th>Importe</th><th>% Participación</th><th>--</th><th>Nombre Asegurado o Fiado</th><th>Póliza</th><th>Endoso</th><th>Moneda</th><th>Serie Recibo</th><th>% Comisión</th><th>Comisión</th><th>Tipo Cambio</th><th>Diferencia</th><th>Incidencia</th></tr>";
            switch (num_ECT){
                case '0':
                    document.getElementById("jsondata").innerHTML = "No se selecciono el número de Estado de Cuenta";
                    break;
                case '1':
                    if(selectedFile ){ //Función para convertir Edo de Cuenta en array de objetos
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
                                  var encontrar;           
                
                                  //Encontrar un valor ahí adentro
                                  search = (key, ArreySICAS) => {
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
                                            TC_EstadoCuenta= Math.trunc(TC_EstadoCuenta*10000)/10000
                                        }
                                        else{
                                            var TC_EstadoCuenta= 1
                                        }
                                          
                                          //Obtener importe mxn
                                          importe_mxn= (ArreySICAS[i]["Importe"]*ArreySICAS[i]["TC"])
                                          if (polizaSicas == key ) {
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
                                            encontrar++;
                                            //Prima Neta
                                            if(aserta["Prima Neta"] != ArreySICAS[i]["PrimaNeta"]){
                                                var diferencia= Math.round((aserta["Prima Neta"] - ArreySICAS[i]["PrimaNeta"] )*100)/100;
                                                errores =  errores + "- Prima Neta\n"
                                                diferencias = String (diferencia) + "\n"
                                            }
                                            //TC = MXN
                                            if (TC ==1){
                                                //Comisión Base
                                            if(ArreySICAS[i]["Importe"] != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                                diferencias = diferencias + String(diferenciaCB) + "\n"
                                                errores = errores + "- Comisión \n"
                                            }
                                            if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                diferencias = diferencias + String(diferenciaPC) + "\n"
                                                errores = errores + "- % Comisión \n"
                                            }
                                            //Comisión de Derechos
                                            if(ArreySICAS[i]["Importe"] != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                                diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                errores = errores + "- Comisión Gtos. Exp.\n"
                                            }
                                            if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                errores = errores + "- % Comisión Maquila\n"
                                            }
                                             //Comisión Especial
                                             if(ArreySICAS[i]["Importe"] != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                                diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                errores = errores + "- Incentivo Prod-Renov\n"
                                            }
                                            if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                errores = errores + "- %Incentivo Prod-Renov\n"
                                            }
                                            }
                                            //TC diferente a 1
                                            else if(TC!=1){
                                                //Comisión Base
                                                if(importe_mxn != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                    var diferenciaCB= Math.round((aserta["Comisión"] - importe_mxn)*100)/100;
                                                    diferencias = diferencias + String(diferenciaCB) + "\n"
                                                    errores = errores + "- Comisión \n"
                                                    if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                        errores = errores + "- Tipo de Cambio\n"
                                                    }
                                                }
                                                if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                    var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                    diferencias = diferencias + String(diferenciaPC) + "\n"
                                                    errores = errores + "- % Comisión \n"
                                                }
                                                //Comisión de Derechos
                                                if(importe_mxn != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                    var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - importe_mxn)*100)/100;
                                                    diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                    errores = errores + "- Comisión Gtos. Exp.\n"
                                                    if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                        errores = errores + "- Tipo de Cambio\n"
                                                    }
                                                }
                                                if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                    var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                    diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                    errores = errores + "- % Comisión Maquila\n"
                                                }
                                                 //Comisión Especial
                                                 if(importe_mxn != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                    var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] -importe_mxn)*100)/100;
                                                    diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                    errores = errores + "- Incentivo Prod-Renov\n"
                                                    if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                        errores = errores + "- Tipo de Cambio\n"
                                                    }
                                                }
                                                if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                    var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                    diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                    errores = errores + "- %Incentivo Prod-Renov\n"
                                                }                                              
                                                
                                            }
                                            
                                            if(errores != ""){
                                               var tabla_sicas = ArreySICAS[i]['Nombre Asegurado o Fiado']+"\t<td>"+ArreySICAS[i]['Poliza']+"\t</td><td>"+ArreySICAS[i]['Endoso']+"\t</td><td>"+ArreySICAS[i]['Moneda']+"\t</td><td>"+ArreySICAS[i]['Serie']+"\t</td><td>"+ArreySICAS[i]['TC']+"\t</td><td>"+ArreySICAS[i]['PrimaNeta']+"\t</td><td>"+ArreySICAS[i]['Tipo Comision']+"\t</td><td>"+ArreySICAS[i]['Importe']+"\t</td><td>"+ArreySICAS[i]['% Participacion']+"\t</td>"
                                               var tabla_EC = "<td>"+aserta['Fiado/Contratante']+"</td><td>"+aserta['No Fianza/']+"</td><td></td><td>"+aserta['Moneda']+"</td><td></td><td>"+porcentaje+"</td><td>"+comision_I+"</td></td><td>"+TC_EstadoCuenta+"</td>"+ "</td><td style='color:var(--b0s-rojo1)'>"+diferencias+"</td>"
                                               tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+ "<td></td>"+tabla_EC +"<td>"+errores+"</td></tr>"
                                            }
                                            
                                        }
                                      }
                                      if(encontrar==0){
                                        //--tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+aserta["Prima Neta"]+"</td><td>"+aserta["% de Comisión"]+"</td><td>"+comision+"</td><td>"+aserta["Comisión"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                                        }
                                        encontrar=0;
                                    }


                                    for(var j=0; j<objetoAserta.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                        var poliza =  String (objetoAserta[j]["No Fianza/"])
                                        aserta=objetoAserta[j];
                                        if(poliza.length == 12){
                                            resultObject = search(poliza, objetoSICAS)
                                        }
                                      
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
                        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Aserta";
                    } 
                break;
                case '2':
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
                         if(selectedFile1_2){ //Función para convertir Edo de Cuenta en array de objetos
                            let fileReader = new FileReader();
                        fileReader.readAsBinaryString(selectedFile1_2);
                        fileReader.onload = (event)=>{
                         let data1_2 = event.target.result;
                         let workbook1_2 = XLSX.read(data1_2,{type:"binary"});       
                         workbook1_2.SheetNames.forEach(sheet => {
                            //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                objetoAserta2 = XLSX.utils.sheet_to_row_object_array(workbook1_2.Sheets[sheet], {range:1}); //Nombre del array
                             });
                             for (var t=0;t<objetoAserta2.length;t++){
                                objetoAserta.push(objetoAserta2[t])
                             }
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
                                      var encontrar;           
                    
                                      //Encontrar un valor ahí adentro
                                      search = (key, ArreySICAS) => {
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
                                                TC_EstadoCuenta= Math.trunc(TC_EstadoCuenta*10000)/10000
                                            }
                                            else{
                                                var TC_EstadoCuenta= 1
                                            }
                                              
                                              //Obtener importe mxn
                                              importe_mxn= (ArreySICAS[i]["Importe"]*ArreySICAS[i]["TC"])
                                              if (polizaSicas == key ) {
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
                                                encontrar++;
                                                //Prima Neta
                                                if(aserta["Prima Neta"] != ArreySICAS[i]["PrimaNeta"]){
                                                    var diferencia= Math.round((aserta["Prima Neta"] - ArreySICAS[i]["PrimaNeta"] )*100)/100;
                                                    errores =  errores + "- Prima Neta\n"
                                                    diferencias = String (diferencia) + "\n"
                                                }
                                                //TC = MXN
                                                if (TC ==1){
                                                    //Comisión Base
                                                if(ArreySICAS[i]["Importe"] != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                    var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                                    diferencias = diferencias + String(diferenciaCB) + "\n"
                                                    errores = errores + "- Comisión \n"
                                                }
                                                if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                    var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                    diferencias = diferencias + String(diferenciaPC) + "\n"
                                                    errores = errores + "- % Comisión \n"
                                                }
                                                //Comisión de Derechos
                                                if(ArreySICAS[i]["Importe"] != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                    var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                                    diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                    errores = errores + "- Comisión Gtos. Exp.\n"
                                                }
                                                if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                    var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                    diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                    errores = errores + "- % Comisión Maquila\n"
                                                }
                                                 //Comisión Especial
                                                 if(ArreySICAS[i]["Importe"] != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                    var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                                    diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                    errores = errores + "- Incentivo Prod-Renov\n"
                                                }
                                                if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                    var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                    diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                    errores = errores + "- %Incentivo Prod-Renov\n"
                                                }
                                                }
                                                //TC diferente a 1
                                                else if(TC!=1){
                                                    //Comisión Base
                                                    if(importe_mxn != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                        var diferenciaCB= Math.round((aserta["Comisión"] - importe_mxn)*100)/100;
                                                        diferencias = diferencias + String(diferenciaCB) + "\n"
                                                        errores = errores + "- Comisión \n"
                                                        if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                            errores = errores + "- Tipo de Cambio\n"
                                                        }
                                                    }
                                                    if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                        var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                        diferencias = diferencias + String(diferenciaPC) + "\n"
                                                        errores = errores + "- % Comisión \n"
                                                    }
                                                    //Comisión de Derechos
                                                    if(importe_mxn != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                        var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - importe_mxn)*100)/100;
                                                        diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                        errores = errores + "- Comisión Gtos. Exp.\n"
                                                        if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                            errores = errores + "- Tipo de Cambio\n"
                                                        }
                                                    }
                                                    if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                        var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                        diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                        errores = errores + "- % Comisión Maquila\n"
                                                    }
                                                     //Comisión Especial
                                                     if(importe_mxn != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                        var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] -importe_mxn)*100)/100;
                                                        diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                        errores = errores + "- Incentivo Prod-Renov\n"
                                                        if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                            errores = errores + "- Tipo de Cambio\n"
                                                        }
                                                    }
                                                    if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                        var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                        diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                        errores = errores + "- %Incentivo Prod-Renov\n"
                                                    }                                              
                                                    
                                                }
                                                
                                                if(errores != ""){
                                                   var tabla_sicas = ArreySICAS[i]['Nombre Asegurado o Fiado']+"\t<td>"+ArreySICAS[i]['Poliza']+"\t</td><td>"+ArreySICAS[i]['Endoso']+"\t</td><td>"+ArreySICAS[i]['Moneda']+"\t</td><td>"+ArreySICAS[i]['Serie']+"\t</td><td>"+ArreySICAS[i]['TC']+"\t</td><td>"+ArreySICAS[i]['PrimaNeta']+"\t</td><td>"+ArreySICAS[i]['Tipo Comision']+"\t</td><td>"+ArreySICAS[i]['Importe']+"\t</td><td>"+ArreySICAS[i]['% Participacion']+"\t</td>"
                                                   var tabla_EC = "<td>"+aserta['Fiado/Contratante']+"</td><td>"+aserta['No Fianza/']+"</td><td></td><td>"+aserta['Moneda']+"</td><td></td><td>"+porcentaje+"</td><td>"+comision_I+"</td></td><td>"+TC_EstadoCuenta+"</td>"+ "</td><td style='color:var(--b0s-rojo1)'>"+diferencias+"</td>"
                                                   tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+ "<td></td>"+tabla_EC +"<td>"+errores+"</td></tr>"
                                                }
                                                
                                            }
                                          }
                                          if(encontrar==0){
                                            //--tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+aserta["Prima Neta"]+"</td><td>"+aserta["% de Comisión"]+"</td><td>"+comision+"</td><td>"+aserta["Comisión"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                                            }
                                            encontrar=0;
                                        }
    
    
                                        for(var j=0; j<objetoAserta.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                            var poliza =  String (objetoAserta[j]["No Fianza/"])
                                            aserta=objetoAserta[j];
                                            if(poliza.length == 12){
                                                resultObject = search(poliza, objetoSICAS)
                                            }
                                          
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
                            document.getElementById("jsondata").innerHTML = "No se adjuntó el 2° Estado de Cuenta de Aserta";
                        }
                         
                        }
                    }else{
                        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Aserta";
                    } 
                break;
                case '3':
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
                         if(selectedFile1_2){ //Función para convertir Edo de Cuenta en array de objetos
                            let fileReader = new FileReader();
                        fileReader.readAsBinaryString(selectedFile1_2);
                        fileReader.onload = (event)=>{
                         let data1_2 = event.target.result;
                         let workbook1_2 = XLSX.read(data1_2,{type:"binary"});       
                         workbook1_2.SheetNames.forEach(sheet => {
                            //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                objetoAserta2 = XLSX.utils.sheet_to_row_object_array(workbook1_2.Sheets[sheet], {range:1}); //Nombre del array
                             });
                             for (var t=0;t<objetoAserta2.length;t++){
                                objetoAserta.push(objetoAserta2[t])
                             }
                             if(selectedFile1_3){ //Función para convertir Edo de Cuenta en array de objetos
                                let fileReader = new FileReader();
                                fileReader.readAsBinaryString(selectedFile1_3);
                                fileReader.onload = (event)=>{
                                 let data1_3 = event.target.result;
                                 let workbook1_3 = XLSX.read(data1_3,{type:"binary"});
                                 console.log(workbook1_3);        
                                 workbook1_3.SheetNames.forEach(sheet => {
                                    //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                        objetoAserta3 = XLSX.utils.sheet_to_row_object_array(workbook1_3.Sheets[sheet], {range:1}); //Nombre del array
                                     console.log("ASERTA3")
                                        console.log(objetoAserta3);
                                 });
                                 for (var t=0;t<objetoAserta3.length;t++){
                                    objetoAserta.push(objetoAserta3[t])
                                 }
                                 console.log(objetoAserta)
                                 if(selectedFile2){ //Función que convierte SICAS en array de objetos
                                    let fileReader = new FileReader();
                                    fileReader.readAsBinaryString(selectedFile2);
                                    fileReader.onload = (event)=>{
                                     let data2 = event.target.result;
                                     let workbook2 = XLSX.read(data2,{type:"binary"});
                                     workbook2.SheetNames.forEach(sheet => {
                                          objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //Nombre del array
                        
                                            //La tabla tiene atributo HIDDEN para que no se vea, pero ahí está.
                                          let tabla ="<table id='Aserta' width='90%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' style='text-align:center;'><tr><th>Nombre Asegurado o Fiado</th><th>Póliza</th><th>Endoso</th><th>Moneda</th><th>Serie Recibo</th><th>Tipo Cambio</th><th>Prima Neta</th><th>Tipo Comisión</th><th>Importe</th><th>% Participación</th><th>--</th><th>Nombre Asegurado o Fiado</th><th>Póliza</th><th>Endoso</th><th>Moneda</th><th>Serie Recibo</th><th>% Comisión</th><th>Comisión</th><th>Tipo Cambio</th><th>Diferencia</th><th>Incidencia</th></tr>";
                                          let resultObject;
                                          let aserta;
                                          var encontrar;           
                        
                                          //Encontrar un valor ahí adentro
                                          search = (key, ArreySICAS) => {
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
                                                    TC_EstadoCuenta= Math.trunc(TC_EstadoCuenta*10000)/10000
                                                }
                                                else{
                                                    var TC_EstadoCuenta= 1
                                                }
                                                  
                                                  //Obtener importe mxn
                                                  importe_mxn= (ArreySICAS[i]["Importe"]*ArreySICAS[i]["TC"])
                                                  if (polizaSicas == key ) {
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
                                                    encontrar++;
                                                    //Prima Neta
                                                    if(aserta["Prima Neta"] != ArreySICAS[i]["PrimaNeta"]){
                                                        var diferencia= Math.round((aserta["Prima Neta"] - ArreySICAS[i]["PrimaNeta"] )*100)/100;
                                                        errores =  errores + "- Prima Neta\n"
                                                        diferencias = String (diferencia) + "\n"
                                                    }
                                                    //TC = MXN
                                                    if (TC ==1){
                                                        //Comisión Base
                                                    if(ArreySICAS[i]["Importe"] != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                        var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                                        diferencias = diferencias + String(diferenciaCB) + "\n"
                                                        errores = errores + "- Comisión \n"
                                                    }
                                                    if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                        var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                        diferencias = diferencias + String(diferenciaPC) + "\n"
                                                        errores = errores + "- % Comisión \n"
                                                    }
                                                    //Comisión de Derechos
                                                    if(ArreySICAS[i]["Importe"] != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                        var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                                        diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                        errores = errores + "- Comisión Gtos. Exp.\n"
                                                    }
                                                    if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                        var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                        diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                        errores = errores + "- % Comisión Maquila\n"
                                                    }
                                                     //Comisión Especial
                                                     if(ArreySICAS[i]["Importe"] != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                        var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                                        diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                        errores = errores + "- Incentivo Prod-Renov\n"
                                                    }
                                                    if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                        var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                        diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                        errores = errores + "- %Incentivo Prod-Renov\n"
                                                    }
                                                    }
                                                    //TC diferente a 1
                                                    else if(TC!=1){
                                                        //Comisión Base
                                                        if(importe_mxn != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                            var diferenciaCB= Math.round((aserta["Comisión"] - importe_mxn)*100)/100;
                                                            diferencias = diferencias + String(diferenciaCB) + "\n"
                                                            errores = errores + "- Comisión \n"
                                                            if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                                errores = errores + "- Tipo de Cambio\n"
                                                            }
                                                        }
                                                        if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                            var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                            diferencias = diferencias + String(diferenciaPC) + "\n"
                                                            errores = errores + "- % Comisión \n"
                                                        }
                                                        //Comisión de Derechos
                                                        if(importe_mxn != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                            var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - importe_mxn)*100)/100;
                                                            diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                            errores = errores + "- Comisión Gtos. Exp.\n"
                                                            if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                                errores = errores + "- Tipo de Cambio\n"
                                                            }
                                                        }
                                                        if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                            var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                            diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                            errores = errores + "- % Comisión Maquila\n"
                                                        }
                                                         //Comisión Especial
                                                         if(importe_mxn != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                            var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] -importe_mxn)*100)/100;
                                                            diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                            errores = errores + "- Incentivo Prod-Renov\n"
                                                            if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                                errores = errores + "- Tipo de Cambio\n"
                                                            }
                                                        }
                                                        if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                            var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                            diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                            errores = errores + "- %Incentivo Prod-Renov\n"
                                                        }                                              
                                                        
                                                    }
                                                    
                                                    if(errores != ""){
                                                       var tabla_sicas = ArreySICAS[i]['Nombre Asegurado o Fiado']+"\t<td>"+ArreySICAS[i]['Poliza']+"\t</td><td>"+ArreySICAS[i]['Endoso']+"\t</td><td>"+ArreySICAS[i]['Moneda']+"\t</td><td>"+ArreySICAS[i]['Serie']+"\t</td><td>"+ArreySICAS[i]['TC']+"\t</td><td>"+ArreySICAS[i]['PrimaNeta']+"\t</td><td>"+ArreySICAS[i]['Tipo Comision']+"\t</td><td>"+ArreySICAS[i]['Importe']+"\t</td><td>"+ArreySICAS[i]['% Participacion']+"\t</td>"
                                                       var tabla_EC = "<td>"+aserta['Fiado/Contratante']+"</td><td>"+aserta['No Fianza/']+"</td><td></td><td>"+aserta['Moneda']+"</td><td></td><td>"+porcentaje+"</td><td>"+comision_I+"</td></td><td>"+TC_EstadoCuenta+"</td>"+ "</td><td style='color:var(--b0s-rojo1)'>"+diferencias+"</td>"
                                                       tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+ "<td></td>"+tabla_EC +"<td>"+errores+"</td></tr>"
                                                    }
                                                    
                                                }
                                              }
                                              if(encontrar==0){
                                                //--tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+aserta["Prima Neta"]+"</td><td>"+aserta["% de Comisión"]+"</td><td>"+comision+"</td><td>"+aserta["Comisión"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                                                }
                                                encontrar=0;
                                            }
        
        
                                            for(var j=0; j<objetoAserta.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                                var poliza =  String (objetoAserta[j]["No Fianza/"])
                                                aserta=objetoAserta[j];
                                                if(poliza.length == 12){
                                                    resultObject = search(poliza, objetoSICAS)
                                                }
                                              
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
                                document.getElementById("jsondata").innerHTML = "No se adjuntó el 3° Estado de Cuenta de Aserta";
                            }
                            }
                        }else{
                            document.getElementById("jsondata").innerHTML = "No se adjuntó el 2° Estado de Cuenta de Aserta";
                        }
                         
                        }
                    }else{
                        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Aserta";
                    } 
                break;
                case '4':
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
                         if(selectedFile1_2){ //Función para convertir Edo de Cuenta en array de objetos
                            let fileReader = new FileReader();
                        fileReader.readAsBinaryString(selectedFile1_2);
                        fileReader.onload = (event)=>{
                         let data1_2 = event.target.result;
                         let workbook1_2 = XLSX.read(data1_2,{type:"binary"});       
                         workbook1_2.SheetNames.forEach(sheet => {
                            //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                objetoAserta2 = XLSX.utils.sheet_to_row_object_array(workbook1_2.Sheets[sheet], {range:1}); //Nombre del array
                             });
                             for (var t=0;t<objetoAserta2.length;t++){
                                objetoAserta.push(objetoAserta2[t])
                             }
                             if(selectedFile1_3){ //Función para convertir Edo de Cuenta en array de objetos
                                let fileReader = new FileReader();
                                fileReader.readAsBinaryString(selectedFile1_3);
                                fileReader.onload = (event)=>{
                                 let data1_3 = event.target.result;
                                 let workbook1_3 = XLSX.read(data1_3,{type:"binary"});
                                 console.log(workbook1_3);        
                                 workbook1_3.SheetNames.forEach(sheet => {
                                    //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                        objetoAserta3 = XLSX.utils.sheet_to_row_object_array(workbook1_3.Sheets[sheet], {range:1}); //Nombre del array
                                     console.log("ASERTA3")
                                        console.log(objetoAserta3);
                                 });
                                 for (var t=0;t<objetoAserta3.length;t++){
                                    objetoAserta.push(objetoAserta3[t])
                                 }
                                 console.log(objetoAserta)
                                 if(selectedFile1_4){ //Función para convertir Edo de Cuenta en array de objetos
                                    let fileReader = new FileReader();
                                    fileReader.readAsBinaryString(selectedFile1_4);
                                    fileReader.onload = (event)=>{
                                     let data1_4 = event.target.result;
                                     let workbook1_4 = XLSX.read(data1_4,{type:"binary"});
                                     console.log(workbook1_4);        
                                     workbook1_4.SheetNames.forEach(sheet => {
                                        //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                            objetoAserta4 = XLSX.utils.sheet_to_row_object_array(workbook1_4.Sheets[sheet], {range:1}); //Nombre del array
                                         console.log("ASERTA4")
                                            console.log(objetoAserta4);
                                     });
                                     for (var t=0;t<objetoAserta4.length;t++){
                                        objetoAserta.push(objetoAserta4[t])
                                     }
                                     console.log(objetoAserta)
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
                                              var encontrar;           
                            
                                              //Encontrar un valor ahí adentro
                                              search = (key, ArreySICAS) => {
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
                                                        TC_EstadoCuenta= Math.trunc(TC_EstadoCuenta*10000)/10000
                                                    }
                                                    else{
                                                        var TC_EstadoCuenta= 1
                                                    }
                                                      
                                                      //Obtener importe mxn
                                                      importe_mxn= (ArreySICAS[i]["Importe"]*ArreySICAS[i]["TC"])
                                                      if (polizaSicas == key ) {
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
                                                        encontrar++;
                                                        //Prima Neta
                                                        if(aserta["Prima Neta"] != ArreySICAS[i]["PrimaNeta"]){
                                                            var diferencia= Math.round((aserta["Prima Neta"] - ArreySICAS[i]["PrimaNeta"] )*100)/100;
                                                            errores =  errores + "- Prima Neta\n"
                                                            diferencias = String (diferencia) + "\n"
                                                        }
                                                        //TC = MXN
                                                        if (TC ==1){
                                                            //Comisión Base
                                                        if(ArreySICAS[i]["Importe"] != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                            var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                                            diferencias = diferencias + String(diferenciaCB) + "\n"
                                                            errores = errores + "- Comisión \n"
                                                        }
                                                        if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                            var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                            diferencias = diferencias + String(diferenciaPC) + "\n"
                                                            errores = errores + "- % Comisión \n"
                                                        }
                                                        //Comisión de Derechos
                                                        if(ArreySICAS[i]["Importe"] != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                            var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                                            diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                            errores = errores + "- Comisión Gtos. Exp.\n"
                                                        }
                                                        if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                            var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                            diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                            errores = errores + "- % Comisión Maquila\n"
                                                        }
                                                         //Comisión Especial
                                                         if(ArreySICAS[i]["Importe"] != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                            var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                                            diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                            errores = errores + "- Incentivo Prod-Renov\n"
                                                        }
                                                        if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                            var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                            diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                            errores = errores + "- %Incentivo Prod-Renov\n"
                                                        }
                                                        }
                                                        //TC diferente a 1
                                                        else if(TC!=1){
                                                            //Comisión Base
                                                            if(importe_mxn != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                                var diferenciaCB= Math.round((aserta["Comisión"] - importe_mxn)*100)/100;
                                                                diferencias = diferencias + String(diferenciaCB) + "\n"
                                                                errores = errores + "- Comisión \n"
                                                                if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                                    errores = errores + "- Tipo de Cambio\n"
                                                                }
                                                            }
                                                            if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                                var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                                diferencias = diferencias + String(diferenciaPC) + "\n"
                                                                errores = errores + "- % Comisión \n"
                                                            }
                                                            //Comisión de Derechos
                                                            if(importe_mxn != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                                var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - importe_mxn)*100)/100;
                                                                diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                                errores = errores + "- Comisión Gtos. Exp.\n"
                                                                if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                                    errores = errores + "- Tipo de Cambio\n"
                                                                }
                                                            }
                                                            if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                                var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                                diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                                errores = errores + "- % Comisión Maquila\n"
                                                            }
                                                             //Comisión Especial
                                                             if(importe_mxn != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                                var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] -importe_mxn)*100)/100;
                                                                diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                                errores = errores + "- Incentivo Prod-Renov\n"
                                                                if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                                    errores = errores + "- Tipo de Cambio\n"
                                                                }
                                                            }
                                                            if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                                var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                                diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                                errores = errores + "- %Incentivo Prod-Renov\n"
                                                            }                                              
                                                            
                                                        }
                                                        
                                                        if(errores != ""){
                                                           var tabla_sicas = ArreySICAS[i]['Nombre Asegurado o Fiado']+"\t<td>"+ArreySICAS[i]['Poliza']+"\t</td><td>"+ArreySICAS[i]['Endoso']+"\t</td><td>"+ArreySICAS[i]['Moneda']+"\t</td><td>"+ArreySICAS[i]['Serie']+"\t</td><td>"+ArreySICAS[i]['TC']+"\t</td><td>"+ArreySICAS[i]['PrimaNeta']+"\t</td><td>"+ArreySICAS[i]['Tipo Comision']+"\t</td><td>"+ArreySICAS[i]['Importe']+"\t</td><td>"+ArreySICAS[i]['% Participacion']+"\t</td>"
                                                           var tabla_EC = "<td>"+aserta['Fiado/Contratante']+"</td><td>"+aserta['No Fianza/']+"</td><td></td><td>"+aserta['Moneda']+"</td><td></td><td>"+porcentaje+"</td><td>"+comision_I+"</td></td><td>"+TC_EstadoCuenta+"</td>"+ "</td><td style='color:var(--b0s-rojo1)'>"+diferencias+"</td>"
                                                           tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+ "<td></td>"+tabla_EC +"<td>"+errores+"</td></tr>"
                                                        }
                                                        
                                                    }
                                                  }
                                                  if(encontrar==0){
                                                    //--tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+aserta["Prima Neta"]+"</td><td>"+aserta["% de Comisión"]+"</td><td>"+comision+"</td><td>"+aserta["Comisión"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                                                    }
                                                    encontrar=0;
                                                }
            
            
                                                for(var j=0; j<objetoAserta.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                                    var poliza =  String (objetoAserta[j]["No Fianza/"])
                                                    aserta=objetoAserta[j];
                                                    if(poliza.length == 12){
                                                        resultObject = search(poliza, objetoSICAS)
                                                    }
                                                  
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
                                    document.getElementById("jsondata").innerHTML = "No se adjuntó el 4° Estado de Cuenta de Aserta";
                                }
                                }
                            }else{
                                document.getElementById("jsondata").innerHTML = "No se adjuntó el 3° Estado de Cuenta de Aserta";
                            }
                            }
                        }else{
                            document.getElementById("jsondata").innerHTML = "No se adjuntó el 2° Estado de Cuenta de Aserta";
                        }
                         
                        }
                    }else{
                        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Aserta";
                    } 
                break;
                case '5':
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
                         if(selectedFile1_2){ //Función para convertir Edo de Cuenta en array de objetos
                            let fileReader = new FileReader();
                        fileReader.readAsBinaryString(selectedFile1_2);
                        fileReader.onload = (event)=>{
                         let data1_2 = event.target.result;
                         let workbook1_2 = XLSX.read(data1_2,{type:"binary"});       
                         workbook1_2.SheetNames.forEach(sheet => {
                            //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                objetoAserta2 = XLSX.utils.sheet_to_row_object_array(workbook1_2.Sheets[sheet], {range:1}); //Nombre del array
                             });
                             for (var t=0;t<objetoAserta2.length;t++){
                                objetoAserta.push(objetoAserta2[t])
                             }
                             if(selectedFile1_3){ //Función para convertir Edo de Cuenta en array de objetos
                                let fileReader = new FileReader();
                                fileReader.readAsBinaryString(selectedFile1_3);
                                fileReader.onload = (event)=>{
                                 let data1_3 = event.target.result;
                                 let workbook1_3 = XLSX.read(data1_3,{type:"binary"});
                                 console.log(workbook1_3);        
                                 workbook1_3.SheetNames.forEach(sheet => {
                                    //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                        objetoAserta3 = XLSX.utils.sheet_to_row_object_array(workbook1_3.Sheets[sheet], {range:1}); //Nombre del array
                                     console.log("ASERTA3")
                                        console.log(objetoAserta3);
                                 });
                                 for (var t=0;t<objetoAserta3.length;t++){
                                    objetoAserta.push(objetoAserta3[t])
                                 }
                                 console.log(objetoAserta)
                                 if(selectedFile1_4){ //Función para convertir Edo de Cuenta en array de objetos
                                    let fileReader = new FileReader();
                                    fileReader.readAsBinaryString(selectedFile1_4);
                                    fileReader.onload = (event)=>{
                                     let data1_4 = event.target.result;
                                     let workbook1_4 = XLSX.read(data1_4,{type:"binary"});
                                     console.log(workbook1_4);        
                                     workbook1_4.SheetNames.forEach(sheet => {
                                        //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                            objetoAserta4 = XLSX.utils.sheet_to_row_object_array(workbook1_4.Sheets[sheet], {range:1}); //Nombre del array
                                         console.log("ASERTA4")
                                            console.log(objetoAserta4);
                                     });
                                     for (var t=0;t<objetoAserta4.length;t++){
                                        objetoAserta.push(objetoAserta4[t])
                                     }
                                     console.log(objetoAserta)
                                     if(selectedFile1_5){ //Función para convertir Edo de Cuenta en array de objetos
                                        let fileReader = new FileReader();
                                        fileReader.readAsBinaryString(selectedFile1_5);
                                        fileReader.onload = (event)=>{
                                         let data1_5 = event.target.result;
                                         let workbook1_5 = XLSX.read(data1_5,{type:"binary"});
                                         console.log(workbook1_5);        
                                         workbook1_5.SheetNames.forEach(sheet => {
                                            //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                                objetoAserta5 = XLSX.utils.sheet_to_row_object_array(workbook1_5.Sheets[sheet], {range:1}); //Nombre del array
                                             console.log("ASERTA5")
                                                console.log(objetoAserta5);
                                         });
                                         for (var t=0;t<objetoAserta5.length;t++){
                                            objetoAserta.push(objetoAserta5[t])
                                         }
                                         console.log(objetoAserta)
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
                                                  var encontrar;           
                                
                                                  //Encontrar un valor ahí adentro
                                                  search = (key, ArreySICAS) => {
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
                                                            TC_EstadoCuenta= Math.trunc(TC_EstadoCuenta*10000)/10000
                                                        }
                                                        else{
                                                            var TC_EstadoCuenta= 1
                                                        }
                                                          
                                                          //Obtener importe mxn
                                                          importe_mxn= (ArreySICAS[i]["Importe"]*ArreySICAS[i]["TC"])
                                                          if (polizaSicas == key ) {
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
                                                            encontrar++;
                                                            //Prima Neta
                                                            if(aserta["Prima Neta"] != ArreySICAS[i]["PrimaNeta"]){
                                                                var diferencia= Math.round((aserta["Prima Neta"] - ArreySICAS[i]["PrimaNeta"] )*100)/100;
                                                                errores =  errores + "- Prima Neta\n"
                                                                diferencias = String (diferencia) + "\n"
                                                            }
                                                            //TC = MXN
                                                            if (TC ==1){
                                                                //Comisión Base
                                                            if(ArreySICAS[i]["Importe"] != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                                var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                                                diferencias = diferencias + String(diferenciaCB) + "\n"
                                                                errores = errores + "- Comisión \n"
                                                            }
                                                            if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                                var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                                diferencias = diferencias + String(diferenciaPC) + "\n"
                                                                errores = errores + "- % Comisión \n"
                                                            }
                                                            //Comisión de Derechos
                                                            if(ArreySICAS[i]["Importe"] != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                                var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                                                diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                                errores = errores + "- Comisión Gtos. Exp.\n"
                                                            }
                                                            if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                                var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                                diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                                errores = errores + "- % Comisión Maquila\n"
                                                            }
                                                             //Comisión Especial
                                                             if(ArreySICAS[i]["Importe"] != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                                var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] - ArreySICAS[i]["Importe"])*100)/100;
                                                                diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                                errores = errores + "- Incentivo Prod-Renov\n"
                                                            }
                                                            if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                                var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                                diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                                errores = errores + "- %Incentivo Prod-Renov\n"
                                                            }
                                                            }
                                                            //TC diferente a 1
                                                            else if(TC!=1){
                                                                //Comisión Base
                                                                if(importe_mxn != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                                                    var diferenciaCB= Math.round((aserta["Comisión"] - importe_mxn)*100)/100;
                                                                    diferencias = diferencias + String(diferenciaCB) + "\n"
                                                                    errores = errores + "- Comisión \n"
                                                                    if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                                        errores = errores + "- Tipo de Cambio\n"
                                                                    }
                                                                }
                                                                if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                                                    var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                                    diferencias = diferencias + String(diferenciaPC) + "\n"
                                                                    errores = errores + "- % Comisión \n"
                                                                }
                                                                //Comisión de Derechos
                                                                if(importe_mxn != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                                                    var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - importe_mxn)*100)/100;
                                                                    diferencias = diferencias +  String(diferenciaCG) + "\n"
                                                                    errores = errores + "- Comisión Gtos. Exp.\n"
                                                                    if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                                        errores = errores + "- Tipo de Cambio\n"
                                                                    }
                                                                }
                                                                if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                                                    var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                                    diferencias = diferencias +  String(diferenciaPM) + "\n"
                                                                    errores = errores + "- % Comisión Maquila\n"
                                                                }
                                                                 //Comisión Especial
                                                                 if(importe_mxn != aserta["Incentivo Prod-Renov"] && comision == 'Comisión Especial'){
                                                                    var diferenciaIP= Math.round((aserta["Incentivo Prod-Renov"] -importe_mxn)*100)/100;
                                                                    diferencias = diferencias +  String(diferenciaIP) + "\n"
                                                                    errores = errores + "- Incentivo Prod-Renov\n"
                                                                    if(ArreySICAS[i]["TC"]!=TC_EstadoCuenta){
                                                                        errores = errores + "- Tipo de Cambio\n"
                                                                    }
                                                                }
                                                                if (aserta["%Incentivo Prod-Renov"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Especial'){
                                                                    var diferenciaPI= Math.round((aserta["%Incentivo Prod-Renov"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                                                    diferencias = diferencias +  String(diferenciaPI) + "\n"
                                                                    errores = errores + "- %Incentivo Prod-Renov\n"
                                                                }                                              
                                                                
                                                            }
                                                            
                                                            if(errores != ""){
                                                               var tabla_sicas = ArreySICAS[i]['Nombre Asegurado o Fiado']+"\t<td>"+ArreySICAS[i]['Poliza']+"\t</td><td>"+ArreySICAS[i]['Endoso']+"\t</td><td>"+ArreySICAS[i]['Moneda']+"\t</td><td>"+ArreySICAS[i]['Serie']+"\t</td><td>"+ArreySICAS[i]['TC']+"\t</td><td>"+ArreySICAS[i]['PrimaNeta']+"\t</td><td>"+ArreySICAS[i]['Tipo Comision']+"\t</td><td>"+ArreySICAS[i]['Importe']+"\t</td><td>"+ArreySICAS[i]['% Participacion']+"\t</td>"
                                                               var tabla_EC = "<td>"+aserta['Fiado/Contratante']+"</td><td>"+aserta['No Fianza/']+"</td><td></td><td>"+aserta['Moneda']+"</td><td></td><td>"+porcentaje+"</td><td>"+comision_I+"</td></td><td>"+TC_EstadoCuenta+"</td>"+ "</td><td style='color:var(--b0s-rojo1)'>"+diferencias+"</td>"
                                                               tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+ "<td></td>"+tabla_EC +"<td>"+errores+"</td></tr>"
                                                            }
                                                            
                                                        }
                                                      }
                                                      if(encontrar==0){
                                                        //--tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+aserta["Prima Neta"]+"</td><td>"+aserta["% de Comisión"]+"</td><td>"+comision+"</td><td>"+aserta["Comisión"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                                                        }
                                                        encontrar=0;
                                                    }
                
                
                                                    for(var j=0; j<objetoAserta.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                                        var poliza =  String (objetoAserta[j]["No Fianza/"])
                                                        aserta=objetoAserta[j];
                                                        if(poliza.length == 12){
                                                            resultObject = search(poliza, objetoSICAS)
                                                        }
                                                      
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
                                        document.getElementById("jsondata").innerHTML = "No se adjuntó el 5° Estado de Cuenta de Aserta";
                                    }
                                    }
                                }else{
                                    document.getElementById("jsondata").innerHTML = "No se adjuntó el 4° Estado de Cuenta de Aserta";
                                }
                                }
                            }else{
                                document.getElementById("jsondata").innerHTML = "No se adjuntó el 3° Estado de Cuenta de Aserta";
                            }
                            }
                        }else{
                            document.getElementById("jsondata").innerHTML = "No se adjuntó el 2° Estado de Cuenta de Aserta";
                        }
                         
                        }
                    }else{
                        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Aserta";
                    } 
                break;
            }
             
    });
    
    function ExportToExcel(type, fn, dl) {// función que convierte a excel
        var elt = document.getElementById('Aserta');
        var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
        return dl ?
          XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
          XLSX.writeFile(wb, fn || ('Incidencias Aserta.' + (type || 'xlsx')));
     }