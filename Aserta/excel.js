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

document.getElementById('button').addEventListener("click", () => {
    var num, num2;
    num_ECT = document.getElementById("select").value
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
                                  let tabla ="<table id='Aserta' width='90%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' style='text-align:center;'> <tr><th>Póliza</th><th>Prima Neta</th><th>% Participación</th><th>Tipo Comisión</th><th>Importe</th><th>Diferencia Importe</th><th>Diferencia en:</th></tr>";
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
                                          if (polizaSicas == key ) {
                                            if(comision == 'Comisión Base o de Neta'){
                                                porcentaje = String(aserta["% de Comisión"])
                                            }
                                            else if(comision == 'Comisión de Derechos'){
                                                porcentaje = String(aserta["% Maquila"])
                                            }
                                            else if(comision == 'Comisión Especial'){
                                                porcentaje = String(aserta["%Incentivo Prod-Renov"])
                                            }
                                            encontrar++;
                                            //Prima Neta
                                            if(aserta["Prima Neta"] != ArreySICAS[i]["PrimaNeta"]){
                                                var diferencia= Math.round((aserta["Prima Neta"] - ArreySICAS[i]["PrimaNeta"] )*100)/100;
                                                errores =  errores + "- Prima Neta\n"
                                                diferencias = String (diferencia) + "\n"
                                            }
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
                                            if(errores != ""){
                                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"\t</td><td>"+aserta["Prima Neta"]+"</td><td>"+porcentaje+"</td><td>"+comision+"</td><td>"+aserta["Comisión"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencias+"</td><td>"+errores+"</td></tr>"; 
                                            }
                                            
                                        }
                                      }
                                      if(encontrar==0){
                                        tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+aserta["Prima Neta"]+"</td><td>"+aserta["% de Comisión"]+"</td><td>"+comision+"</td><td>"+aserta["Comisión"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
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
                case 2:
                    if(selectedFile ){ //Función para convertir Edo de Cuenta en array de objetos
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
                                      let tabla ="<table id='Aserta' width='90%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' style='text-align:center;'> <tr><th>Póliza</th><th>Prima Neta</th><th>% Participación</th><th>Tipo Comisión</th><th>Importe</th><th>Diferencia Importe</th><th>Diferencia en:</th></tr>";
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
                                              if (polizaSicas == key ) {
                                                if(comision == 'Comisión Base o de Neta'){
                                                    porcentaje = String(aserta["% de Comisión"])
                                                }
                                                else if(comision == 'Comisión de Derechos'){
                                                    porcentaje = String(aserta["% Maquila"])
                                                }
                                                else if(comision == 'Comisión Especial'){
                                                    porcentaje = String(aserta["%Incentivo Prod-Renov"])
                                                }
                                                encontrar++;
                                                //Prima Neta
                                                if(aserta["Prima Neta"] != ArreySICAS[i]["PrimaNeta"]){
                                                    var diferencia= Math.round((aserta["Prima Neta"] - ArreySICAS[i]["PrimaNeta"] )*100)/100;
                                                    errores =  errores + "- Prima Neta\n"
                                                    diferencias = String (diferencia) + "\n"
                                                }
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
                                                if(errores != ""){
                                                    tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"\t</td><td>"+aserta["Prima Neta"]+"</td><td>"+porcentaje+"</td><td>"+comision+"</td><td>"+aserta["Comisión"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencias+"</td><td>"+errores+"</td></tr>"; 
                                                }
                                                
                                            }
                                          }
                                          if(encontrar==0){
                                            tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+aserta["Prima Neta"]+"</td><td>"+aserta["% de Comisión"]+"</td><td>"+comision+"</td><td>"+aserta["Comisión"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
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
                         
                        }
                    }else{
                        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Aserta";
                    } 
                break;
                case 3:
                break;
                case 4:
                break;
                case 5:
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