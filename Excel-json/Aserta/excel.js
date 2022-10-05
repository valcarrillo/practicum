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
    //console.log(num_ECT);

    switch (num_ECT){
        case '0':
            document.getElementById("jsondata").innerHTML = "No se selecciono número de Estados de cuenta";
            break;
        //Caso de tener 1 estado de cuenta
        case '1':
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
                 if(selectedFile2){ //Función que convierte SICAS en array de objetos
                    let fileReader = new FileReader();
                    fileReader.readAsBinaryString(selectedFile2);
                    fileReader.onload = (event)=>{
                     let data2 = event.target.result;
                     let workbook2 = XLSX.read(data2,{type:"binary"});
                     workbook2.SheetNames.forEach(sheet => {
                          objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //Nombre del array
        
                            //La tabla tiene atributo HIDDEN para que no se vea, pero ahí está.
                          let tabla ="<table id='Aserta' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Tipo Comisión</th><th>Total Comisión</th><th>Diferencia comisión</th><th>Diferencia en:</th></tr>";
                          let resultObject;
                          let sicas;
                          var encontrar;           
        
                          //Encontrar un valor ahí adentro
                          search = (key, ArreyAserta) => {
                              for (let i=0; i < ArreyAserta.length; i++) {
                                var polizaAserta = String(ArreyAserta[i]["No Fianza/"])
                                var comision = String(sicas["Tipo Comision"])
                                  if (polizaAserta == key) {
                                    encontrar++;
                                    if ( comision == 'Comisión Base o de Neta' ){
                                        if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                            var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                            console.log("La diferencia es de"+diferencia);
                                            tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                        }
                                        else if (ArreyAserta[i]["Comisión"] !=sicas["Importe"] ){
                                            var diferencia= Math.round((ArreyAserta[i]["Comisión"] -sicas["Importe"])*100)/100;
                                            tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"\t</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión</td></tr>";
                                        }
                                    }
                                    else if (comision == 'Comisión de Derechos'){
                                        if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                            var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                            //console.log("La diferencia es de"+diferencia);
                                            tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td<td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                        }
                                        else if(ArreyAserta[i]["Comisión Gtos. Exp."]!= sicas["Importe"]) {
                                            var diferencia= Math.round((ArreyAserta[i]["Comisión Gtos. Exp."] -sicas["Importe"])*100)/100;
                                            //console.log("La diferencia es de"+diferencia);
                                            tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión Gtos.</td></tr>";
                                        }
                                    }
                                    else if (comision == 'Comisión Especial'){
                                        if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                            var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                            //console.log("La diferencia es de"+diferencia);
                                            tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td<td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                        }
                                        else if(ArreyAserta[i]["Incentivo Prod-Renov"]!= sicas["Importe"]) {
                                            var diferencia= Math.round((ArreyAserta[i]["Incentivo Prod-Renov"] -sicas["Importe"])*100)/100;
                                            //console.log("La diferencia es de"+diferencia);
                                            tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión Gtos.</td></tr>";
                                        }
                                    }
                                  }
        
                              }
                              if(encontrar==0){
                                tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                                }
                                encontrar=0;
                            }
                            for(var j=0; j<objetoSICAS.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                var poliza =  String (objetoSICAS[j].Poliza)
                                sicas=objetoSICAS[j];
                              resultObject = search(poliza, objetoAserta)
                              //console.log("Número de registros en sicas: "+j);
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
                document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Berkley";
            }
            break;
            //Caso 2 Estados de Cuente
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
                         console.log(workbook1_2);        
                         workbook1_2.SheetNames.forEach(sheet => {
                            //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                objetoAserta2 = XLSX.utils.sheet_to_row_object_array(workbook1_2.Sheets[sheet], {range:1}); //Nombre del array
                             console.log("ASERTA2")
                                console.log(objetoAserta2);
                         });
                         for (var t=0;t<objetoAserta2.length;t++){
                            objetoAserta.push(objetoAserta2[t])

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
                                  let tabla ="<table id='Aserta' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Tipo Comisión</th><th>Total Comisión</th><th>Diferencia comisión</th><th>Diferencia en:</th></tr>";
                                  let resultObject;
                                  let sicas;
                                  var encontrar;           
                    
                                  //Encontrar un valor ahí adentro
                                  search = (key, ArreyAserta) => {
                                      for (let i=0; i < ArreyAserta.length; i++) {
                                        var polizaAserta = String(ArreyAserta[i]["No Fianza/"])
                                        var comision = String(sicas["Tipo Comision"])
                                          if (polizaAserta == key) {
                                            encontrar++;
                                            if ( comision == 'Comisión Base o de Neta' ){
                                                if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                    var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                    console.log("La diferencia es de"+diferencia);
                                                    tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                }
                                                else if (ArreyAserta[i]["Comisión"] !=sicas["Importe"] ){
                                                    var diferencia= Math.round((ArreyAserta[i]["Comisión"] -sicas["Importe"])*100)/100;
                                                    tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"\t</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión</td></tr>";
                                                }
                                            }
                                            else if (comision == 'Comisión de Derechos'){
                                                if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                    var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                    //console.log("La diferencia es de"+diferencia);
                                                    tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td<td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                }
                                                else if(ArreyAserta[i]["Comisión Gtos. Exp."]!= sicas["Importe"]) {
                                                    var diferencia= Math.round((ArreyAserta[i]["Comisión Gtos. Exp."] -sicas["Importe"])*100)/100;
                                                    //console.log("La diferencia es de"+diferencia);
                                                    tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión Gtos.</td></tr>";
                                                }
                                            }
                                            else if (comision == 'Comisión Especial'){
                                                if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                    var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                    //console.log("La diferencia es de"+diferencia);
                                                    tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td<td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                }
                                                else if(ArreyAserta[i]["Incentivo Prod-Renov"]!= sicas["Importe"]) {
                                                    var diferencia= Math.round((ArreyAserta[i]["Incentivo Prod-Renov"] -sicas["Importe"])*100)/100;
                                                    //console.log("La diferencia es de"+diferencia);
                                                    tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión Gtos.</td></tr>";
                                                }
                                            }
                                          }
                    
                                      }
                                      if(encontrar==0){
                                        tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                                        }
                                        encontrar=0;
                                    }
                                    for(var j=0; j<objetoSICAS.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                        var poliza =  String (objetoSICAS[j].Poliza)
                                        sicas=objetoSICAS[j];
                                      resultObject = search(poliza, objetoAserta)
                                      //console.log("Número de registros en sicas: "+j);
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
                //Caso tener 3 estados de cuenta
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
                         console.log(workbook1_2);        
                         workbook1_2.SheetNames.forEach(sheet => {
                            //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                objetoAserta2 = XLSX.utils.sheet_to_row_object_array(workbook1_2.Sheets[sheet], {range:1}); //Nombre del array
                             console.log("ASERTA2")
                                console.log(objetoAserta2);
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
                                      let tabla ="<table id='Aserta' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Tipo Comisión</th><th>Total Comisión</th><th>Diferencia comisión</th><th>Diferencia en:</th></tr>";
                                      let resultObject;
                                      let sicas;
                                      var encontrar;           
                        
                                      //Encontrar un valor ahí adentro
                                      search = (key, ArreyAserta) => {
                                          for (let i=0; i < ArreyAserta.length; i++) {
                                            var polizaAserta = String(ArreyAserta[i]["No Fianza/"])
                                            var comision = String(sicas["Tipo Comision"])
                                              if (polizaAserta == key) {
                                                encontrar++;
                                                if ( comision == 'Comisión Base o de Neta' ){
                                                    if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                        var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                        console.log("La diferencia es de"+diferencia);
                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                    }
                                                    else if (ArreyAserta[i]["Comisión"] !=sicas["Importe"] ){
                                                        var diferencia= Math.round((ArreyAserta[i]["Comisión"] -sicas["Importe"])*100)/100;
                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"\t</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión</td></tr>";
                                                    }
                                                }
                                                else if (comision == 'Comisión de Derechos'){
                                                    if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                        var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                        //console.log("La diferencia es de"+diferencia);
                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td<td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                    }
                                                    else if(ArreyAserta[i]["Comisión Gtos. Exp."]!= sicas["Importe"]) {
                                                        var diferencia= Math.round((ArreyAserta[i]["Comisión Gtos. Exp."] -sicas["Importe"])*100)/100;
                                                        //console.log("La diferencia es de"+diferencia);
                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión Gtos.</td></tr>";
                                                    }
                                                }
                                                else if (comision == 'Comisión Especial'){
                                                    if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                        var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                        //console.log("La diferencia es de"+diferencia);
                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td<td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                    }
                                                    else if(ArreyAserta[i]["Incentivo Prod-Renov"]!= sicas["Importe"]) {
                                                        var diferencia= Math.round((ArreyAserta[i]["Incentivo Prod-Renov"] -sicas["Importe"])*100)/100;
                                                        //console.log("La diferencia es de"+diferencia);
                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión Gtos.</td></tr>";
                                                    }
                                                }
                                              }
                        
                                          }
                                          if(encontrar==0){
                                            tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                                            }
                                            encontrar=0;
                                        }
                                        for(var j=0; j<objetoSICAS.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                            var poliza =  String (objetoSICAS[j].Poliza)
                                            sicas=objetoSICAS[j];
                                          resultObject = search(poliza, objetoAserta)
                                          //console.log("Número de registros en sicas: "+j);
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
                    document.getElementById("jsondata").innerHTML = "No se adjuntó el 1° Estado de Cuenta de Aserta";
                }
                break;
                //Caso de tener 4 Estados de Cuenta
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
                             console.log(workbook1_2);        
                             workbook1_2.SheetNames.forEach(sheet => {
                                //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                    objetoAserta2 = XLSX.utils.sheet_to_row_object_array(workbook1_2.Sheets[sheet], {range:1}); //Nombre del array
                                 console.log("ASERTA2")
                                    console.log(objetoAserta2);
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
                                              let tabla ="<table id='Aserta' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Tipo Comisión</th><th>Total Comisión</th><th>Diferencia comisión</th><th>Diferencia en:</th></tr>";
                                              let resultObject;
                                              let sicas;
                                              var encontrar;           
                                    
                                              //Encontrar un valor ahí adentro
                                              search = (key, ArreyAserta) => {
                                                  for (let i=0; i < ArreyAserta.length; i++) {
                                                    var polizaAserta = String(ArreyAserta[i]["No Fianza/"])
                                                    var comision = String(sicas["Tipo Comision"])
                                                      if (polizaAserta == key) {
                                                        encontrar++;
                                                        if ( comision == 'Comisión Base o de Neta' ){
                                                            if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                                var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                                console.log("La diferencia es de"+diferencia);
                                                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                            }
                                                            else if (ArreyAserta[i]["Comisión"] !=sicas["Importe"] ){
                                                                var diferencia= Math.round((ArreyAserta[i]["Comisión"] -sicas["Importe"])*100)/100;
                                                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"\t</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión</td></tr>";
                                                            }
                                                        }
                                                        else if (comision == 'Comisión de Derechos'){
                                                            if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                                var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                                //console.log("La diferencia es de"+diferencia);
                                                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td<td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                            }
                                                            else if(ArreyAserta[i]["Comisión Gtos. Exp."]!= sicas["Importe"]) {
                                                                var diferencia= Math.round((ArreyAserta[i]["Comisión Gtos. Exp."] -sicas["Importe"])*100)/100;
                                                                //console.log("La diferencia es de"+diferencia);
                                                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión Gtos.</td></tr>";
                                                            }
                                                        }
                                                        else if (comision == 'Comisión Especial'){
                                                            if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                                var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                                //console.log("La diferencia es de"+diferencia);
                                                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td<td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                            }
                                                            else if(ArreyAserta[i]["Incentivo Prod-Renov"]!= sicas["Importe"]) {
                                                                var diferencia= Math.round((ArreyAserta[i]["Incentivo Prod-Renov"] -sicas["Importe"])*100)/100;
                                                                //console.log("La diferencia es de"+diferencia);
                                                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión Gtos.</td></tr>";
                                                            }
                                                        }
                                                      }
                                    
                                                  }
                                                  if(encontrar==0){
                                                    tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                                                    }
                                                    encontrar=0;
                                                }
                                                for(var j=0; j<objetoSICAS.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                                    var poliza =  String (objetoSICAS[j].Poliza)
                                                    sicas=objetoSICAS[j];
                                                  resultObject = search(poliza, objetoAserta)
                                                  //console.log("Número de registros en sicas: "+j);
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
                        document.getElementById("jsondata").innerHTML = "No se adjuntó el 1° Estado de Cuenta de Aserta";
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
                                 console.log(workbook1_2);        
                                 workbook1_2.SheetNames.forEach(sheet => {
                                    //console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                                        objetoAserta2 = XLSX.utils.sheet_to_row_object_array(workbook1_2.Sheets[sheet], {range:1}); //Nombre del array
                                     console.log("ASERTA2")
                                        console.log(objetoAserta2);
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
                                         let workbook1_4 = XLSX.read(data1_3,{type:"binary"});
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
                                                      let tabla ="<table id='Aserta' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Tipo Comisión</th><th>Total Comisión</th><th>Diferencia comisión</th><th>Diferencia en:</th></tr>";
                                                      let resultObject;
                                                      let sicas;
                                                      var encontrar;           
                                            
                                                      //Encontrar un valor ahí adentro
                                                      search = (key, ArreyAserta) => {
                                                          for (let i=0; i < ArreyAserta.length; i++) {
                                                            var polizaAserta = String(ArreyAserta[i]["No Fianza/"])
                                                            var comision = String(sicas["Tipo Comision"])
                                                              if (polizaAserta == key) {
                                                                encontrar++;
                                                                if ( comision == 'Comisión Base o de Neta' ){
                                                                    if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                                        var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                                        console.log("La diferencia es de"+diferencia);
                                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                                    }
                                                                    else if (ArreyAserta[i]["Comisión"] !=sicas["Importe"] ){
                                                                        var diferencia= Math.round((ArreyAserta[i]["Comisión"] -sicas["Importe"])*100)/100;
                                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"\t</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión</td></tr>";
                                                                    }
                                                                }
                                                                else if (comision == 'Comisión de Derechos'){
                                                                    if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                                        var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                                        //console.log("La diferencia es de"+diferencia);
                                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td<td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                                    }
                                                                    else if(ArreyAserta[i]["Comisión Gtos. Exp."]!= sicas["Importe"]) {
                                                                        var diferencia= Math.round((ArreyAserta[i]["Comisión Gtos. Exp."] -sicas["Importe"])*100)/100;
                                                                        //console.log("La diferencia es de"+diferencia);
                                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión Gtos.</td></tr>";
                                                                    }
                                                                }
                                                                else if (comision == 'Comisión Especial'){
                                                                    if(ArreyAserta[i]["Prima Neta"] != sicas["PrimaNeta"]){
                                                                        var diferencia= Math.round((ArreyAserta[i]["Prima Neta"] -sicas["PrimaNeta"])*100)/100;
                                                                        //console.log("La diferencia es de"+diferencia);
                                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td<td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Prima Neta</td></tr>";
                                                                    }
                                                                    else if(ArreyAserta[i]["Incentivo Prod-Renov"]!= sicas["Importe"]) {
                                                                        var diferencia= Math.round((ArreyAserta[i]["Incentivo Prod-Renov"] -sicas["Importe"])*100)/100;
                                                                        //console.log("La diferencia es de"+diferencia);
                                                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+ArreyAserta[i]["No Fianza/"]+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencia+"</td><td>Comisión Gtos.</td></tr>";
                                                                    }
                                                                }
                                                              }
                                            
                                                          }
                                                          if(encontrar==0){
                                                            tabla= tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Tipo Comision"]+"</td><td>"+sicas["Importe"]+"</td><td></td><td style='background-color:var(--bs-rojo2)'>NO SE ENCONTRÓ</td></tr>";
                                                            }
                                                            encontrar=0;
                                                        }
                                                        for(var j=0; j<objetoSICAS.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley
                                                            var poliza =  String (objetoSICAS[j].Poliza)
                                                            sicas=objetoSICAS[j];
                                                          resultObject = search(poliza, objetoAserta)
                                                          //console.log("Número de registros en sicas: "+j);
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
                            document.getElementById("jsondata").innerHTML = "No se adjuntó el 1° Estado de Cuenta de Aserta";
                        }                        
                        break;
                    }    
});

function ExportToExcel(type, fn, dl) {// función que convierte a excel
    var elt = document.getElementById('Aserta');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    return dl ?
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || ('Aseguradora Aserta.' + (type || 'xlsx')));
 }