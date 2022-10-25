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

            document.getElementById("jsondata").innerHTML = "No se selecciono el número de Estado de Cuenta";

        //Caso de tener 1 estado de cuenta

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
                                  if (polizaSicas == key ) {
                                    encontrar++;
                                    //Prima Neta
                                    if(aserta["Prima Neta"] != ArreySICAS[i]["PrimaNeta"] ){
                                        var diferencia= Math.round((aserta["Prima Neta"] - ArreySICAS[i]["PrimaNeta"] )*100)/100;
                                        errores =  errores + "-Prima Neta"
                                        diferencias = String (diferencia)
                                    }
                                    //Comisión Base
                                    if(ArreySICAS[i]["Importe"] != aserta["Comisión"] && comision == 'Comisión Base o de Neta'){
                                        var diferenciaCB= Math.round((aserta["Comisión"] - ArreySICAS[i]["Importe"])*100)/100;
                                        diferencias = diferencias + "\n" + String(diferenciaCB)
                                        errores = errores + "\n- Importe Comisión Base"
                                    }
                                    if (aserta["% de Comisión"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión Base o de Neta'){
                                        var diferenciaPC= Math.round((aserta["% de Comisión"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                        diferencias = diferencias + "\n" + String(diferenciaPC)
                                        errores = errores + "\n- % Comisión Base"
                                    }
                                    //Comisión de Derechos
                                    if(ArreySICAS[i]["Importe"] != aserta["Comisión Gtos. Exp."] && comision == 'Comisión de Derechos'){
                                        var diferenciaCG= Math.round((aserta["Comisión Gtos. Exp."] - ArreySICAS[i]["Importe"])*100)/100;
                                        diferencias = diferencias + "\n" + String(diferenciaCG)
                                        errores = errores + "\n- Importe Comisión Gtos. Exp."
                                    }
                                    if (aserta["% Maquila"] != ArreySICAS[i]["% Participacion"] && comision == 'Comisión de Derechos'){
                                        var diferenciaPM= Math.round((aserta["% Maquila"] -ArreySICAS[i]["% Participacion"] )*100)/100;
                                        diferencias = diferencias + "\n" + String(diferenciaPM)
                                        errores = errores + "\n- % Comisión Maquila"
                                    }
                                    if(errores != ""){
                                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+key+"\t</td><td>"+aserta["Prima Neta"]+"</td><td>"+aserta["% de Comisión"]+"</td><td>"+comision+"</td><td>"+aserta["Comisión"]+"</td><td style='color:var(--bs-rojo1)'>"+diferencias+"</td><td>"+errores+"</td></tr>"; 
                                    }
                                    
                                    //Comisión Derechos
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
                document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Berkley";
            }  
    });
    
    function ExportToExcel(type, fn, dl) {// función que convierte a excel
        var elt = document.getElementById('Aserta');
        var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
        return dl ?
          XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
          XLSX.writeFile(wb, fn || ('Incidencias Aserta.' + (type || 'xlsx')));
     }