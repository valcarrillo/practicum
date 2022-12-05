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
let objetoSICAStransformado=[]; 
let objetoAtlas=[]; //Array de objetos en el que se va a guardar CHUBB
let jsonObj;

document.getElementById('button').addEventListener("click", () => {
    var reng_SICAS=0;
    var reng_EC=0; 
    var tabla = " <table id='Atlas' width='90%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' HIDDEN ><tr><td colspan='10'>SICAS</td><td>--</td><td colspan ='11'>ESTADOS DE CUENTA</td></tr><tr><th>Nombre Asegurado o Fiado</th><th>Póliza</th><th>Endoso</th><th>Moneda</th><th>Serie Recibo</th><th>Tipo Cambio</th><th>Prima Neta</th><th>Tipo Comisión</th><th>Importe</th><th>% Participación</th><th>--</th><th>Nombre Asegurado o Fiado</th><th>Póliza</th><th>Endoso</th><th>Moneda</th><th>Serie Recibo</th><th>% Comisión</th><th>Comisión</th><th>Tipo Cambio</th><th>Diferencia Comisión MXN</th><th>Incidencia</th></tr>";
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
        objeto = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet], {range:13});
                //objetoCHUBB=Object.assign(objetoCHUBB,objeto);
                //objeto={};
                for(var j=0; j<objeto.length; j++){
                    objetoAtlas.push(objeto[j]);
                }
         });
        }
        }
        console.log("objeto Atlas:");
         console.log(objetoAtlas);
         if(selectedFile2){ //Función que convierte SICAS en array de objetos
            let fileReader = new FileReader();
            fileReader.readAsBinaryString(selectedFile2);
            fileReader.onload = (event)=>{
             let data2 = event.target.result;
             let workbook2 = XLSX.read(data2,{type:"binary"});
             workbook2.SheetNames.forEach(sheet => {
                  objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //Nombre del array
                  console.log(objetoSICAS);
                  jsonObj={};
                  // Función para transformar Sicas
                  try {
                  for (var w=0; w<objetoSICAS.length-1;w++){
                    var pol = objetoSICAS[w].Poliza.toString().split('-');  //Divide la póliza de SICAS por '-' . La posición 2 es la fianza y la 3 es la inclusión
                    var SICASpoliza = pol[3];
                    var pol_sig = objetoSICAS[w+1].Poliza.toString().split('-'); //Divide la póliza de SICAS por '-' . La posición 2 es la fianza y la 3 es la inclusión
                    var SICASpoliza_sig = pol_sig[3];
                    var pol_sig_sig = objetoSICAS[w+2].Poliza.toString().split('-');  //Divide la póliza de SICAS por '-' . La posición 2 es la fianza y la 3 es la inclusión
                    var SICASpolizasig_sig = pol_sig_sig[3];
                    var tipo_comision= String(objetoSICAS[w]["Tipo Comision"])  
                    var tipo_comision_sig= String(objetoSICAS[w+1]["Tipo Comision"]) 
                    var tipo_comision_sig_sig= String(objetoSICAS[w+1]["Tipo Comision"]) 
                    var serie = String(objetoSICAS[w]["Serie"]) 
                    var serie_sig = String(objetoSICAS[w+1]["Serie"]) 
                    var serie_transformada = ""
                    var serie_final = ""
                    if (SICASpoliza==SICASpoliza_sig && tipo_comision !='Comisión Especial' && tipo_comision_sig !='Comisión Especial' && serie==serie_sig){
                        var abono = objetoSICAS[w].Importe + objetoSICAS[w+1].Importe
                         serie_transformada = objetoSICAS[w].Serie.toString().split('/')
                         serie_final = +serie_transformada[0] * 100
                        //console.log("\nRenglon :"+w+"'n\Abono="+ abono +"\nSerie"+serie_final)
                        jsonObj={"Concepto": objetoSICAS[w]["Nombre Asegurado o Fiado"],"Poliza": SICASpoliza,"Moneda":objetoSICAS[w]["Moneda"],"Recibo":serie_final, "FechaPago":objetoSICAS[w]["FechaPago"],"TC": objetoSICAS[w]["TC"], "Prima neta":objetoSICAS[w]["PrimaNeta"], "Abono":abono};
                     objetoSICAStransformado.push(jsonObj);
                    jsonObj={};
                    w++;
                    }
                    else if (tipo_comision !='Comisión Especial'){
                        var abono = objetoSICAS[w].Importe
                        serie_transformada = objetoSICAS[w].Serie.toString().split('/')
                         serie_final = +serie_transformada[0] * 100
                        //console.log("\nRenglon :"+w+"'n\Abono="+ abono +"\nSerie"+serie_final)
                        jsonObj={"Concepto": objetoSICAS[w]["Nombre Asegurado o Fiado"],"Poliza": SICASpoliza,"Moneda":objetoSICAS[w]["Moneda"],"Endoso":objetoSICAS[w]["Endoso"],"Recibo":serie_final, "FechaPago":objetoSICAS[w]["FechaPago"],"TC": objetoSICAS[w]["TC"], "Prima neta":objetoSICAS[w]["PrimaNeta"], "Abono":abono};
                     objetoSICAStransformado.push(jsonObj);
                    jsonObj={};
                    }
                  }
                  console.log(objetoSICAStransformado,recibo)  
                    search = (key, ArraySICAS,serie) => {
                        reng_SICAS = ArraySICAS.length
                        for (let i=0; i < ArraySICAS.length; i++) {
                            var errores="";
                          var diferencias ="";
                          var encontrar =0;
                          var endosoSicas = String[ArraySICAS[i]["Endoso"]]
                          var endoso = "";
                          
                          if (endosoSicas == undefined){
                            endoso= "0"
                          }
                          console.log(endoso)
                            if(key == ArraySICAS[i].Poliza && recibo == ArraySICAS[i].Recibo){
                                encontrar = 1;
                                //Comparar abono
                                if(String(ArraySICAS[i]['Abono']) != atlas['Abono']){
                                var diferenciaCB= Math.round((atlas["Abono"] - ArraySICAS[i]["Abono"])*100)/100;
                                  diferencias = diferencias + String(diferenciaCB) + "\n"
                                  errores = errores + "- Comisión \n"
                                }
                                else{
                                    var diferenciaCB= Math.round((atlas["Abono"] - ArraySICAS[i]["Abono"])*100)/100;
                                    diferencias = diferencias + String(diferenciaCB) + "\n" 
                                }
                                if (endoso != atlas['Endoso']){
                                  errores = errores + "-Endoso"
                                }
                                var tabla_sicas = "\t<td>"+ArraySICAS[i]['Poliza']+"\t</td><td>"+endoso+"\t</td><td>"+ArraySICAS[i]['Moneda']+"\t</td><td>'"+ArraySICAS[i]['Recibo']+"'\t</td><td>"+ArraySICAS[i]['TC']+"\t</td><td>"+ArraySICAS[i]['Prima neta']+"\t</td><td>\t</td><td>"+ArraySICAS[i]['Abono']+"\t</td><td>\t</td>"
                                var tabla_EC = "<td></td><td>"+atlas['Póliza']+"</td><td>"+atlas['Endoso']+"</td><td>"+atlas['Moneda']+"</td><td>"+atlas['Recibo']+"</td><td></td><td>"+atlas['Abono']+"</td><td style='color:var(--b0s-rojo1)'></td><td>"+diferencias+"</td>"
                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+ "<td></td>"+tabla_EC +"<td>"+errores+"</td></tr>"  
                            }
                            
                        }
                        if (encontrar == 0){
                            //Arreglar columnas
                            var tabla_EC = "<td></td><td>"+atlas['Póliza']+"</td><td>"+endoso+"</td><td>"+atlas['Moneda']+"</td><td>"+atlas['Recibo']+"</td><td></td>"+atlas['Abono']+"<td></td><td></td><td style='color:var(--b0s-rojo1)'></td><td>NO SE ENCONTRÓ</td>"
                            var tabla_sicas ="Invalido<td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td><td>Invalido</td>"
                            tablaNA=tablaNA+"<tr><td style='background-color:var(--bs-azul3)'>"+tabla_sicas+"</td><td>--</td>"+tabla_EC+"</tr>"     
                        }
                    }
                    //Dar formato a la tabla 
                    
                      for(var j=0; j<objetoAtlas.length; j++){ //Ciclo que va a buscar cada poliza de SICAS en Berkley                          
                       var poliza = String(objetoAtlas[j].Póliza)
                       var recibo = String(objetoAtlas[j].Recibo)
                       atlas = objetoAtlas[j]
                       reng_EC = reng_EC + 1;
                       resultObject = search(poliza, objetoSICAStransformado,recibo)
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
                       document.getElementById("jsondata").innerHTML = "Renglones Estado de Cuenta: "+reng_EC+"\nRenglones SICAS: "+reng_SICAS+"\n";
                        ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente.
                        
                        //document.getElementById("jsondata").innerHTML = "Renglones Estado de Cuenta: "+reng_EC+"\nRenglones SICAS: "+reng_SICAS+"\n";
                      }
                      //setInterval("location.reload()",10);
            }
          catch (error) { //Si hay un error aquí se muestra.
            document.getElementById("jsondata").innerHTML = "Algo salió mal al leer el documento. Revise que el encabezado tenga el formato correcto. Error: "+error;
          }
        });
             
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
    var elt = document.getElementById('Atlas');
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    return dl ?
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || ('Incidencias Aserta.' + (type || 'xlsx')));
 }