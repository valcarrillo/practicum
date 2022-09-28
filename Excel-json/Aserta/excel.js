//https://levelup.gitconnected.com/how-to-convert-excel-file-into-json-object-by-using-javascript-9e95532d47c5
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
let objetoAserta; //Array de objetos en el que se va a guardar Berkley

document.getElementById('button').addEventListener("click", () => {
    var num, num2;
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
                  let fianzas ="<table width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Campo</th><th>BERKLEY FIANZAS</th><th>SICAS</th></tr>";
                  let resultObject;
                  let sicas;
                  var encontrar;

                  comparar = (aser, sic) => {
                    if(aser == sic){
                        fianzas=fianzas+"<tr style='color:green;'>"
                    }else{
                        fianzas=fianzas+"<tr style='background-color:var(--bs-rojo2);'>"
                    }
                  }
                 //Encontrar un valor ahí adentro
                  search = (key, inclu, inputArray) => {
                      for (let i=0; i < inputArray.length; i++) {
                          if (inputArray[i].FIANZA == key) {
                            inputArray[i].
                            console.log(inputArray[i].INCLUSION)
                            if (inputArray[i].INCLUSION == inclu) {
                            encontrar=1;
                            fianzas= fianzas+"<tr style='background-color:var(--bs-azul3)'><td>Num póliza</td><td>"+inputArray[i].FIANZA+"</td><td>"+key+"</td></tr>";
                            fianzas= fianzas+"<tr style='background-color:var(--bs-azul4)'><td>Inclusión</td><td>"+inputArray[i].INCLUSION+"</td><td>"+inclu+"</td></tr>";
                            comparar(inputArray[i]["PRIMA NETA"], sicas["PrimaNeta"]);
                            fianzas=fianzas+"<td>Prima Neta</td><td>"+inputArray[i]["PRIMA NETA"]+"</td><td>"+sicas["PrimaNeta"]+"</td></tr>";
                            comparar(inputArray[i]["% COMISION"], sicas["% Participacion"]);
                            fianzas= fianzas+"<td>% Comisión</td><td>"+inputArray[i]["% COMISION"]+"</td><td>"+sicas["% Participacion"]+"</td></tr>";
                            comparar(inputArray[i].COMISIONES, sicas["Importe"]);
                            fianzas= fianzas+"<td>Importe</td><td>"+inputArray[i].COMISIONES+"</td><td>"+sicas["Importe"]+"</td></tr>";
                            document.getElementById("jsondata").innerHTML = fianzas+"<tr><td></td><td></td><td></td></tr></table>";
                              return inputArray[i];
                            }
                          }
                      }
                      if(encontrar==0){
                      fianzas= fianzas+"<tr style='background-color:var(--bs-rojo1)'><td>Num póliza</td><td>No se encontró</td><td>"+key+"-"+inclu+"</td></tr>";
                      "<p>"+key+" de SICAS  no se encontró</p>";
                      }
                      encontrar=0; 
                    }
                    //;modificar poliza
                    for(var j=0; j<objetoSICAS.length; j++){
                        var pol = objetoSICAS[j].Poliza.split('-'),
                        poliza = pol[2];
                        inclusion=pol[3];
                        if(typeof inclusion === 'undefined'){
                            inclusion=0;
                        }
                        num = +poliza;
                        sicas=objetoSICAS[j];//renglon de sicas
                      resultObject = search(num, inclusion, objetoAserta);
                      console.log(resultObject);
                      console.log("Número de registros en sicas: "+j);
                    }
                    if(resultObject==0){
                        document.getElementById("jsondata").innerHTML = "No se encontró ninguna fianza";

                    }
                            }
                );
             
            }
        } else{

        }
        // document.getElementById("jsondata2").innerHTML = "No se adjuntó nada";
        }
        
    }else{
        //document.getElementById("jsondata").innerHTML = "No se adjuntó nada";
    }
    
    
});

