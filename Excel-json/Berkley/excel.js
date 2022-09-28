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
let objetoBerkley; //Array de objetos en el que se va a guardar Berkley

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
                objetoBerkley = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet], {range:3}); //Nombre del array
             console.log(objetoBerkley);
            // document.getElementById("jsondata").innerHTML = JSON.stringify(objetoBerkley,undefined,4)
         });

         if(selectedFile2){ //Función que convierte SICAS en array de objetos
            let fileReader = new FileReader();
            fileReader.readAsBinaryString(selectedFile2);
            fileReader.onload = (event)=>{
             let data2 = event.target.result;
             let workbook2 = XLSX.read(data2,{type:"binary"});
             //console.log(workbook2);
             workbook2.SheetNames.forEach(sheet => {
                  objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //Nombre del array

                  let fianzas;
                  let tabla ="<table width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001'> <tr><th>Póliza</th><th>Prima Neta</th><th>% Comisión</th><th>Total Comisión</th><th>Diferencia comisión</th></tr>";
                  let resultObject;
                  let sicas;
                  var encontrar;

                  comparar = (berk, sic) => {
                    if(berk != sic){
                        var diferencia= berk-sic;
                        tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+inputArray[i].FIANZA+"-"+inputArray[i].INCLUSION+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo2);'>"+diferencia+"</td></tr></table>";
                    }else{
                        fianzas=" ";
                    }
                  }
                  

                  //Encontrar un valor ahí adentro
                  search = (key, inclu, inputArray) => {
                      for (let i=0; i < inputArray.length; i++) {
                          if (inputArray[i].FIANZA == key) {
                            if (inputArray[i].INCLUSION == inclu) {
                            encontrar=1;
                            //comparar(inputArray[i]["PRIMA NETA"], sicas["PrimaNeta"]);
                            if(inputArray[i]["PRIMA NETA"] != sicas["PrimaNeta"]){
                                var diferencia= berk-sic;
                                tabla=tabla+"<tr><td style='background-color:var(--bs-azul3)'>"+inputArray[i].FIANZA+"-"+inputArray[i].INCLUSION+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Importe"]+"</td><td style='color:var(--bs-rojo2);'>"+diferencia+"</td></tr>";
                            }
                           // fianzas=fianzas+"<td>Prima Neta</td><td>"+inputArray[i]["PRIMA NETA"]+"</td><td>"+sicas["PrimaNeta"]+"</td></tr>";
                            //comparar(inputArray[i]["% COMISION"], sicas["% Participacion"]);
                            //fianzas= fianzas+"<td>% Comisión</td><td>"+inputArray[i]["% COMISION"]+"</td><td>"+sicas["% Participacion"]+"</td></tr>";
                            //comparar(inputArray[i].COMISIONES, sicas["Importe"]);
                           // fianzas= fianzas+"<td>Importe</td><td>"+inputArray[i].COMISIONES+"</td><td>"+sicas["Importe"]+"</td></tr>";
                              return inputArray[i];
                            }
                          }
                      }
                      if(encontrar==0){
                      tabla= tabla+"<tr style='background-color:var(--bs-rojo1)'><td>"+inputArray[i].FIANZA+"-"+inputArray[i].INCLUSION+"</td><td>"+sicas["PrimaNeta"]+"</td><td>"+sicas["% Participacion"]+"</td><td>"+sicas["Importe"]+"</td><td>NO SE ENCONTRÓ</td>";
                      }
                      encontrar=0; 
                    }
                    for(var j=0; j<objetoSICAS.length; j++){
                        var pol = objetoSICAS[j].Poliza.split('-'),
                        poliza = pol[2];
                        inclusion=pol[3];
                        if(typeof inclusion === 'undefined'){
                            inclusion=0;
                        }
                        num = +poliza;
                        sicas=objetoSICAS[j];
                      resultObject = search(num, inclusion, objetoBerkley);
                      console.log(resultObject);
                      console.log("Número de registros en sicas: "+j);
                      document.getElementById("jsondata").innerHTML = tabla+"</table>";
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

