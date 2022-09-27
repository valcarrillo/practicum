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
                objetoBerkley = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet], {range:1}); //Nombre del array
             console.log(objetoBerkley);
            // document.getElementById("jsondata").innerHTML = JSON.stringify(objetoBerkley,undefined,4)

                //Forma de obtener un renglón
               //console.log(objetoBerkley[2]);
               //convtir un subarray en array
              /*  const array1 = objetoBerkley[2];
                num2 = +objetoBerkley[2].FIANZA;
                console.log("Berkley");
                console.log("Fianza: "+objetoBerkley[2].FIANZA);
                console.log("Prima Neta: "+objetoBerkley[2]["PRIMA NETA"]);
                console.log("El % de comisión es: "+objetoBerkley[2]["% COMISION"]);
                console.log("IVA: "+objetoBerkley[2]["IVA"]);
                console.log("Los gastos fueron de: "+objetoBerkley[2].GASTOS);
                console.log("El total de comisión es: "+objetoBerkley[2]["TOTAL COMISION"]); //[] Para nombres compuestos
                console.log("Movimiento: "+objetoBerkley[2].MOVIMIENTO);*/
                // console.log(array1);   
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
                 // console.log(objetoSICAS);
                 // document.getElementById("jsondata2").innerHTML = JSON.stringify(objetoSICAS,undefined,4);
                  //console.log(objetoSICAS[7].Poliza);
                   /*
                  console.log("SICAS");
                  console.log("Número de Póliza: "+num);
                  console.log("Prima Neta: "+objetoSICAS[2]["PrimaNeta"]);
                  console.log("% Participacion: "+objetoSICAS[2]["% Participacion"]);
                  console.log("Importe: "+objetoSICAS[2]["Importe"]);
                  console.log("Importe pendiente: "+objetoSICAS[2]["Importe pendiente"]);
                  console.log("Endoso: "+objetoSICAS[2]["Endoso"]);*/

                 


                  let fianzas;
                  let resultObject;
                  let sicas;
                  var encontrar;

                  comparar = (berk, sic) => {
                    if(berk == sic){
                        fianzas=fianzas+" SON IGUALES </p>"
                    }else{
                        fianzas=fianzas+" SON DIFERENTES </p>"
                    }
                  }
                  //Encontrar un valor ahí adentro
                  search = (key, inputArray) => {
                      for (let i=0; i < inputArray.length; i++) {
                          if (inputArray[i].FIANZA == key) {
                            encontrar=1;
                            fianzas= fianzas+"<p> Num Póliza: "+inputArray[i].FIANZA+" de Berkley es igual a "+key+" de SICAS "; 
                            fianzas= fianzas+"<p> Prima Neta: "+inputArray[i]["PRIMA NETA"]+" de Berkley es igual a "+sicas["PrimaNeta"]+" de SICAS "; 
                            comparar(inputArray[i]["PRIMA NETA"], sicas["PrimaNeta"]);
                            fianzas= fianzas+"<p> % Comisión: "+inputArray[i]["% COMISION"]+" de Berkley es igual a "+sicas["% Participacion"]+" de SICAS "; 
                            comparar(inputArray[i]["% COMISION"], sicas["% Participacion"]);
                            fianzas= fianzas+"<p> Iva: "+inputArray[i]["IVA"]+" de Berkley es igual a no sé de SICAS "; 
                            comparar(inputArray[i]["IVA"], 0);
                            fianzas= fianzas+"<p> Importe: "+inputArray[i].COMISIONES+" de Berkley es igual a "+sicas["Importe"]+" de SICAS "; 
                            comparar(inputArray[i].COMISIONES, sicas["Importe"]);
                            fianzas= fianzas+"<p> Total comisión: "+inputArray[i]["TOTAL COMISION"]+" de Berkley es igual a "+sicas["Importe pendiente"]+" de SICAS";
                            comparar(inputArray[i]["TOTAL COMISION"], sicas["Importe pendiente"]);
                            fianzas= fianzas+"<br>" 
                            document.getElementById("jsondata").innerHTML = fianzas;
                              return inputArray[i];
                          }
                      }
                      if(encontrar==0){
                      fianzas= fianzas+"<p>"+key+" de SICAS  no se encontró</p>";
                      }
                      encontrar=0; 
                    }
                    //Moddifcación de polizas
                    for(var j=0; j< objetoSICAS.length; j++){
                        var pol = objetoSICAS[j].Poliza.split('-'),
                        poliza = pol[2];
                        num = +poliza;
                        sicas=objetoSICAS[j];
                      resultObject = search(num, objetoBerkley);
                      console.log(resultObject);
                      console.log("Número de registros en sicas: "+j);
                    }
                    if(resultObject==0){
                        document.getElementById("jsondata").innerHTML = "No se encontró ninguna fianza";

                    }

                    //console.log(num);
                    //console.log(num2);
                   /* if(num==num2){
                        document.getElementById("jsondata2").innerHTML = num+" "+ num2+" Son iguales";
                        console.log("Son iguales");
                    }else{
                        document.getElementById("jsondata2").innerHTML = num+" "+ num2+" Son diferentes";
                        console.log("Son diferentes");
                    }*/
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

