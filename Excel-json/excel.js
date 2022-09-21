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
    if(selectedFile){ //Función para convertir Edo de Cuenta en array de objetos
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);
        fileReader.onload = (event)=>{
         let data1 = event.target.result;
         let workbook1 = XLSX.read(data1,{type:"binary"});
         console.log(workbook1);        
         workbook1.SheetNames.forEach(sheet => {
            console.log(workbook1.Sheets[sheet]);                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
                objetoBerkley = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet], {range:3}); //Nombre del array
             console.log(objetoBerkley);
             document.getElementById("jsondata").innerHTML = JSON.stringify(objetoBerkley,undefined,4)

                //Forma de obtener un renglón
               console.log(objetoBerkley[2]);
               //convtir un subarray en array
                const array1 = objetoBerkley[2];
                console.log("Fianza: "+objetoBerkley[2].FIANZA);
                console.log("Los gastos fueron de: "+objetoBerkley[2].GASTOS);
                console.log("El total de comisión es: "+objetoBerkley[2]["TOTAL COMISION"]); //[] Para nombres compuestos
                console.log("El % de comisión es: "+objetoBerkley[2]["% COMISION"]);
                console.log(array1);   

                //Encontrar un valor ahí adentro
                search = (key, inputArray) => {
                    for (let i=0; i < inputArray.length; i++) {
                        if (inputArray[i].FIANZA == key) {
                            return inputArray[i];
                        }
                    }
                  }
                  let resultObject = search("97386", objetoBerkley);
                console.log(resultObject);
         });
         
        }
        
    }
    if(selectedFile2){ //Función que convierte SICAS en array de objetos
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile2);
        fileReader.onload = (event)=>{
         let data2 = event.target.result;
         let workbook2 = XLSX.read(data2,{type:"binary"});
         console.log(workbook2);
         workbook2.SheetNames.forEach(sheet => {
              objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //Nombre del array
              console.log(objetoSICAS);
              document.getElementById("jsondata2").innerHTML = JSON.stringify(objetoSICAS,undefined,4)
              
                        });
         
        }
    }
});