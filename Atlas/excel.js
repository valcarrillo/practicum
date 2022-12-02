//REFERENCIAS DEL CÓDIGO.
//https://levelup.gitconnected.com/how-to-convert-excel-file-into-json-object-by-using-javascript-9e95532d47c5
//convertir en EXCEL https://codepedia.info/javascript-export-html-table-data-to-excel

let selectedFile=[]; //En esta variable se almacenan los archivos leídos del Estado de cuenta. Es un array para que se puedan leer varios. Cada archivo es un valor del array.
let selectedFile2; //En esta variable se almacena el archivo de SICAS
let numarchivos=0; //Esta variable nos va a servir para registrar cuántos Estados de cuenta se leyeron.
console.log(window.XLSX); //Manda a la consola la lectura de xlsx.
document.getElementById('inputA').addEventListener("change", (event) => { //Cuando se suben los estados de cuenta, almacena cada uno en un valor del array
    const files = event.target.files;
    for (numarchivos=0; numarchivos < files.length; numarchivos++) {
        selectedFile[numarchivos] = event.target.files[numarchivos]; //selectedFile[0]=file[0]
     }// Esta función determinará el valor de numarchivos. Si son 5, numarchivos = 5 en este punto.
})
document.getElementById('inputSicas').addEventListener("change", (event) => {// Lee SICAS. Solo se puede subir un documento
    selectedFile2 = event.target.files[0];
}
)

let objetoSICAS; //Array de objetos en el que se van a guardar los datos de SICAS. Cada renglón será un objeto.
let objetoBerkley=[]; //Array de objetos en el que se van a guardar los datos de SICAS. Cada renglón será un objeto. A diferencia de la variable anterior, esta tiene =[] porque en la función siguiente se va a insertar cada objeto de todos los archivos. Si se lo quitas, no deja hacer el push.
//var fechareciente = new Date("2000-01-02"); // (YYYY-MM-DD) Variable que nos servirá para conocer la fecha más reciente.
//var fechaantigua = new Date("2100-01-01"); // (YYYY-MM-DD) Variable que nos servirá para conocer la fecha más antigua.
//const month = ["Nada","ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]; //Este arrar sirve para convertir los números en meses. Si pones month[date.getMonth()], va a dar el nombre de ese mes.
var reng_EC=0; //Esta variable guardará la cantidad de renglones que se contaron en el estado de cuenta, que es igual a la cantidad de objetos dentro de objetoBerkley. Este dato se desplegará en la página

document.getElementById('button').addEventListener("click", () => { //Cuando se da clic en el botón Comparar sucede los siguiente:
    if(selectedFile){ //Función para convertir Edo de Cuenta en array de objetos
        for(i=0; i<numarchivos; i++){ //ciclo que lee cada selected file de Estado de cuenta. Si no lo encuentra marcha error
            let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile[i]); //convierte el archivo den un binary string
        fileReader.onload = (event)=>{
         let data1 = event.target.result;
         let workbook1 = XLSX.read(data1,{type:"binary"});  //lee los resultados binarios como excel   
         workbook1.SheetNames.forEach(sheet => {
                                                           // EL RANGO ES LO GRANDE DEL ENCABEZADO
            objeto = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet], {range:13}); //objeto es el Nombre del array, y la función siguiente es la que lo convierte en un objeto de arrays. El primer campo es el documento, el segundo el tamaño de las filas a ignorar antes del encabezado
            console.log(objeto); //se imprime el objeto para ver que se lea bien
            for(var j=1; j<objeto.length-1; j++){// Cada objeto, es decir, cada renglón del estado de cuenta se inserta dentro de objetoBerkley. Esta función sirve para unir todos los renglones de varios Estados de cuenta en un solo objeto. Se pone -1 porque la última línea del edo de cuenta no se considera
                objetoBerkley.push(objeto[j]);
            }
        });
        }
    }
    console.log("objetoBerkley:");
    console.log(objetoBerkley); //Se imprime el array de objetos de Berkley completo.

         if(selectedFile2){ //Función que convierte SICAS en array de objetos
            let fileReader = new FileReader();
            fileReader.readAsBinaryString(selectedFile2);//convierte el archivo den un binary string
            fileReader.onload = (event)=>{
             let data2 = event.target.result;
             let workbook2 = XLSX.read(data2,{type:"binary"}); //lee los resultados binarios como excel   
             workbook2.SheetNames.forEach(sheet => {                   //No tiene range porque el rango antes del encabezado debe de ser 0.
                  objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //objeto es el Nombre del array, y la función siguiente es la que lo convierte en un objeto de arrays. El primer campo es el documento, el segundo el tamaño de las filas a ignorar antes del encabezado
                  console.log("objetoSICAS:"); //Se imprime el array de objetos de SICAS
                  console.log(objetoSICAS);
                  reng_SICAS=objetoSICAS.length; // Esta variable nos dirá cuántos renglones/objetos se leyeron en el archivo de SICAS. Este dato se desplegará en la página
                    //La tabla tiene atributo HIDDEN para que no se vea, pero ahí está.
                    //La tabla de diferencias es la que tiene el encabezado de toda la tabla, por lo que al unirse debe de ir primero siempre.
                    //En la tabla de diferencias regsitra cuando se encuentra una póliza pero hay diferencia de comisión
                  let tabladiferencias ="<table id='BerkleyFianzas' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' hidden><tr><th>SICAS</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th>BERKLEY</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>";
                  //En el renglón de arriba la tabla tiene [Sicas                                ][Berkley                              ]
                  //En el renglón de abajo la tabla tiene  [Nombre|Poliza|Endoso|Moneda|Serie....][Nombre|Poliza|Endoso|Moneda|Serie....]
                  //Se divide en dos para que no sea un renglón larguísimo.
                  tabladiferencias=tabladiferencias+"<tr><td>Nombre Asegurado o Fiado</td><td>Poliza</td><td>Endoso</td><td>Moneda</td><td>Serie Recibo</td><td>Tipo Cambio</td><td>PrimaNeta</td><td>Tipo Comision</td><td>Importe</td><td>% Participacion</td><td></td><td>Nombre Asegurado o Fiado</td><td>Poliza</td><td>Endoso</td><td>Moneda</td><td>Serie Recibo</td><td>% Comisión</td><td>Comisión</td><td>Tipo Cambio</td><td>Diferencia</td><td>Incidencia</td></tr>";
                  
                  let resultObject; //se guardan los resultados de hacer el search. Solo sirve para comprobar la operación. 
                  let berkley; //En esta variable solo se guarda un objeto de Berkley, es decir, un renglón, que es aquel que se va a buscar entre los renglones de SICAS en la función search()
                  var encontrar; //Es una bandera que nos ayudará a detectar si se encontró una póliza o no, y en caso de que no, se registrará en no encontrados.
                  let tablanoencontrados="";    //Aqui se guardan los datos no encontrados. Se unen al final para tener más orden.
                  let tablaiguales=""; //Aquí se almacenan los que no tienen diferencias. Solo es por estilo.

                  /************************************/
                  /***********FUNCIÓN SEARCH**********/
                  //Esta función se llama más abajo//
                  //Encontrar una póliza de Berkley dento de todos los de SICAS

                  /*
                    Para sacar el valor de un arrar de objetos se hace de la manera siguiente:
                    NombredelArray[valor].Nombredelcampo para nombres simples o:
                    NombredelArray[valor]["Nombre del campo"] si tiene nombres complejos o con caracteres especiales.
                    Ejemplo: ArraySICAS[0].Poliza  o ArraySICAS[0]["PrimaNeta"]

                    En caso de que fuera un solo objeto, se quita el valor
                    NombredelArray.Nombredelcampo para nombres simples o:
                    NombredelArray["Nombre del campo"] si tiene nombres complejos o con caracteres especiales.
                    Ejemplo: berkley.COMISIONES o berkley["PRIMA NETA"]
                  */

                  search = (key, endo, ArraySICAS) => { //Se manda la póliza (key), inclusión (inclu) y endoso (endo) de un objeto Berkley con todo el Array de Sicas
                      for (let i=0; i < ArraySICAS.length; i++) { //Este ciclo comparará el renglón de Berkley con cada renglón de SICAS
                        encontrar=0; //Encontrar =0 significa que no se ha encontrado nada
                        var pol = ArraySICAS[i].Poliza.toString().split('-');  //Divide la póliza de SICAS por '-' . La posición 2 es la fianza y la 3 es la inclusión
                        SICASpoliza = pol[3]; //póliza de sicas
                        //SICASinclusion= pol[3]; //inclusión de sicas
                       /* if(typeof  SICASinclusion === 'undefined'){ //Si SICAS no tiene inclusión, se toma como inclusión 0
                            SICASinclusion=0;
                        }else{
                            SICASinclusion= + SICASinclusion; //Si la inclusión existe, se sumará uno al valor, porque Berkley registra la inclusión 0 como 1
                        }*/
                        numpoliza = +SICASpoliza; //se convierte la póliza en un número para borrar los 0s anteriores
                        var SICASendoso=ArraySICAS[i].Endoso; //Aquí se toma en endoso de SICAS
                        if(typeof SICASendoso === ' '){
                            SICASendoso= 0; //Si el endoso no está en SICAS entonces se registra como 1
                        }else{
                            SICASendoso= +SICASendoso +1; //Si hay endoso, se le suma 1, porque Berkley registra el primer endoso como 1
                        }
                        //Busqueda
                          if (numpoliza == key) { //póliza SICAS == póliza Berkley)
                            //if (SICASinclusion == inclu) { //inclusión SICAS == inclusión Berkley?
                                if (SICASendoso == endo) { //endoso SICAS == endoso Berkley?
                                    encontrar=1; //Si los tres campos anteriores coinciden, se encontó la póliza. Entonces la bandera es 1

                                //FUNCIONES QUE REDONDEAN LOS VALORES Y MULTIPLICAN POR EL TIPO DE CAMBIO
                                importeSicas=Math.round(ArraySICAS[i]["Importe"]*ArraySICAS[i].TC*100)/100;// Esto es para que solo cuente los primeros dos decimales
                                importeBerkley=Math.round(berkley["Abono"]*berkley["Moneda"]*100)/100;
                                    //Esta operación saca la diferencia de los datos anteriores
                                    var diferencia= Math.round((importeBerkley -importeSicas)*100)/100;
                                    var tipodif; //aquí se registrará en dónde se encuentra la diferencia en cado de que exita.
                                    if(berkley["Prima neta"] !=ArraySICAS[i]["PrimaNeta"]){
                                        tipodif="Prima Neta y % Comisión"; //La diferencia estuvo en la prima neta y % de comisión
                                    }else if(berkley["Prima neta"] !=ArraySICAS[i]["PrimaNeta"]){
                                        tipodif="Prima Neta"; //diferencia en prima neta
                                    }else if(berkley["Moneda"] !=ArraySICAS[i].TC){
                                            tipodif="Tipo de Cambio"; //la diferencia está en el tipo de cambio
                                    }else{
                                        tipodif="Total Comisión"; //En caso de no ser la diferencia en niguna de las anteriores, simplemente se registra como diferencia en Total de comisión
                                    }
                                if(importeBerkley != importeSicas){ //Si el importe es diferente, es decir, si hay diferencia o la resta en diferente a 0, se registra en tabal de diferencias. El campo Póliza incluye la póliza y la inclusión"</td><td>"+berkley.COMISIONES+"</td><td></td><td>"+diferencia+"</td><td>"+tipodif+"</td></tr>";
                                    tabladiferencias=tabladiferencias+"<tr><td>"+ArraySICAS[i]["Nombre Asegurado o Fiado"]+"</td><td>"+ArraySICAS[i].Poliza+"</td><td>"+ArraySICAS[i].Endoso+"</td><td>"+ArraySICAS[i].Moneda+"</td><td>'"+ArraySICAS[i].Serie+"'</td><td>"+ArraySICAS[i].TC+"</td><td>"+ArraySICAS[i].PrimaNeta+"</td><td>"+ArraySICAS[i]["Tipo Comision"]+"</td><td>"+ArraySICAS[i].Importe+"</td><td>"+ArraySICAS[i]["% Participacion"]+"</td><td></td><td>"+berkley["Póliza"]+"</td><td>";
                                }else{ //Si la resta es == 0, es decir, son iguales, se registra en la tabla de iguales
                                    tablaiguales=tablaiguales+"<tr><td>"+ArraySICAS[i]["Nombre Asegurado o Fiado"]+"</td><td>"+ArraySICAS[i].Poliza+"</td><td>"+ArraySICAS[i].Endoso+"</td><td>"+ArraySICAS[i].Moneda+"</td><td>'"+ArraySICAS[i].Serie+"'</td><td>"+ArraySICAS[i].TC+"</td><td>"+ArraySICAS[i].PrimaNeta+"</td><td>"+ArraySICAS[i]["Tipo Comision"]+"</td><td>"+ArraySICAS[i].Importe+"</td><td>"+ArraySICAS[i]["% Participacion"]+"</td><td></td><td>"+berkley["Póliza"]+"</td><td>"+diferencia+"</td><td></td></tr>";
                                }
                              return ArraySICAS[i]; //Se regresa el objeto de SICAS que se encontró.
                               // }
                            }
                          }
                      }
                      if(encontrar==0){ // Si al terminar la comparación no encontró la póliza, encontrar==0, entonces se regsitra en no encontrados
                      tablanoencontrados= tablanoencontrados+"<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>"+berkley["Póliza"]+"</td><td></td><td></td><td>NO SE ENCONTRÓ</td></tr>";
                        return ArraySICAS;
                        }
                      encontrar=0; // Se reestablece la bandera
                    }
                    
                    //#################################################################
                    //###########FUNCIÓN QUE MANDA A LLAMAR LA BÚSQUEDA################ 
                    //#################################################################
                   
                        try { //Si el primer objeto de SICAS tiene póliza al igual que el primer objeto de Berkley, entonces hace la comparación.
                                //Si no encuentra estos campos en los primeros objetos, mandará que no se pudo leer el archivo
                             if(typeof objetoSICAS[0].Poliza === 'undefined' || !(objetoBerkley[0].hasOwnProperty('Póliza'))){
                                document.getElementById("jsondata").innerHTML = "No se pudo leer el documento. Revise haber adjuntado el correcto.";
                              }else{
                            reng_EC=objetoBerkley.length; //Se cuentan los renglones/objetos de Berkley

                            for(var j=0; j<objetoBerkley.length-1; j++){ //Ciclo que va a mandar cada póliza de Berkley a la función search()
                                
                                //Busca la fecha más antigua y la más reciente en el Estado de Cuenta de Berkley para el nombre del xlsx.
                                    //var fechas =objetoBerkley[j]["FECHA APLICACION"]; //Nombre del campo de fecha
                                    //const [dia, mes, anio] = fechas.toString().split('/'); //divide por '/'. El primero es el día, luego el mes y año
                                   // const fecha1 = new Date(+anio, +mes - 1, +dia); //Convierte los datos anteriores en una fecha. A mes se le resta 1 porque Date va de 0 a 11
                                   // if(fecha1>fechareciente){ //Se compara si la fecha es más reciente. En caso de ser así, cambia el valor de la fecha más reciente, que en principio es el "2000-01-02"
                                   /*     fechareciente=fecha1;
                                    }
                                    if(fecha1<fechaantigua){ //Se compara si la fecha es más antigua. En caso de ser así, cambia el valor de la fecha más antigua, que en principio es el "2100-01-01";
                                        fechaantigua=fecha1;
                                    }*/
                                
                                berkley=objetoBerkley[j]; //berkley es donde se guarda SOLO UN objeto/renglón de Berkley, el que se va a mandar a buscar
                                poliza=objetoBerkley[j].Póliza //aquí se guarda la póliza de berkley
                                //inclusion=objetoBerkley[j].INCLUSION //aqui se guarda la inclusión
                                movimiento=objetoBerkley[j].Endoso //aquí el movimiento o endoso. Ver terminología
                            
                                //Manda a llamar a la función de búsqueda y el resultado lo pone en resultObject.
                                                //póliza, inclusión, endoso o movimiento, todos los objetos de SICAS
                            resultObject = search(poliza, movimiento, objetoSICAS);
                            //En el renglón de abajo se manda a la página la cantidad de renglones que se encontraron en cada documento. Si se subieron varios estados de cuenta de Berkley, los suma todos
                            document.getElementById("numregistros").innerHTML = "Renglones Estado de Cuenta: "+reng_EC+"\nRenglones SICAS: "+reng_SICAS+"\n";
                            //Manda la tabla a la página, pero no aparecerá por el atributo HIDDEN. Primero la tabla de diferencias porque tiene el encabezado, luego los iguales, luego no encontrados y al final las fechas.
                            document.getElementById("jsondata").innerHTML = tabladiferencias+tablaiguales+tablanoencontrados; // DEL "+fechaantigua.getDate()+" "+month[+fechaantigua.getMonth()+1]+" "+fechaantigua.getFullYear()+" AL "+fechareciente.getDate()+" "+month[+fechareciente.getMonth()+1]+" "+fechareciente.getFullYear();;//+month[messicas]+" Año: "+aniosicas; //Se manda la tabla pero no se va a ver porque tiene HIDDEN
                        }
                            ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente. Se manda el parámetro type.
                    }                          
                } catch (error) { //Si hay un error aquí se muestra.
                            document.getElementById("jsondata").innerHTML = "Algo salió mal al leer el documento. Revise que el encabezado tenga el formato correcto. Error: "+error;
                          }
            }
            ); 
            }
        } else{//En caso de no adjuntarse nada en SICAS aquí manda el error
             document.getElementById("jsondata").innerHTML = "No se adjuntó nada en SICAS";
        }
    }else{//En caso de no adjuntarse nada en Berkley se manda el error. Primero revisará que haya en Berkley y luego en SICAS
        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de Berkley";
    }
    
});
//fn: filename, dl: download, type:xlsx que se manda al llamar la función
function ExportToExcel(type, fn, dl) {// función que convierte a excel
    var elt = document.getElementById('BerkleyFianzas'); //Nombre de la tabla: 'BerkleyFianzas'
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    //Nombre del documento
    var nombre ='CONCILIACIÓN ATLAS FIANZAS DEL '+".";
    return dl ? //Va a tratar de forzar un client-side download.
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || (nombre + (type || 'xlsx')));
}
