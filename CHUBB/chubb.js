//REFERENCIAS DEL CÓDIGO.
//https://levelup.gitconnected.com/how-to-convert-excel-file-into-json-object-by-using-javascript-9e95532d47c5
//convertir en EXCEL https://codepedia.info/javascript-export-html-table-data-to-excel

let selectedFile=[]; //En esta variable se almacenan los archivos leídos del Estado de cuenta. Es un array para que se puedan leer varios. Cada archivo es un valor del array.
let selectedFile2; //En esta variable se almacena el archivo de SICAS
let numarchivos=0; //Esta variable nos va a servir para registrar cuántos Estados de cuenta se leyeron.
console.log(window.XLSX);//Manda a la consola la lectura de xlsx.
document.getElementById('inputA').addEventListener("change", (event) => { //Cuando se suben los estados de cuenta, almacena cada uno en un valor del array
    const files = event.target.files;
    for (numarchivos=0; numarchivos < files.length; numarchivos++) {
        selectedFile[numarchivos] = event.target.files[numarchivos]; //selectedFile[0]=file[0]
     } // Esta función determinará el valor de numarchivos. Si son 5, numarchivos = 5 en este punto.
})
document.getElementById('inputSicas').addEventListener("change", (event) => {// Lee SICAS. Solo se puede subir un documento
    selectedFile2 = event.target.files[0];
}
)
let objetoSICAS; //Array de objetos en el que se van a guardar los datos de SICAS. Cada renglón será un objeto.
let objetoCHUBB=[]; //Array de objetos en el que se van a guardar los datos de SICAS. Cada renglón será un objeto. A diferencia de la variable anterior, esta tiene =[] porque en la función siguiente se va a insertar cada objeto de todos los archivos. Si se lo quitas, no deja hacer el push.
var fechareciente = new Date("2000-01-02"); // (YYYY-MM-DD) Variable que nos servirá para conocer la fecha más reciente.
var fechaantigua = new Date("2100-01-01"); // (YYYY-MM-DD) Variable que nos servirá para conocer la fecha más antigua.
const month = ["Nada","ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]; //Este arrar sirve para convertir los números en meses. Si pones month[date.getMonth()], va a dar el nombre de ese mes.
var reng_EC=0;  //Esta variable guardará la cantidad de renglones que se contaron en el estado de cuenta, que es igual a la cantidad de objetos dentro de objetoCHUBB. Este dato se desplegará en la página

document.getElementById('button').addEventListener("click", () => { //Cuando se da clic en el botón Comparar sucede los siguiente:
    //Este es un objeto vacío que se va insertar en objeto CHUBB como el primer objeto. 
    /*Este no va a aparecer en los resultados pero se creó porque para transformar los objetos leídos en CHUBB
    se leen a partir del anterior, de esta manera: objetoCHUBB[j-1]. Para que pudiera leer el primero
    se agrega un objeto antes. Sino marca error por objetoCHUBB[-1] está fuera de los límites
    */
    jsonObj={"Asegurado": "NO SIRVE SOLO PARA PRUEBA", "ClaveId": "AAA", "PolizaId": "00000000","Endoso": "00000", "Recibo": "0", "TotalRecibo": "0","PNetaMto": "000", "ComisionMto": "00000", "ComisionSobreRecargoMto": "000", "TipoMov": "0"};
    objetoCHUBB.push(jsonObj);
    jsonObj={};//Se vacía esta variable que nos servirá para insertar objetos.
    if(selectedFile){ //Función para convertir Edo de Cuenta en array de objetos
        for(i=0; i<numarchivos; i++){ //ciclo que lee cada selected file de Estado de cuenta. Si no lo encuentra marca error
            let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile[i]);//convierte el archivo den un binary string
        fileReader.onload = (event)=>{
         let data1 = event.target.result;
         let workbook1 = XLSX.read(data1,{type:"binary"});   //lee los resultados binarios como excel      
         workbook1.SheetNames.forEach(sheet => {
                                                           // {range:1} EL RANGO ES LO GRANDE DEL ENCABEZADO, pero aquí no está porque el rango es 0                                        
                objeto = XLSX.utils.sheet_to_row_object_array(workbook1.Sheets[sheet]); //objeto es el Nombre del array, y la función siguiente es la que lo convierte en un objeto de arrays.
                for(var j=1; j<objeto.length; j++){
                    objetoCHUBB.push(objeto[j]);// Cada objeto, es decir, cada renglón del estado de cuenta se inserta dentro de objetoCHUBB. Esta función sirve para unir todos los renglones de varios Estados de cuenta en un solo objeto. Se pone -1 porque la última línea del edo de cuenta no se considera
                }
         });
        }
        }
        console.log("objetoCHUBB:");
         console.log(objetoCHUBB);//Se imprime el array de objetos de CHUBB completo.
         
         if(selectedFile2){ //Función que convierte SICAS en array de objetos
            let fileReader = new FileReader();
            fileReader.readAsBinaryString(selectedFile2);//convierte el archivo den un binary string
            fileReader.onload = (event)=>{
             let data2 = event.target.result;
             let workbook2 = XLSX.read(data2,{type:"binary"}); //lee los resultados binarios como excel  
             workbook2.SheetNames.forEach(sheet => {             //No tiene range porque el rango antes del encabezado debe de ser 0.
                  objetoSICAS = XLSX.utils.sheet_to_row_object_array(workbook2.Sheets[sheet]); //objeto es el Nombre del array, y la función siguiente es la que lo convierte en un objeto de arrays. El primer campo es el documento, el segundo el tamaño de las filas a ignorar antes del encabezado
                  console.log(objetoSICAS);//Se imprime el array de objetos de SICAS
                  var reng_SICAS=objetoSICAS.length;// Esta variable nos dirá cuántos renglones/objetos se leyeron en el archivo de SICAS. Este dato se desplegará en la página
                  //La tabla tiene atributo HIDDEN para que no se vea, pero ahí está.
                  //La tabla de diferencias es la que tiene el encabezado de toda la tabla, por lo que al unirse debe de ir primero siempre.
                  //En la tabla de diferencias regsitra cuando se encuentra una póliza pero hay diferencia de comisión
                  let tabladiferencias ="<table id='CHUBB' width='80%' border='1' cellpadding='0' cellspacing='0' bordercolor='#0000001' hidden><tr><th>SICAS</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th>CHUBB</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>";
                  //En el renglón de arriba la tabla tiene [Sicas                                ][CHUBB                              ]
                  //En el renglón de abajo la tabla tiene  [Nombre|Poliza|Endoso|Moneda|Serie....][Nombre|Poliza|Endoso|Moneda|Serie....]
                  //Se divide en dos para que no sea un renglón larguísimo.
                  tabladiferencias=tabladiferencias+"<tr><td>Nombre Asegurado o Fiado</td><td>Poliza</td><td>Endoso</td><td>Moneda</td><td>Serie Recibo</td><td>Tipo Cambio</td><td>PrimaNeta</td><td>Tipo Comision</td><td>Importe</td><td>% Participacion</td><td></td><td>Nombre Asegurado o Fiado</td><td>Poliza</td><td>Endoso</td><td>Moneda</td><td>Serie Recibo</td><td>% Comisión</td><td>Comisión</td><td>Tipo Cambio</td><td>Diferencia</td><td>Incidencia</td></tr>";
                  
                  let resultObject; //se guardan los resultados de hacer el search. Solo sirve para comprobar la operación. 
                  let noencontrados=[]; //En esta variable se guardan como arreglo aquellos objetos que no fueron encontrados en la función search() para mandarse a la función search2()
                  var encontrar; //Es una bandera que nos ayudará a detectar si se encontró una póliza o no, y en caso de que no, se registrará en no encontrados.
                  let tablanoencontrados="";   //Aqui se guardan los datos no encontrados. Se unen al final para tener más orden.
                  let tablaiguales=""; //Aquí se almacenan los que no tienen diferencias. Solo es por estilo.
                  let tablatipo6="" // En esta tabla se registrarán todas las comisiones por internet.

                  /************************************/
                  /***********FUNCIONES SEARCH**********/

                  /* HAY 2 FUNCIONES DE SEARCH
                  
                  SEARCH(): busca las pólizas de CHUBB entre las de SICAS con endoso. 
                  SEARCH2(): busca las pólizas no encontradas en la función anterior entre las de SICAS sin endoso. En esta última no se comparan endosos.
                  
                  Se hicieron dos funciones porque CHUBB le pone endosos a todo, mientras que Hoken solo a partir del segundo movimiento.
                  */
                  //Estas funciones se llaman más abajo//
                  //Encontrar una póliza de CHUBB dento de todos los de SICAS

                  /*
                    Para sacar el valor de un arrar de objetos se hace de la manera siguiente:
                    NombredelArray[valor].Nombredelcampo para nombres simples o:
                    NombredelArray[valor]["Nombre del campo"] si tiene nombres complejos o con caracteres especiales.
                    Ejemplo: ArraySICAS[0].Poliza  o ArraySICAS[0]["PrimaNeta"]

                    En caso de que fuera un solo objeto, se quita el valor
                    NombredelArray.Nombredelcampo para nombres simples o:
                    NombredelArray["Nombre del campo"] si tiene nombres complejos o con caracteres especiales.
                    Ejemplo: CHUBB.COMISIONES o CHUBB["PRIMA NETA"]
                  */

                 /***********FUNCIÓN SEARCH**********/
                  search = (poliza, CHUBB, ArraySICAS) => { //Se manda la póliza de CHUBB, el objeto de CHUBB y todos los objetos de SICAS que tienen endoso
                    for (let i=0; i < ArraySICAS.length; i++) {//Este ciclo comparará el renglón de CHUBB con cada renglón de SICAS
                        encontrar=0; //Encontrar =0 significa que no se ha encontrado nada
                        var SICASendoso=ArraySICAS[i].Endoso; //Aquí se toma en endoso de SICAS
                        if (ArraySICAS[i].Poliza == poliza) { //poliza SICAS == poliza CHUBB?
                            if (SICASendoso == CHUBB.Endoso) {//endoso SICAS == endoso CHUBB?
                               if (ArraySICAS[i].Serie == CHUBB.Serie) {// serie SICAS == serie CHUBB?
                                    if (ArraySICAS[i]["Tipo Comision"] == CHUBB["Tipo Comisión"]) { // tipo comisión SICAS == tipo comisión CHUBB?
                                        encontrar=1; //Si los cuatro campos anteriores coinciden, se encontó la póliza. Entonces la bandera es 1
                                
                                        //FUNCIÓN QUE REDONDEA SICAS Y MULTIPLICA POR TIPO DE CAMBIO
                                importeSicas=Math.round(ArraySICAS[i]["Importe"]*ArraySICAS[i].TC*100)/100;// Esto es para que solo cuente los primeros dos decimales
                                    //Diferencia en los totales de comisión
                                    var diferencia= Math.round((CHUBB.Importe -importeSicas)*100)/100;
                                    var tipodif; //aquí se registrará en dónde se encuentra la diferencia en cado de que exita.
                                    if(CHUBB.PNetaMto !=ArraySICAS[i]["PrimaNeta"] && CHUBB["% Comision"] !=ArraySICAS[i]["% Participacion"]){
                                        tipodif="Prima Neta y % Comisión"; //La diferencia estuvo en la prima neta y % de comisión
                                    }else if(CHUBB.PNetaMto !=ArraySICAS[i]["PrimaNeta"]){
                                        tipodif="Prima Neta"; //diferencia en prima neta
                                    }else if(CHUBB["% Comision"] !=ArraySICAS[i]["% Participacion"]){
                                        tipodif="% Comisión"; //diferencia en % de comisión
                                    }else{
                                        tipodif="Total Comisión"; //En caso de no ser la diferencia en niguna de las anteriores, simplemente se registra como diferencia en Total de comisión
                                    }

                                if(CHUBB.Importe!= importeSicas){//Si el importe es diferente, es decir, si hay diferencia o la resta en diferente a 0, se registra en tabal de diferencias. El campo Póliza incluye la póliza y la inclusión
                                    tabladiferencias=tabladiferencias+"<tr><td>"+CHUBB.Asegurado+"</td><td>"+ArraySICAS[i].Poliza+"</td><td>"+ArraySICAS[i].Endoso+"</td><td>"+ArraySICAS[i].Moneda+"</td><td>'"+ArraySICAS[i].Serie+"'</td><td>"+ArraySICAS[i].TC+"</td><td>"+ArraySICAS[i].PrimaNeta+"</td><td>"+ArraySICAS[i]["Tipo Comision"]+"</td><td>"+ArraySICAS[i].Importe+"</td><td>"+ArraySICAS[i]["% Participacion"]+"</td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td>"+diferencia+"</td><td>"+tipodif+"</td></tr>";
                                }else{ //Si la resta es == 0, es decir, son iguales, se registra en la tabla de iguales
                                    tablaiguales=tablaiguales+"<tr><td>"+CHUBB.Asegurado+"</td><td>"+ArraySICAS[i].Poliza+"</td><td>"+ArraySICAS[i].Endoso+"</td><td>"+ArraySICAS[i].Moneda+"</td><td>'"+ArraySICAS[i].Serie+"'</td><td>"+ArraySICAS[i].TC+"</td><td>"+ArraySICAS[i].PrimaNeta+"</td><td>"+ArraySICAS[i]["Tipo Comision"]+"</td><td>"+ArraySICAS[i].Importe+"</td><td>"+ArraySICAS[i]["% Participacion"]+"</td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td>"+diferencia+"</td><td></td></tr>";
                                }
                              return ArraySICAS; //Se regresa el objeto de SICAS que se encontró.
                            }
                            
                            }
                           }
                          }
                    }
                      if(encontrar==0){ // Si al terminar la comparación no encontró la póliza, encontrar==0, entonces se manda al arreglo de no encontrados
                        noencontrados.push(CHUBB); //Este arreglo se manda a search2()
                        return CHUBB;
                        }
                      encontrar=0; // Se reestablece la bandera
                    }

                    /***********FUNCIÓN SEARCH2**********/
                  search2 = (poliza, CHUBB, ArraySICAS) => { //Se manda la póliza de CHUBB, el objeto de CHUBB y todos los objetos de SICAS que NO tienen endoso
                    for (let i=0; i < ArraySICAS.length; i++) {//Este ciclo comparará el renglón de CHUBB con cada renglón de SICAS
                        encontrar=0;//Encontrar =0 significa que no se ha encontrado nada
                        if (ArraySICAS[i].Poliza == poliza) {//poliza SICAS == poliza CHUBB?
                            //No se compara endoso
                            if (ArraySICAS[i].Serie == CHUBB.Serie) {// serie SICAS == serie CHUBB?
                                if (ArraySICAS[i]["Tipo Comision"] == CHUBB["Tipo Comisión"]) { // tipo comisión SICAS == tipo comisión CHUBB?
                                encontrar=1; //Si los cuatro campos anteriores coinciden, se encontó la póliza. Entonces la bandera es 1

                            //FUNCIÓN QUE REDONDEA Y MULTIPLICA POR EL TIPO DE CAMBIO
                            importeSicas=Math.round(ArraySICAS[i]["Importe"]*ArraySICAS[i].TC*100)/100;// Esto es para que solo cuente los primeros dos decimales
                                 //Diferencia en los totales de comisión
                                var diferencia= Math.round((CHUBB.Importe -importeSicas)*100)/100;
                                var tipodif; //aquí se registrará en dónde se encuentra la diferencia en cado de que exita.
                                if(CHUBB.PNetaMto !=ArraySICAS[i]["PrimaNeta"] && CHUBB["% Comision"] !=ArraySICAS[i]["% Participacion"]){
                                    tipodif="Prima Neta y % Comisión"; //La diferencia estuvo en la prima neta y % de comisión
                                }else if(CHUBB.PNetaMto !=ArraySICAS[i]["PrimaNeta"]){
                                    tipodif="Prima Neta"; //diferencia en prima neta
                                }else if(CHUBB["% Comision"] !=ArraySICAS[i]["% Participacion"]){
                                    tipodif="% Comisión"; //diferencia en % de comisión
                                }else{
                                    tipodif="Total Comisión"; //En caso de no ser la diferencia en niguna de las anteriores, simplemente se registra como diferencia en Total de comisión
                                }
                             
                            if(CHUBB.Importe != importeSicas){//Si el importe es diferente, es decir, si hay diferencia o la resta en diferente a 0, se registra en tabal de diferencias. El campo Póliza incluye la póliza y la inclusión
                                tabladiferencias=tabladiferencias+"<tr><td>"+CHUBB.Asegurado+"</td><td>"+ArraySICAS[i].Poliza+"</td><td>"+ArraySICAS[i].Endoso+"</td><td>"+ArraySICAS[i].Moneda+"</td><td>'"+ArraySICAS[i].Serie+"'</td><td>"+ArraySICAS[i].TC+"</td><td>"+ArraySICAS[i].PrimaNeta+"</td><td>"+ArraySICAS[i]["Tipo Comision"]+"</td><td>"+ArraySICAS[i].Importe+"</td><td>"+ArraySICAS[i]["% Participacion"]+"</td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td>"+diferencia+"</td><td>"+tipodif+"</td></tr>";
                            }else{//Si la resta es == 0, es decir, son iguales, se registra en la tabla de iguales
                                tablaiguales=tablaiguales+"<tr><td>"+CHUBB.Asegurado+"</td><td>"+ArraySICAS[i].Poliza+"</td><td>"+ArraySICAS[i].Endoso+"</td><td>"+ArraySICAS[i].Moneda+"</td><td>'"+ArraySICAS[i].Serie+"'</td><td>"+ArraySICAS[i].TC+"</td><td>"+ArraySICAS[i].PrimaNeta+"</td><td>"+ArraySICAS[i]["Tipo Comision"]+"</td><td>"+ArraySICAS[i].Importe+"</td><td>"+ArraySICAS[i]["% Participacion"]+"</td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+ CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td>"+diferencia+"</td><td></td></tr>";
                            }
                          return ArraySICAS; //Se regresa el objeto de SICAS que se encontró.
            
                                }
                        }
                          }
                    }
                      if(encontrar==0){ // Si al terminar la comparación no encontró la póliza, encontrar==0, entonces:
                        //En caso de ser tipo comisón Bono por Internet, se registra en una tabla aparte
                        if( CHUBB["Tipo Comisión"]=="Bono Internet/ Comisión por derechos"){
                            tablatipo6=tablatipo6+"<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>"+CHUBB["Tipo Comisión"] +"</td><td></td><td></td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+ CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td></td><td>NO SE ENCONTRÓ</td></tr>";
                        }else{// si no se registra en no encontradps
                        tablanoencontrados= tablanoencontrados+"<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>"+CHUBB["Tipo Comisión"] +"</td><td></td><td></td><td></td><td>"+CHUBB.Asegurado+"</td><td>"+poliza+"</td><td>"+CHUBB.Endoso+"</td><td></td><td>'"+ CHUBB.Serie+"'</td><td>"+CHUBB["% Comision"]+"</td><td>"+CHUBB.Importe+"</td><td></td><td></td><td>NO SE ENCONTRÓ</td></tr>";
                        return CHUBB;
                        }
                        }
                      encontrar=0; // Se reestablece la bandera
                    }

                /************************************/
                  /***********TRANSFORMACIÓN DE LOS DATOS**********/
                  /* CHUBB desglosa las pólizas en incisos, mientras que SICAS no,
                  por eso se transformarán los datos, sumando las primas netas y comisiones de pólizas iguales */
                    var pnetamto =0; //Aquí se guardará la suma de las primas netas de una póliza desglosada
                    var comisionmto=0, comisionrecargo=0,  comisionespecial=0; //Aquí se guardan las sumas de las comisiones de una póliza
                    var claveanterior="", polizaanterior="", endosoanterior=""; //Sirve para conocer la clave, póliza y endoso anterior y saber si es igual con la que se esta trabajando, para ver si se suman o no
                    var registers=0; //numero de registros
                    let objetoNuevochubb=[]; //En este arreglo se guardarán los objetos de CHUBB transformados.
                    let jsonObj;
                    //SICAS se va a dividir en dos, en los que tienen endoso y los que no
                    let Sicassinendoso=[]; //Arreglo de objetos de SICAS sin endoso
                    let Sicasconendoso=[]; //Arreglo de objetos de SICAS con endoso
                
                try {//Si el primer objeto de SICAS tiene póliza al igual que el primer objeto de CHUBB, entonces hace la comparación.
                    //Si no encuentra estos campos en los primeros objetos, mandará que no se pudo leer el archivo
                    if(typeof objetoSICAS[0].Poliza === 'undefined' || !(objetoCHUBB[0].hasOwnProperty('PolizaId'))){
                        document.getElementById("jsondata").innerHTML = "No se pudo leer el documento. Revise haber adjuntado el correcto.";
                    }else{
                        for(var j=0; j<objetoSICAS.length; j++){ //Se toma cada renglón de SICAS.
                            if(objetoSICAS[j].Endoso===''){//Si no tiene endoso, se registra como 0 y se inserta en SICAS sin endoso
                                objetoSICAS[j].Endoso='0';
                                Sicassinendoso.push(objetoSICAS[j]);
                                
                            }else{
                                Sicasconendoso.push(objetoSICAS[j]); //Si sí tiene endoso, se inserta en SICAS con endoso
                                
                            }

                        }
                        console.log("Sicassinendoso"); //Se imprime sicas sin endoso
                        console.log(Sicassinendoso);
                        console.log("Sicasconendoso"); //Se imprime sicas con endoso
                        console.log(Sicasconendoso);
                           //Este es un objeto vacío que se va insertar en objeto CHUBB como el primer objeto. 
                        /*Este no va a aparecer en los resultados pero se creó porque para transformar los objetos leídos en CHUBB
                        se leen a partir del anterior, de esta manera: objetoCHUBB[j-1]. Para que pudiera leer el primero
                        se agrega un objeto antes. Sino marca error por objetoCHUBB[-1] está fuera de los límites
                        */
                        jsonObj={"Asegurado": "NO SIRVE SOLO PARA PRUEBA", "ClaveId": "AAA", "PolizaId": "00000000","Endoso": "00000", "Recibo": "0", "TotalRecibo": "0","PNetaMto": "000", "ComisionMto": "00000", "ComisionSobreRecargoMto": "000", "TipoMov": "0"};
                        objetoCHUBB.push(jsonObj);
                        jsonObj={};
                        console.log(objetoCHUBB);
                        for(var j=1; j<objetoCHUBB.length; j++){ //Ciclo que va a transformar cada poliza de CHUBB
                            
                            if(objetoCHUBB[j].hasOwnProperty('PolizaId')){ //Si tiene Póliza
                                //Busca la fecha más antigua y la más reciente en SICAS para el nombre del xlsx. 
                                if(objetoCHUBB[j].hasOwnProperty('Fecha')){       //Nombre del campo de fecha                
                                    var fechas =objetoCHUBB[j].Fecha.split(' '); //Se divide por un espacio porque CHUBB tiene fecha y hora separada por un espacio
                                    const [mes, dia, anio] = fechas[0].split('/'); //Se divide la fecha en mes, dia y año
                                    const fecha1 = new Date(+anio, +mes - 1, +dia); //Convierte los datos anteriores en una fecha. A mes se le resta 1 porque Date va de 0 a 11
                                    if(fecha1>fechareciente){//Se compara si la fecha es más reciente. En caso de ser así, cambia el valor de la fecha más reciente, que en principio es el "2000-01-02"
                                        fechareciente=fecha1;
                                    }
                                    if(fecha1<fechaantigua){//Se compara si la fecha es más antigua. En caso de ser así, cambia el valor de la fecha más antigua, que en principio es el "2100-01-01";
                                        fechaantigua=fecha1;
                                    }
                                }

                            //console.log(objetoCHUBB[j]);
                            CHUBB=objetoCHUBB[j]; //CHUBB es donde se guarda SOLO UN objeto/renglón de CHUBB, el que se va a mandar a buscar
                            //Las pólizas en CHUBB tienen dos partes, Clave Id y Póliza
                            clave=objetoCHUBB[j].ClaveId; 
                            if(clave != null){ //Si la clave NO es nula, se eliminan los espacios para evitar problemas de comparación
                                clave = clave.toString().replace(/\s/g, '');
                            }else{ //Sino, la clave se considera como nula
                                clave='';
                                objetoCHUBB[j].ClaveId=clave;
                            }
                            if(!(objetoCHUBB[j].hasOwnProperty('Asegurado'))){ //Si CHUBB no tiene asegurado, se considera como un espacio en blanco. Esto sirve en especial para las comisiones por internet
                                objetoCHUBB[j].Asegurado=' ';
                            }
                            poliza=objetoCHUBB[j].PolizaId; //Póliza de CHUBB
                            poliza = poliza.toString().replace(/\s/g, ''); //Se eliminan los espacios en blanco
                            endoso=objetoCHUBB[j].Endoso;
                            recibo=objetoCHUBB[j].Recibo;

                            //SE VA A CREAR UN NUEVO ARRAY DE OBJETOS JSON SUMANDO LAS PÓLIZAS QUE SEAN IGUALES
                            if(clave==claveanterior && poliza==polizaanterior && endoso==endosoanterior){ //Si la clave, la poliza y el endoso son iguales
                                pnetamto=pnetamto+objetoCHUBB[j].PNetaMto; //Se suma la prima neta
                                comisionmto=comisionmto+objetoCHUBB[j].ComisionMto; //Se suma la comisión base o neta monto
                                comisionrecargo=comisionrecargo+objetoCHUBB[j].ComisionSobreRecargoMto; // Se suma la comisón sobre recargos
                                comisionespecial=comisionespecial+objetoCHUBB[j].Comision2 //Se suman las comisiones especiales
                            }else{
                            // Si no es igual entonces se van a insertar los datos guardados arriba en el array, es decir, los datos de chubb[j-1]
                                if(pnetamto!=0){//CHUBB no tiene % de comisión, entonces aquí se calcula
                                    porcentajecomision=Math.round((comisionmto/pnetamto)*100);
                                }else{
                                    porcentajecomision="";
                                }
                                //Si CHUBB tiene los campos de Recibo y Total registo, se guardan en un campo llamado serie para compararse con la serie de SICAS
                                //Ejemplo: Recibo=1, Total registros=2  se guarda como Serie=001/002
                                if(objetoCHUBB[j-1].hasOwnProperty('Recibo') && objetoCHUBB[j-1].hasOwnProperty('TotalRecibo')){ 
                                    serie=objetoCHUBB[j-1].Recibo.toString().padStart(3, "0")+"/"+objetoCHUBB[j-1].TotalRecibo.toString().padStart(3, "0");
                                }else{ //Si no tiene los campos anteriores se guarda como 000/000
                                   serie="000/000";
                                }
                                /*CHUBB registra las comisiones de forma vertical, es decir, una póliza tiene varios tipos de comisiones, pero SICAS de manera horizonal.
                                En las líneas siguientes, por cada tipo de comisión se hace un objeto/renglón nuevo en caso de que este sea mayor a 0*/
                                if(objetoCHUBB[j-1].TipoMov==6){//Si el tipo de movimiento es 6, entonces la comisión es bono por internet
                                    tipocomision="Bono Internet/ Comisión por derechos";
                                    jsonObj={"Asegurado": objetoCHUBB[j-1].Asegurado, "ClaveId": claveanterior, "PolizaId": polizaanterior,"Endoso": endosoanterior, "Serie": serie,"PNetaMto": Math.round(pnetamto*100)/100, "Importe": Math.round(comisionmto*100)/100, "% Comision": porcentajecomision, "Tipo Comisión": tipocomision};
                                    objetoNuevochubb.push(jsonObj);
                                    jsonObj={};
                                }else{ //Si el movimiento es de otro tipo, es comisión base o neta, y además puede tener otras dos comisiones
                                    if(comisionrecargo>0){
                                        tipocomision="Comisión de Recargos";
                                        jsonObj={"Asegurado": objetoCHUBB[j-1].Asegurado, "ClaveId": claveanterior, "PolizaId": polizaanterior,"Endoso": endosoanterior, "Serie": serie,"PNetaMto": Math.round(pnetamto*100)/100, "Importe": Math.round(comisionrecargo*100)/100, "% Comision": porcentajecomision, "Tipo Comisión": tipocomision};
                                        objetoNuevochubb.push(jsonObj);
                                        jsonObj={};
                                    }
                                    if(comisionespecial>0){
                                        tipocomision="Comisión Especial";
                                        jsonObj={"Asegurado": objetoCHUBB[j-1].Asegurado, "ClaveId": claveanterior, "PolizaId": polizaanterior,"Endoso": endosoanterior, "Serie": serie,"PNetaMto": Math.round(pnetamto*100)/100, "Importe": Math.round(comisionespecial*100)/100, "% Comision": porcentajecomision, "Tipo Comisión": tipocomision};
                                        objetoNuevochubb.push(jsonObj);
                                        jsonObj={};
                                    }
                                    //La comisión base o neta siempre se va a registrar
                                    tipocomision="Comisión Base o de Neta";
                                    //Se insertan los datos
                                    jsonObj={"Asegurado": objetoCHUBB[j-1].Asegurado, "ClaveId": claveanterior, "PolizaId": polizaanterior,"Endoso": endosoanterior, "Serie": serie,"PNetaMto": Math.round(pnetamto*100)/100, "Importe": Math.round(comisionmto*100)/100, "% Comision": porcentajecomision, "Tipo Comisión": tipocomision}; 
                                   
                                    //console.log(jsonObj);
                                    objetoNuevochubb.push(jsonObj);
                                    jsonObj={}; //Se vacía el campo para evitar que se metan más datos

                                    //Se reestablecen los valores anteriores con los valores del objeto actual, es decir, CHUBB[j]
                                    pnetamto=objetoCHUBB[j].PNetaMto;
                                    comisionmto=objetoCHUBB[j].ComisionMto;
                                    comisionrecargo=objetoCHUBB[j].ComisionSobreRecargoMto;
                                    comisionespecial=objetoCHUBB[j].Comision2
                                }
                           
                            }
                            //Se reestablecen los valores anteriores con los valores del objeto actual, es decir, CHUBB[j]
                            polizaanterior=poliza;
                            endosoanterior=endoso;
                            claveanterior=clave;
                            
                            }
                        }
                        console.log(objetoNuevochubb);//Se imprimen los objetos de CHUBB ya tranformados
                        var reng_EC=objetoNuevochubb.length; //Se cuentan los renglones/objetos de CHUBB
                     
                    //###########FUNCIÓN QUE MANDA A LLAMAR LA BÚSQUEDA################ 
                     for(var j=1; j<objetoNuevochubb.length; j++){ //Se va a mandar cada renglón del nuevo CHUBB a comparar
                            poliza=objetoNuevochubb[j].ClaveId+" "+objetoNuevochubb[j].PolizaId; //Se juntan la clave y la póliza de CHUBB para que sea igual a la póliza en SICAS
                            resultObject = search(poliza, objetoNuevochubb[j], Sicasconendoso); //Se manda la póliza de CHUBB, el objeto de CHUBB y todos los objetos de SICAS con endoso
                            registers++; //conocer la cantidad de registros
                     }
                    for(var j=0; j<noencontrados.length; j++){//Se va a mandar cada renglón de los no encontrados en search() a comparar
                        poliza=noencontrados[j].ClaveId+" "+noencontrados[j].PolizaId; //Se juntan la clave y la póliza de CHUBB para que sea igual a la póliza en SICAS
                        search2(poliza, noencontrados[j], Sicassinendoso); //Se manda la póliza de CHUBB, el objeto de CHUBB y todos los objetos de SICAS sin endoso
                    }
                        //En el renglón de abajo se manda a la página la cantidad de renglones que se encontraron en cada documento. Si se subieron varios estados de cuenta de CHUBB, los suma todos
                        document.getElementById("numregistros").innerHTML = "Renglones Estado de Cuenta: "+reng_EC+"\nRenglones SICAS: "+reng_SICAS+"\n";
                         //Manda la tabla a la página, pero no aparecerá por el atributo HIDDEN. Primero la tabla de diferencias porque tiene el encabezado, luego los iguales, luego no encontrados y al final las fechas.
                        document.getElementById("jsondata").innerHTML = tabladiferencias+tablaiguales+tablanoencontrados+tablatipo6+"<tr><td>DEL</td><td>"+fechaantigua.getDate()+" "+month[+fechaantigua.getMonth()+1]+" "+fechaantigua.getFullYear()+"</td><td>AL</td><td>"+fechareciente.getDate()+" "+month[+fechareciente.getMonth()+1]+" "+fechareciente.getFullYear()+"</td><td># Registros</td><td>"+(registers)+"</td><td></td><td></td></tr></table>"; // DEL "+fechaantigua.getDate()+" "+month[+fechaantigua.getMonth()+1]+" "+fechaantigua.getFullYear()+" AL "+fechareciente.getDate()+" "+month[+fechareciente.getMonth()+1]+" "+fechareciente.getFullYear();;//+month[messicas]+" Año: "+aniosicas; //Se manda la tabla pero no se va a ver porque tiene HIDDEN
                    }
                        ExportToExcel('xlsx'); //Se llama la función para que convierta a XLSX directamente. Se manda el parámetro type.
                        objetoCHUBB=[]; //Se vacía el array de CHUBB por si se necesita cargar nuevos documentos
                    } catch (error) { //Si hay un error aquí se muestra.
                        document.getElementById("jsondata").innerHTML = "Algo salió mal al leer el documento. Revise que el encabezado tenga el formato correcto. Error: "+error;
                      }
                    }
            );
             
            }
        } else{//En caso de no adjuntarse nada en SICAS aquí manda el error
             document.getElementById("jsondata").innerHTML = "No se adjuntó nada en SICAS";
        }    
    }else{//En caso de no adjuntarse nada en CHUBB se manda el error. Primero revisará que haya en CHUBB y luego en SICAS
        document.getElementById("jsondata").innerHTML = "No se adjuntó el Estado de Cuenta de CHUBB";
    }
    
});
//fn: filename, dl: download, type:xlsx que se manda al llamar la función
function ExportToExcel(type, fn, dl) {// función que convierte a excel
    var elt = document.getElementById('CHUBB');//Nombre de la tabla: 'CHUBB'
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
     //Nombre del documento
    var nombre ='CONCILIACIÓN CHUBB DEL '+fechaantigua.getDate()+" "+month[+fechaantigua.getMonth()+1]+" "+fechaantigua.getFullYear()+" AL "+fechareciente.getDate()+" "+month[+fechareciente.getMonth()+1]+" "+fechareciente.getFullYear()+".";
    return dl ? //Va a tratar de forzar un client-side download.
      XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }):
      XLSX.writeFile(wb, fn || (nombre + (type || 'xlsx')));
}
