
function CajaTextoEntrada(texto, valorDefecto)
{
    var iexplorAux = WScript.CreateObject('InternetExplorer.Application');
    iexplorAux.navigate('about:blank');//pagina vacia
    iexplorAux.Visible = 0;//mantiene pagina vacia

    while (iexplorAux.Busy) {}//espera a que este abierto el InternetExplorer
    // InternetExplorer esta preparado
 
    var interpreteIE = iexplorAux.Document.Script;//recupera el interprete de scripts del InternetExplorer
    var entradaUsuario = interpreteIE.prompt(texto, valorDefecto); //crea una ventana de entrada de texto del InternetExplorer
    iexplorAux.Quit();//cierra el InternetExplorer
    return entradaUsuario;
}

var mensaje = 'elige uno';
var tiempoLimite = 0; //timeout
var botones = 1; // OK and cancel
var icono = 48;// Exclamation 
var consola = new ActiveXObject('WScript.Shell'); 
var eleccion = consola.Popup(mensaje, tiempoLimite, 'titulo ventana', botones + icono);
mensaje='falso';
if(eleccion=='1')mensaje='cierto';

var nombre = CajaTextoEntrada('mete tu nombre '+mensaje, 'yo');
if (nombre != null){
    WScript.Echo(nombre);//msgbox sencillo
}
else
    WScript.Echo('No input');