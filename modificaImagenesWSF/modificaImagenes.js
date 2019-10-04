
var s_fuente = 'C:\\Users\\nombre-usuario\\Desktop\\origen.jpg';//debe existir origen.txt en el escritorio, poner nombre de usuario
var s_destino = 'C:\\Users\\nombre-usuario\\Desktop\\temporal.jpg';//poner nombre de usuario
WScript.Echo(s_fuente);//msgbox sencillo

var variant1 = leeArchivoBinario(s_fuente);
var s_imagenb64 = deBinarioAbase64(variant1);
var contexto1 = abreCanvasIEconImagen(s_imagenb64);
//modifica el canvas
contexto1.fillStyle = 'rgba(0, 255, 0, 1)';
contexto1.fillRect(0, 0, 20, 20);
//recoge la modificacion en formato cadena imagen b64
var imagenb64Mod = dameImagenB64CanvasMod();
WScript.Echo(imagenb64Mod);
cierraIE();

var variant2 = deBase64Abinario(imagenb64Mod);
escribeArchivoBinario(s_destino,variant2);
WScript.Echo('guardado archivo');