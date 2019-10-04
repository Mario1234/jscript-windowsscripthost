
var iexplorAux;
var lienzo;
//abre una venatan de internet explorer con un canvas
//crea una imagen con la cadena de datos de imagen s_imagenb64 en base64
//dibuja la imagen en el canvas
//devuelve el contexto para que se pueda modificar
//s_imagenb64 : string
function abreCanvasIEconImagen(s_imagenb64){
	iexplorAux = WScript.CreateObject('InternetExplorer.Application');
	iexplorAux.navigate('about:blank');//pagina vacia
	iexplorAux.Visible = 1;//mantiene pagina visible
	iexplorAux.Left = 50;
	iexplorAux.Top = 100;
	iexplorAux.Height = 300;
	iexplorAux.Width = 480;
	iexplorAux.MenuBar = 0;
	iexplorAux.ToolBar = 0;
	iexplorAux.StatusBar = 0;

	while (iexplorAux.Busy) {};//espera a que este abierto el InternetExplorer
	// InternetExplorer esta preparado
	iexplorAux.document.open;
	iexplorAux.document.writeln('<HTML><BODY></BODY></HTML>');
	iexplorAux.document.close;
	while (iexplorAux.Busy) {};

	WScript.Echo('codificadob64');
	var imagen1 = iexplorAux.document.createElement('image');
	lienzo = iexplorAux.document.createElement('canvas');
	lienzo.width = 1224;
	lienzo.height = 768;
	lienzo.style.zIndex = 8;
	//lienzo.style.position = 'absolute';
	lienzo.style.border = '1px solid';

	var cuerpo = iexplorAux.document.body;
	cuerpo.appendChild(lienzo);
	cuerpo.appendChild(imagen1);
	imagen1.src='data:image/jpeg;charset=utf-8;base64,'+s_imagenb64;
	var contexto1 = lienzo.getContext('2d');
	WScript.Echo('va a pintar imagen');
	WScript.Echo(imagen1.src);
	contexto1.drawImage(imagen1, 0, 0);
	return contexto1;
}

//devuelve cadena de datos del canvas en base64
//se debe usar despues de modificar el canvas para 
//recoger estos datos con la info de la modificacion
function dameImagenB64CanvasMod(){
	var imagenb64ModMeta = lienzo.toDataURL("image/jpeg");
	return imagenb64ModMeta.replace('data:image/jpeg;base64,','');
}

//cierra el InternetExplorer
function cierraIE(){
	iexplorAux.Quit();
}