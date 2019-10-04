function CajaSeleccionArchivo()
{
	var iexplorAux = WScript.CreateObject('InternetExplorer.Application');
	iexplorAux.navigate('about:blank');//pagina vacia
	iexplorAux.Visible = 0;//mantiene pagina invisible

	while (iexplorAux.Busy) {};//espera a que este abierto el InternetExplorer
	// InternetExplorer esta preparado
	iexplorAux.document.open;
	iexplorAux.document.writeln('<HTML><BODY><INPUT ID="FileSelect" NAME="FileSelect" TYPE="file"><BODY></HTML>');
	iexplorAux.document.close;
	while (iexplorAux.Busy) {};
	//var elemento = iexplorAux.Document.getElementById('FileSelect');
	var elemento = iexplorAux.document.all.item('FileSelect');
	elemento.focus=true;
	elemento.click();
	var entradaUsuario = elemento.value;
	iexplorAux.Quit();//cierra el InternetExplorer
	return entradaUsuario;
}
var archivo = CajaSeleccionArchivo();
if (archivo != null){
    WScript.Echo(archivo);//msgbox sencillo
}
else
    WScript.Echo('No input');
