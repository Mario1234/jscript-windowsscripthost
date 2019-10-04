var fuente = 'C:\\Users\\nombre_usuario\\Desktop\\origen.txt';//debe existir origen.txt en el escritorio, poner nombre de usuario
var destino = 'C:\\Users\\nombre_usuario\\Desktop\\temporal.txt';//poner nombre de usuario
WScript.Echo(fuente);//msgbox sencillo
var sistArchAux = WScript.CreateObject('Scripting.FileSystemObject');
var inArchivo = sistArchAux.GetFile(fuente);
var outArchivo = sistArchAux.CreateTextFile(destino,true);
var lector = inArchivo.OpenAsTextStream(1,-2);
var linea;
while(!lector.AtEndOfStream){
	linea = lector.ReadLine();
	outArchivo.Write(linea);
}
outArchivo.Close();