//lectura y escritura de archivos binarios
//utiliza el tipo generico variant de visual basic script
//es un tpo protegido, no se puede acceder a su contenido
//leer base64.js para mas info

//-------lee archivo binario---------
//s_fuente : string
function leeArchivoBinario(s_fuente){
	var flujoLectura = WScript.CreateObject('ADODB.Stream');
	flujoLectura.Type = 1;//binario
	flujoLectura.Open();
	flujoLectura.LoadFromFile(s_fuente);
	var variant1 = flujoLectura.Read();
	flujoLectura.close();
	return variant1;//variant1 : VBVariant, objeto con los bytes del binario leido
}


//-------escribe imagen en archivo binario---------
//variant2 : VBVariant, objeto con los bytes del binario
//s_destino : string, ruta del archivo destino
function escribeArchivoBinario(s_destino, variant2){
	var flujoEscritura = WScript.CreateObject('ADODB.Stream');
	flujoEscritura.Type = 1;//binario
	flujoEscritura.Open();
	flujoEscritura.Write(variant2);
	flujoEscritura.SaveToFile(s_destino, 2);//2=si existe sobreescribe
	flujoEscritura.close();
}