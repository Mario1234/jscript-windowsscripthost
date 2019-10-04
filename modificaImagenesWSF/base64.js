//Utilizando la herramienta de creacion de objetos xml de microsoft
//se codifica y descodifica un binario a base 64
//el archivo binario se pasa como parametro como una matriz de bytes dentro de un variant
//variant es el tipo generico de visual basic script, esta protegido y no se puede acceder
//a su contenido, solo se pueden usar herramientas como WScript.CreateObject('ADODB.Stream') o new ActiveXObject("Msxml2.DOMDocument.6.0")
//para convertirlo a string

//-----------codifica la imagen en base64-------------
function deBinarioAbase64(variant1){
	var oXML = new ActiveXObject("Msxml2.DOMDocument.6.0");//crea un objeto de documento XML
	var oNode = oXML.createElement("aux");
	oNode.dataType = "bin.base64";//se le dice que codifique en base64 un binario
	oNode.nodeTypedValue = variant1;
	var s_imagenb64 = oNode.text;
	oXML=null;
	oNode=null;
	return s_imagenb64;//s_imagenb64 : string
}

//-----------decodifica la imagen de base64 a binario-------------
//s_imagenb64Mod : string
function deBase64Abinario(s_imagenb64Mod){
	oXML = new ActiveXObject("Msxml2.DOMDocument.6.0");//crea un objeto de documento XML
	var oNode = oXML.createElement("aux");
	oNode.dataType = "bin.base64";//se le dice que codifique en base64 un binario
	oNode.text=s_imagenb64Mod;
	var variant2 = oNode.nodeTypedValue;
	oXML=null;
	oNode=null;
	return variant2;
}