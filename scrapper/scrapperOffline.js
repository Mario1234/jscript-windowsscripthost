
//contenido de la pagina:
//<html><head><title>A very simple webpage</title><basefont size=4></head><body bgcolor=FFFFFF><h1>A very simple webpage. This is an "h1" level header.</h1><h2>This is a level h2 header.</h2><h6>This is a level h6 header.  Pretty small!</h6><p>This is a standard paragraph.</p><p align=center>Now I\'ve aligned it in the center of the screen.</p><p align=right>Now aligned to the right</p><p><b>Bold text</b></p><p><strong>Strongly emphasized text</strong>  Can you tell the difference vs. bold?</p><p><i>Italics</i></p><p><em>Emphasized text</em>  Just like Italics!</p><p>Here is a pretty picture: <img src=example/prettypicture.jpg alt="Pretty Picture"></p><p>Same thing, aligned differently to the paragraph: <img align=top src=example/prettypicture.jpg alt="Pretty Picture"></p><hr><h2>How about a nice ordered list!</h2><ol>  <li>This little piggy went to market  <li>This little piggy went to SB228 class  <li>This little piggy went to an expensive restaurant in Downtown Palo Alto  <li>This little piggy ate too much at Indian Buffet.  <li>This little piggy got lost</ol><h2>Unordered list</h2><ul>  <li>First element  <li>Second element  <li>Third element</ul><hr><h2>Nested Lists!</h2><ul>  <li>Things to to today:    <ol>      <li>Walk the dog      <li>Feed the cat      <li>Mow the lawn    </ol>  <li>Things to do tomorrow:    <ol>      <li>Lunch with mom      <li>Feed the hamster      <li>Clean kitchen    </ol></ul><p>And finally, how about some <a href=http://www.yahoo.com/>Links?</a></p><p>Or let\'s just link to <a href=../../index.html>another page on this server</a></p><p>Remember, you can view the HTMl code from this or any other page by using the "View Page Source" command of your browser.</p></body></html>
var solicitadorHTTP1 = WScript.CreateObject('Msxml2.ServerXMLHTTP.6.0');
//solicitadorHTTP1.open('GET', 'https://web.ics.purdue.edu/~gchopra/class/public/pages/webdesign/05_simple.html', false);
//solicitadorHTTP1.send();
var paginaHTML1 = new ActiveXObject('MSXML2.DOMDocument.6.0');// "Msxml2.DOMDocument.6.0");  
paginaHTML1.validateOnParse = true;  
paginaHTML1.async = false;
//var s_paginaHTML1 = solicitadorHTTP1.responseText;
//s_paginaHTML1 = s_paginaHTML1.replace(/\n/g,'');//quitar saltos de linea
var s_paginaHTML2 = '<html><head><title>A very simple webpage</title><basefont size=4></head><body bgcolor=FFFFFF><h1>A very simple webpage. This is an "h1" level header.</h1><h2>This is a level h2 header.</h2><h6>This is a level h6 header.  Pretty small!</h6><p>This is a standard paragraph.</p><p align=center>Now I\'ve aligned it in the center of the screen.</p><p align=right>Now aligned to the right</p><p><b>Bold text</b></p><p><strong>Strongly emphasized text</strong>  Can you tell the difference vs. bold?</p><p><i>Italics</i></p><p><em>Emphasized text</em>  Just like Italics!</p><p>Here is a pretty picture: <img src=example/prettypicture.jpg alt="Pretty Picture"></p><p>Same thing, aligned differently to the paragraph: <img align=top src=example/prettypicture.jpg alt="Pretty Picture"></p><hr><h2>How about a nice ordered list!</h2><ol>  <li>This little piggy went to market  <li>This little piggy went to SB228 class  <li>This little piggy went to an expensive restaurant in Downtown Palo Alto  <li>This little piggy ate too much at Indian Buffet.  <li>This little piggy got lost</ol><h2>Unordered list</h2><ul>  <li>First element  <li>Second element  <li>Third element</ul><hr><h2>Nested Lists!</h2><ul>  <li>Things to to today:    <ol>      <li>Walk the dog      <li>Feed the cat      <li>Mow the lawn    </ol>  <li>Things to do tomorrow:    <ol>      <li>Lunch with mom      <li>Feed the hamster      <li>Clean kitchen    </ol></ul><p>And finally, how about some <a href=http://www.yahoo.com/>Links?</a></p><p>Or let\'s just link to <a href=../../index.html>another page on this server</a></p><p>Remember, you can view the HTMl code from this or any other page by using the "View Page Source" command of your browser.</p></body></html>';
s_paginaHTML2 = s_paginaHTML2.replace(/.*><body/g,'<body');//quita los metadatos HTML y cabeceras
s_paginaHTML2 = s_paginaHTML2.replace(/<\/html>/g,'');//quita el de fin de html

//arreglar la pagina
s_paginaHTML2 = s_paginaHTML2.replace(/<hr>/g,'<hr/>');//cierra hr
s_paginaHTML2 = s_paginaHTML2.replace(/href=[^<]+>/g,'href=com>');//quita enlaces
s_paginaHTML2 = s_paginaHTML2.replace(/<img src=.*"Pretty Picture">/g,'');//<img src="null" alt="Pretty Picture"/>  quita imagenes
s_paginaHTML2 = s_paginaHTML2.replace(/size=4>/g,'size=4/>');//cierra etiqueta basefont
s_paginaHTML2 = s_paginaHTML2.replace(/=[^<\/]+>/g,function (x) { //pone comillas a los atributos de etiquetas abiertas
	var sol = '=\"';
	sol+=x.slice(1, (x.length-1));
	sol+='\">';
    return sol;
  });
s_paginaHTML2 = s_paginaHTML2.replace(/=[^<]+\/>/g,function (x) { //pone comillas a los atributos de etiquetas cerradas: basefont
	var sol = '=\"';
	sol+=x.slice(1, (x.length-2));
	sol+='\"/>';
    return sol;
  });
//arregla fin de li s
s_paginaHTML2 = s_paginaHTML2.replace(/<li>[^<]+<\//g,function (x) {//arreglas li s seguidos de cierre de ol s y ul s
	var sol = x.slice(0, (x.length-1));
	return sol+"/li></";
  });
s_paginaHTML2 = s_paginaHTML2.replace(/<li>[^<]+<li>/g,function (x) {//solo arregla los li s impares que van seguidos de apertura li
	var sol = x.slice(0, (x.length-3));
	return sol+"/li><li>";
  });
s_paginaHTML2 = s_paginaHTML2.replace(/<li>[^<]+<li>/g,function (x) {//repetido para los pares
	var sol = x.slice(0, (x.length-3));
	return sol+"/li><li>";
});
s_paginaHTML2 = s_paginaHTML2.replace(/<li>[^<]+<ol>/g,function (x) {//arreglas li s seguidos de apertura ol s
	var sol = x.slice(0, (x.length-3));
	return sol+"/li><ol>";
  });
  
//parsea cadena pagina a arbol DOM
paginaHTML1.loadXML(s_paginaHTML2);  
if (paginaHTML1.parseError.errorCode != 0) {
	var miError1 = paginaHTML1.parseError;
	WScript.Echo("Hubo un error: " + miError1.reason);
} else {
	var listaNodosXML1 = paginaHTML1.getElementsByTagName('li'); 	
	for (var i=0; i<listaNodosXML1.length; i++) {
		WScript.Echo(listaNodosXML1.item(i).text);
	}
	//WScript.Echo(paginaHTML1.xml);
}
WScript.Quit(0);


