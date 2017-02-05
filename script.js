//VARIABLES GLOBALES

//Variable de Objecto de ActiveXObject
var ws = new ActiveXObject("WScript.Shell");
/*
Por ejemplo google.es
DC=google,DC=es

Por ejemplo maps.google.es
DC=maps,DC=google,DC=es
*/
var dominio= "";

function initialize(){
//Dimensiones de la ventana
window.resizeTo(650,460);

//Variables de servidores
var nombreServidor, nombreCliente;
var btnNombreServidor = document.getElementById("btnNombreServidor");
var btnPingServidor = document.getElementById("btnPingServidor");
var btnExplorarServidor = document.getElementById("btnExplorarServidor");

//Variables de clientes
var btnNombreCliente = document.getElementById ("btnNombreCliente");
var btnNombreClienteTS = document.getElementById("btnNombreClienteTS");
var btnPingCliente = document.getElementById("btnPingCliente");
var btnSCCMResourceExplorer = document.getElementById("btnSCCMResourceExplorer");
var btnObtenerContrasenya = document.getElementById("btnObtenerContrasenya");
var lapsContrasenya = document.getElementById("lapsContrasenya");

//Variables de Usuarios
var btnUserAttribute = document.getElementById ("btnUserAttribute");
var obtenerAtributoUsuario = document.getElementById ("obtenerAtributoUsuario");

//Variables Robocopy
//var btnRobocopy = document.getElementById ("btnRobocopy");

//Lista de opciones de menú
var server = document.getElementById("server");
var client = document.getElementById("client");
var user = document.getElementById("user");
        
//ID de las cajas contenedoras
var serverContent = document.getElementById("serverContent");
var clientContent = document.getElementById("clientContent");
var userContent = document.getElementById("userContent");
var robocopyContent = document.getElementById("robocopyContent");
        
//Opción elegida
var option = "servidor";


//JSON
obtenerServidores (servidores);
obtenerClientes (clientes);


//Botones evento
btnNombreServidor.addEventListener ("click", function() {
	var nombreServidor = document.getElementById("nombreServidor");
	//var serverUser = document.getElementById("serverUser");
	//var serverPassword = document.getElementById("serverPassword");
	conexionTerminalServer(nombreServidor, "s");
});


btnNombreCliente.addEventListener ("click", conexionCMRC);
btnNombreClienteTS.addEventListener("click", function(){
	var nombreCliente = document.getElementById("nombreCliente");
	conexionTerminalServer(nombreCliente, "cl");
});

btnPingCliente.addEventListener ("click", function () {
	nombreCliente = document.getElementById("nombreCliente");
	hacerPing(nombreCliente.value);
});

btnPingServidor.addEventListener ("click", function () {
	nombreServidor = document.getElementById("nombreServidor");
	hacerPing(nombreServidor.value);
});

btnExplorarServidor.addEventListener ("click", function () {
	nombreServidor = document.getElementById("nombreServidor");
	exploradorRemoto ("\\\\"+nombreServidor.value);
});


btnSCCMResourceExplorer.addEventListener ("click", function () {
	var nombreCliente = document.getElementById("nombreCliente");
	SCCMResourceExplorer (nombreCliente.value);
});

btnObtenerContrasenya.addEventListener ("click", function () {
	var nombreCliente = document.getElementById("nombreCliente");
	var comando = "ldifde -d \""+dominio+ "\" -f output.txt   -r \"(&(cn="+nombreCliente.value+"))\" -l \"ms-Mcs-AdmPwd\" ";
	ejecutarComandoCMD (comando);
	setTimeout (function(){
		var resultadoComando = leerArchivo ("output.txt");
		lapsContrasenya.value = resultadoComando;
	}, 4000);
});

btnUserAttribute.addEventListener ("click", function () {
	var cuentaUsuario = document.getElementById ("cuentaUsuario");
	var comando = "ldifde -d \"" + dominio + "\" -f outputUser.txt -r \"(&(cn="+cuentaUsuario.value+"))\" -l pwdLastSet,displayName,distinguishedName,mDBOverHardQuotaLimit,mDBOverQuotaLimit,mDBStorageQuota";
	ejecutarComandoCMD (comando);
	setTimeout (function(){
		var resultadoComando = leerArchivo ("outputUser.txt");
		obtenerAtributoUsuario.value = resultadoComando;
	}, 4000);
});

//Botones menú de opciones
server.addEventListener ("click", function(){
            if (option != "servidor") {
                serverContent.style.display = "block";
                serverContent.style.visibility = "visible";
                clientContent.style.display = "none";
                clientContent.style.visibility = "hidden";
                userContent.style.display = "none";
                userContent.style.visibility = "hidden";
                option = "servidor";
            }
});
        
client.addEventListener ("click", function(){
            if (option != "cliente") {
                clientContent.style.display = "block";
                clientContent.style.visibility = "visible";
                serverContent.style.display = "none";
                serverContent.style.visibility = "hidden";
                userContent.style.display = "none";
                userContent.style.visibility = "hidden";
                option = "cliente";
            }
});
        
user.addEventListener ("click", function(){
            if (option != "usuario") {
                userContent.style.display = "block";
                userContent.style.visibility = "visible";
                serverContent.style.display = "none";
                serverContent.style.visibility = "hidden";
                clientContent.style.display = "none";
                clientContent.style.visibility = "hidden";
                option = "usuario";
            }
});

}



//FUNCIONES
function conexionTerminalServer (nombreEquipo, tipoEquipo) {
	
	if(nombreEquipo.value != "" ){
		
		if (tipoEquipo == "s") {
		var serverUser = document.getElementById("serverUser");
		var serverPassword = document.getElementById("serverPassword");
			if ( serverUser.value != "" && serverPassword.value != "" )  {
			registrarCredenciales ( nombreEquipo.value, serverUser.value, serverPassword.value);
			}
		}
		if (tipoEquipo == "cl") {
		
		}
	var rdpexe = 'C:\\WINDOWS\\system32\\mstsc.exe';
	//var ws = new ActiveXObject("WScript.Shell");
	ws.Exec(rdpexe+" /v:"+nombreEquipo.value);
	} else {
	alert("No se ha introducido un nombre de equipo");
	}

}

//Listar los Servidores
function obtenerServidores (servidores) {
	var browsersServer = document.getElementById("browsersServer");
	var out = "";
	var s="";
	for(var i=0; i<servidores.length; i++){
		s = servidores[i].servidor;
		out += "<option value='"+s+"'/><br />";
	}
	
	browsersServer.innerHTML = out;
}

//Listar los Clientes
function obtenerClientes (clientes) {
	var browsersClient = document.getElementById("browsersClient");
	var out = "";
	var s="";
	for(var i=0; i<clientes.length; i++){
		s = clientes[i].cliente.toLowerCase();
		out += "<option value='"+s+"'/><br />";
	}
	
	browsersClient.innerHTML = out;
}

function conexionCMRC () {
	nombreCliente = document.getElementById("nombreCliente");
	if(nombreCliente.value != null){
	var rdpexe = 'C:\\Program Files (x86)\\Microsoft Configuration Manager\\AdminConsole\\bin\\i386\\CmRcViewer.exe';
	//var ws= new ActiveXObject("WScript.Shell");
	ws.Exec(rdpexe+" "+nombreCliente.value);
	}
}



function hacerPing (host) {
	if(host != null){
	ws.run("%COMSPEC% /k ping "+host+" -t");
	}
}

function robocopy(origen, destino){
	if(origen.value != "" && destino.value != "" ){
	
	alert(origen.value);
	
	}
}

//Explorador Remoto
function exploradorRemoto (directorioCompartido) {
	var rdpexe = 'C:\\Windows\\explorer.exe';
	ws.Exec(rdpexe+" "+directorioCompartido);
}


//SCCM Resource Explorer
function SCCMResourceExplorer (host){
	var rdpexe = 'C:\\Program Files (x86)\\Microsoft Configuration Manager\\AdminConsole\\bin\\ResourceExplorer.exe';
	var parameters = "-s -sms:ResExplrQuery= \"SELECT * FROM SMS_R_SYSTEM WHERE Name=\'"+host+"\' \" ";
	ws.Exec(rdpexe+" "+parameters);
}

function registrarCredenciales (serverName, user, password) {
	ws.Exec("cmdkey /generic:termsrv/"+serverName+" /user:"+user+" /pass:"+password);
	setTimeout (function(){
		ws.Exec("cmdkey /delete:termsrv/"+serverName);
	}, 5000);
}

//Ejecutar comando
function ejecutarComandoCMD (cmd) {
	ws.Exec (cmd);
}

//Leer archivo
function leerArchivo (filename) {
	var ForReading = 1;
    	var fso = new ActiveXObject("Scripting.FileSystemObject");

   	// Open the file for output.
    	//var filename = archivo;

    	// Open the file for input.
    	var f = fso.OpenTextFile(filename, ForReading);

    	// Read from the file.
    	if (f.AtEndOfStream)
        	   return ("");
    	else
        	   return (f.ReadAll());

	f.close();
}


//addEventListener ("load", initialize);
