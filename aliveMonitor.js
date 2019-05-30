var WS_NOTVISIVLE   = 0;    
var WS_ACT_NOMAL    = 1;   
var WS_ACT_MIN      = 2;   
var WS_ACT_MAX      = 3;   
var WS_NOTACT_NOMAL = 4;   
var WS_ACT_DEF      = 5;  
var WS_NOTACT_MIN   = 7;

var INTERVAL = 5000;
// var KEY_DELAY = 5000;

var FileName = "testApp.exe";	// exe file name.
var AppName  = "testApp";		// application name.
var isRunning = false;

var locator;
var service;
var set;

var wshell = new ActiveXObject( "WScript.Shell" );

WScript.Sleep(INTERVAL);
wshell.Run(FileName, WS_ACT_NOMAL);
// WScript.Sleep(KEY_DELAY);
// wshell.SendKeys("{ENTER}");

while (true) {

	isRunning = false;

	locator = WScript.CreateObject("WbemScripting.SWbemLocator.1");
	service = locator.ConnectServer();
	set = service.ExecQuery("select * from Win32_Process");

 	for (var e = new Enumerator(set); !e.atEnd(); e.moveNext()) {
    	var p = e.item();
    	if (p.Caption == FileName) {
    		isRunning = true;
		}
	}

	if(!isRunning){
		wshell.Run(FileName, WS_ACT_NOMAL);
		WScript.Sleep(INTERVAL);
		wshell.AppActivate( AppName );
		// WScript.Sleep(KEY_DELAY);
		// wshell.SendKeys("{ENTER}");

	}
	WScript.Sleep(INTERVAL);
}

// 実行方法
// testApp.exeのパスを通す
// $ cscript.exe aliveMonitor.js

