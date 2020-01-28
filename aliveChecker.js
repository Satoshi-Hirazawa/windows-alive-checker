// ウインドウスタイル
var WS_NOTVISIVLE   = 0;	//ウインドウは非表示 
var WS_ACT_NOMAL    = 1;	//ウインドウはアクティブ、サイズは通常(規定値)
var WS_ACT_MIN      = 2;	//ウインドウはアクティブ、サイズは最小
var WS_ACT_MAX      = 3;	//ウインドウはアクティブ、サイズは最大
var WS_NOTACT_NOMAL = 4;	//ウインドウは非アクティブ、サイズは通常
var WS_ACT_DEF      = 5;	//ウインドウはアクティブ、サイズは前回終了時と同じ(アプリによって動作は異なる)
var WS_NOTACT_MIN   = 7;	//ウインドウは非アクティブ、サイズは最小

var INTERVAL = 5000;
// var KEY_DELAY = 5000;

var FileName = "TestApp.exe";	// exe file name.
var AppName  = "TestApp";		// application name.
var isRunning = false;

var locator;
var service;
var set;

// Shell関連の操作を提供するオブジェクトを取得
var wshell = new ActiveXObject( "WScript.Shell" );

WScript.Sleep(INTERVAL);
wshell.Run(FileName, WS_ACT_NOMAL);
// WScript.Sleep(KEY_DELAY);
// wshell.SendKeys("{ENTER}");	// 起動後ENTERキーを入力させる

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
		wshell.AppActivate(AppName);	// アプリケーションをアクティブにする
		// WScript.Sleep(KEY_DELAY);
		// wshell.SendKeys("{ENTER}");	// 起動後ENTERキーを入力させる
	}
	WScript.Sleep(INTERVAL);
}

// 実行方法
// TestApp.exeのパスを通す
// $ cscript.exe aliveChecker.js

