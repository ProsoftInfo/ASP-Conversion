// Load modern browser compatibility helpers beside this script.
(function () {
	var scripts, currentScript, source, compatSource, loader;
	if (window.ITMSModernCompat || document.querySelector('script[src*="itms-modern-compat.js"]')) {
		return;
	}
	scripts = document.getElementsByTagName("script");
	currentScript = document.currentScript || scripts[scripts.length - 1];
	source = currentScript ? currentScript.getAttribute("src") || "" : "";
	compatSource = "/Scripts/itms-modern-compat.js";
	loader = document.createElement("script");
	loader.type = "text/javascript";
	loader.src = compatSource;
	(document.head || document.documentElement).appendChild(loader);
}());

// The code below enables the rollover and rollout effect on the tab buttons
// The code to call this function is in the following comment statement
// onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)"

function tabrollover(thisvar){
	//alert(thisvar.className);
	thisvar.className="TabTableRoll";
}

function tabrollout(thisvar){
	//alert(thisvar.className);
	thisvar.className="TabTable";
}

function toolrollover(thisvar){
	//alert(thisvar.className);
	thisvar.className="ToolBarCellRoll";
}

function toolrollout(thisvar){
	//alert(thisvar.className);
	thisvar.className="ToolBarCell";
}
function toolClick(thisvar){
	//alert(thisvar.className);
	thisvar.className="ToolBarCellClick";
}
