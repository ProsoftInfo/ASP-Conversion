(function (window, document) {
	"use strict";

	function xmlObject(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function returnRoot(root) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	window.saveXML = function () {
		var root = xmlRoot(xmlObject("OutData"));
		var now = document.querySelector('input[name="radSelection"][value="Y"]');
		if (root) {
			root.setAttribute("Done", now && now.checked ? "Y" : "N");
			returnRoot(root);
		}
		window.close();
		return false;
	};

	window.addEventListener("beforeunload", function () {
		var root = xmlRoot(xmlObject("OutData"));
		if (root) {
			returnRoot(root);
		}
	});
}(window, document));
