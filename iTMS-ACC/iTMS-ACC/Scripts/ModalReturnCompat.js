(function (window, document) {
	"use strict";

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function dialogId() {
		var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		return match ? decodeURIComponent(match[1]) : "";
	}

	function dialogArguments() {
		var args = window.dialogArguments;
		var id;
		if (!args && window.opener && window.opener.__itmsDialogArgs) {
			id = dialogId();
			if (id && Object.prototype.hasOwnProperty.call(window.opener.__itmsDialogArgs, id)) {
				args = window.opener.__itmsDialogArgs[id];
				window.dialogArguments = args;
			}
		}
		return args;
	}

	function notifyDialogValue(id, value) {
		if (!id || !window.opener) {
			return;
		}
		try {
			if (window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
				window.opener.ITMSModernCompat._receiveDialogValue(id, value);
				return;
			}
		} catch (ignoreDirectReturn) {}
		try {
			window.opener.postMessage({ type: "itms-dialog-return", id: id, value: value }, window.location.origin || "*");
		} catch (ignoreMessageReturn) {}
	}

	function returnValue(value) {
		var id;
		window.returnValue = value;
		window.returnvalue = value;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(value);
			return;
		}
		id = dialogId();
		notifyDialogValue(id, value);
	}

	function install(provider) {
		var fired = false;
		var handler = function () {
			var value;
			if (fired) {
				return;
			}
			fired = true;
			value = typeof provider === "function" ? provider() : provider;
			returnValue(value);
		};
		window.window_onunload = handler;
		window.addEventListener("beforeunload", handler);
		return handler;
	}

	window.ITMSModalReturnCompat = {
		returnValue: returnValue,
		returnAndClose: function (value) {
			returnValue(value);
			window.close();
		},
		install: install,
		xmlRoot: xmlRoot,
		xmlIsland: function (name) {
			return xmlRoot(name);
		},
		dialogArguments: dialogArguments,
		dialogArgumentsRoot: function () {
			return xmlRoot(dialogArguments());
		},
		fieldValue: function (name, formName) {
			var frm = document.forms[formName || "formname"] || document.forms[0];
			var item = frm && frm.elements ? frm.elements[name] : null;
			return item ? item.value : "";
		}
	};
}(window, document));
