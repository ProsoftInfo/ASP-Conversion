<SCRIPT>
(function (window, document) {
	"use strict";

	var config = window.__itmsDataListCompatConfig || {};
	var selectedValue = "-1:0";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.FormName || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function fieldValue(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setFieldValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function asArray(collection) {
		return Array.prototype.slice.call(collection || []);
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function childElements(node, name) {
		var wanted = name && String(name).toLowerCase();
		return asArray(node && node.childNodes).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function dialogArguments() {
		var args = window.dialogArguments;
		var id;
		if (!args && window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.dialogArgumentsRoot) {
			args = window.ITMSModalReturnCompat.dialogArgumentsRoot();
		}
		if (!args && window.opener && window.opener.__itmsDialogArgs) {
			id = dialogId();
			if (id && Object.prototype.hasOwnProperty.call(window.opener.__itmsDialogArgs, id)) {
				args = window.opener.__itmsDialogArgs[id];
				window.dialogArguments = args;
			}
		}
		return args;
	}

	function dialogRoot() {
		var args = dialogArguments();
		return args && args.documentElement || args && args.XMLDocument && args.XMLDocument.documentElement || null;
	}

	function createDialogNode(nodeName) {
		var args = dialogArguments();
		var root = dialogRoot();
		if (args && args.createElement) {
			return args.createElement(nodeName);
		}
		if (args && args.XMLDocument && args.XMLDocument.createElement) {
			return args.XMLDocument.createElement(nodeName);
		}
		if (root && root.ownerDocument) {
			return root.ownerDocument.createElement(nodeName);
		}
		return document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function dialogId() {
		var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		return match ? decodeURIComponent(match[1]) : "";
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

	function closeWith(value) {
		returnValue(value);
		window.close();
	}

	function queryString(sArguments) {
		var result = String(sArguments || "") + "&Query=" + trim(fieldValue("Query"));
		var searchBy = field("SearchBy");
		var hSearch = field("hSearch");
		if (searchBy) {
			if (!hSearch || Number(hSearch.value) > 0) {
				result += "&SearchBy=" + trim(searchBy.value);
			} else {
				result += "&SearchBy=";
			}
		}
		return result;
	}

	function keyValuesFromParts(parts, indexes) {
		return (indexes || []).map(function (index) {
			return trim(parts[index] || "");
		}).join("\u0001");
	}

	function keyValuesFromNode(node, attrs) {
		return (attrs || []).map(function (name) {
			return trim(attr(node, name));
		}).join("\u0001");
	}

	function selectedNodeKey(node) {
		return keyValuesFromNode(node, config.keyAttrs || []);
	}

	function rowKey(parts) {
		return keyValuesFromParts(parts, config.keyIndexes || []);
	}

	function removeKey(parts) {
		return keyValuesFromParts(parts, config.removeIndexes || []);
	}

	function removeNodeKey(node) {
		return keyValuesFromNode(node, config.removeAttrs || config.keyAttrs || []);
	}

	function removeMatching(parts, removeMode) {
		var root = dialogRoot();
		var key = removeMode ? removeKey(parts) : rowKey(parts);
		if (!root) {
			return;
		}
		childElements(root, config.nodeName).forEach(function (node) {
			if ((removeMode ? removeNodeKey(node) : selectedNodeKey(node)) === key) {
				root.removeChild(node);
			}
		});
	}

	function clearDialogNodes() {
		var root = dialogRoot();
		if (!root) {
			return;
		}
		childElements(root, config.nodeName).forEach(function (node) {
			root.removeChild(node);
		});
	}

	function appendDialogNode(parts) {
		var root = dialogRoot();
		var node;
		if (!root) {
			return;
		}
		node = createDialogNode(config.nodeName);
		(config.attrs || []).forEach(function (item) {
			setAttr(node, item.name, parts[item.index] || "");
		});
		root.appendChild(node);
	}

	function isMultiSelect() {
		return trim(fieldValue("hSelectMode")).toUpperCase() === "M";
	}

	function escapeHtml(value) {
		return String(value == null ? "" : value)
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;")
			.replace(/'/g, "&#39;");
	}

	function selectedListValue(node) {
		return (config.removeAttrs || config.keyAttrs || []).map(function (name) {
			return attr(node, name);
		}).join(":");
	}

	function selectedListHtml() {
		var root = dialogRoot();
		var html = '<br><TABLE class="TableOutLineOnly" cellspacing="1" width="100%">';
		childElements(root, config.nodeName).forEach(function (node) {
			html += '<tr><td class="ExcelDisplayCell">';
			html += '<input type="checkbox" name="chk" value="' + escapeHtml(selectedListValue(node)) + '" checked onclick="RemoveNode(this)">';
			html += '</td>';
			(config.displayAttrs || config.keyAttrs || []).forEach(function (name) {
				html += '<td class="ExcelDisplayCell">' + escapeHtml(attr(node, name)) + '</td>';
			});
			html += '</tr>';
		});
		html += "</table><br>";
		return html;
	}

	function simpleSelectedValue() {
		var checked = asArray(document.querySelectorAll('input[type="radio"]:checked,input[type="checkbox"]:checked'));
		if (!checked.length) {
			return selectedValue;
		}
		if (config.collectAll) {
			return checked.map(function (item) {
				return item.value || "";
			}).join(",").replace(/~~/g, '"');
		}
		return checked[0].value || selectedValue;
	}

	window.showpage = function (sArguments) {
		if (config.mode === "multi") {
			setFieldValue("hButtPress", "Page");
			setAttr(dialogRoot(), "Action", "Page");
			setAttr(dialogRoot(), "PassQuery", queryString(sArguments));
			closeWith(dialogRoot());
			return;
		}
		selectedValue = queryString(sArguments);
		closeWith(selectedValue);
	};

	window.sendValue = function () {
		if (config.mode === "multi") {
			setFieldValue("hButtPress", "Done");
			setAttr(dialogRoot(), "Action", "Done");
			closeWith(dialogRoot());
			return;
		}
		selectedValue = simpleSelectedValue();
		closeWith(selectedValue);
	};

	window.sendNewValue = function () {
		selectedValue = "AN";
		closeWith(selectedValue);
	};

	window.XmlFun = function (obj) {
		var parts = String(obj && obj.value || "").split(":");
		if (config.mode !== "multi") {
			return;
		}
		if (isMultiSelect()) {
			if (obj.checked) {
				appendDialogNode(parts);
			} else {
				removeMatching(parts, false);
			}
			window.DispList();
			return;
		}
		clearDialogNodes();
		if (obj.checked) {
			appendDialogNode(parts);
		}
	};

	window.Init = function () {
		if (config.mode !== "multi") {
			return;
		}
		asArray(document.querySelectorAll('input[type="checkbox"][name="pKey"],input[type="radio"][name="pKey"]')).forEach(function (item) {
			var parts = String(item.value || "").split(":");
			childElements(dialogRoot(), config.nodeName).forEach(function (node) {
				if (selectedNodeKey(node) === rowKey(parts)) {
					item.checked = true;
				}
			});
		});
		if (isMultiSelect()) {
			window.DispList();
		}
	};

	window.RemoveNode = function (obj) {
		var parts = String(obj && obj.value || "").split(":");
		if (obj && obj.checked) {
			return;
		}
		removeMatching(parts, true);
		asArray(document.querySelectorAll('input[name="pKey"]')).forEach(function (item) {
			var itemParts = String(item.value || "").split(":");
			if (rowKey(itemParts) === removeKey(parts)) {
				item.checked = false;
			}
		});
		window.DispList();
	};

	window.DispList = function () {
		var host = document.getElementById("idSelList") || window.idSelList;
		if (host) {
			host.innerHTML = selectedListHtml();
		}
	};

	window.ChkEntKey = function (eventArg) {
		var eventObj = eventArg || null;
		if (eventObj && eventObj.key === "Enter") {
			eventObj.preventDefault && eventObj.preventDefault();
			return false;
		}
		return true;
	};

	window.window_onunload = function () {
		if (config.mode === "multi") {
			if (!trim(fieldValue("hButtPress"))) {
				setAttr(dialogRoot(), "Action", "CLOSE");
				clearDialogNodes();
			}
			returnValue(dialogRoot());
			return;
		}
		returnValue(selectedValue);
	};

	window.addEventListener("beforeunload", function () {
		window.window_onunload();
	});
}(window, document));
</SCRIPT>
