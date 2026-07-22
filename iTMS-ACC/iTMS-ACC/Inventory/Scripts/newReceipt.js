(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		var controls = frm && frm.elements || [];
		if (controls[name]) {
			return controls[name];
		}
		for (var i = 0; i < controls.length; i += 1) {
			if (String(controls[i].name || "").toLowerCase() === String(name).toLowerCase()) {
				return controls[i];
			}
		}
		return document.getElementById(name) || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function selectedValue(name) {
		var control = field(name);
		if (!control) {
			return "";
		}
		if (control.options && control.selectedIndex >= 0) {
			return control.options[control.selectedIndex].value;
		}
		return control.value || "";
	}

	function selectedText(name) {
		var control = field(name);
		return control && control.options && control.selectedIndex >= 0 ? control.options[control.selectedIndex].text : "";
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name);
	}

	function xmlDocument(value) {
		if (!value) {
			return null;
		}
		if (value.XMLDocument) {
			return value.XMLDocument;
		}
		if (value._doc) {
			return value._doc;
		}
		if (value.nodeType === 9) {
			return value;
		}
		return value.ownerDocument || null;
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function elementChildren(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (!wanted || String(children[i].nodeName).toLowerCase() === wanted)) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function getAttrAt(node, index) {
		return trim(node && node.attributes && node.attributes[index] ? node.attributes[index].nodeValue : "");
	}

	function loadXml(islandName, text) {
		var island = xmlIsland(islandName);
		if (island && typeof island.loadXML === "function") {
			return island.loadXML(text);
		}
		return false;
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		window.open(url, "_blank");
		return null;
	}

	function popupWindowSize(type) {
		var parts;
		if (typeof window.GetWindowSizeForPopup === "function") {
			parts = String(window.GetWindowSizeForPopup(type || "1")).split(":");
			if (parts.length >= 3) {
				return {
					program: parts[0],
					height: parts[1],
					width: parts[2]
				};
			}
		}
		return {
			program: "ItemSelectCommon.asp",
			height: "500",
			width: "750"
		};
	}

	function clearOptions(control) {
		if (control && control.options) {
			control.options.length = 0;
		}
	}

	function addOption(control, text, value) {
		var option;
		if (!control) {
			return;
		}
		option = document.createElement("option");
		option.text = text;
		option.value = value;
		control.add(option);
	}

	function clearSelectionLists() {
		clearOptions(field("selFrombox"));
		clearOptions(field("selTobox"));
	}

	function selectedRoot(returnValue, islandName) {
		return xmlRoot(returnValue) || xmlRoot(xmlIsland(islandName));
	}

	function normalizeAttributeList(value) {
		var pieces = trim(value).split(",");
		var result = [];
		var left;
		var attrParts;
		for (var i = 0; i < pieces.length; i += 1) {
			left = (pieces[i].split(":")[0] || "");
			attrParts = left.split("#");
			if (trim(attrParts[1])) {
				result.push(attrParts[0]);
			}
		}
		return result.join(",");
	}

	function loadClassXml(selectedClasses, callback) {
		var orgId = selectedValue("selUnit") || (field("hOrgID") && field("hOrgID").value) || "";
		var xhr = syncGet("../Master/XMLSelectItemClass.asp?sOrgID=" + encodeURIComponent(orgId) + "&sText=" + encodeURIComponent(selectedClasses));
		var classCodes = [];
		var root;
		if (xhr.responseText) {
			loadXml("OutData", xhr.responseText);
			root = xmlRoot(xmlIsland("OutData"));
			elementChildren(root, "CLASS").forEach(function (node) {
				classCodes.push(getAttr(node, "CLASSCODE"));
			});
		}
		callback(classCodes.join(":"));
	}

	window.Search = function () {
		var usage = field("selDepart") || field("seldepart");
		var orgId = field("hOrgID") && field("hOrgID").value || "";
		var itemType = selectedValue("selItmType") || field("hIType") && field("hIType").value || "";
		var size;
		var url;
		if (usage && selectedValue(usage.name) === "select") {
			alert("select Usage");
			usage.focus();
			return false;
		}
		size = popupWindowSize("1");
		url = "../../Common/" + size.program + "?orgID=" + encodeURIComponent(orgId) + "&sIType=" + encodeURIComponent(itemType) + "&hSelectMode=R&Flag=1";
		openModal(url, xmlIsland("OutData") || xmlIsland("ItemData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (returnValue) {
			var root = selectedRoot(returnValue, "OutData");
			var nodes = elementChildren(root);
			var node;
			var attrList;
			if (!nodes.length) {
				alert("No Items found");
				return;
			}
			node = nodes[0];
			attrList = normalizeAttributeList(getAttr(node, "AttributeList"));
			if (field("hItmCode")) {
				field("hItmCode").value = getAttr(node, "ItemCode") || getAttr(node, "ITEMCODE");
			}
			if (field("hAttributeList")) {
				field("hAttributeList").value = attrList;
			}
			form().submit();
		});
		return false;
	};

	window.CheckType = function (obj) {
		var unit = field("selUnit");
		var depart = field("selDepart");
		var addType = field("selAddType");
		var itemType = field("selItmType");
		clearSelectionLists();
		if (unit && unit.selectedIndex === 0) {
			alert("Select Unit");
			unit.focus();
			if (obj) {
				obj.selectedIndex = 0;
			}
			return false;
		}
		if (depart && depart.selectedIndex === 0) {
			alert("Select Usage");
			depart.focus();
			if (obj) {
				obj.selectedIndex = 0;
			}
			return false;
		}
		if (depart && depart.value === "PRD" && addType && addType.selectedIndex === 0) {
			alert("Select Type");
			addType.focus();
			return false;
		}
		if (itemType && itemType.selectedIndex === 0) {
			alert("Select Item Type");
			itemType.focus();
			if (obj) {
				obj.selectedIndex = 0;
			}
			return false;
		}
		window.AddClass(function (selectedClasses) {
			var xhr;
			var from = field("selFrombox");
			if (!selectedClasses) {
				return;
			}
			if (from) {
				from.multiple = false;
			}
			xhr = syncGet("itmRecXMLSelect.asp?Check=N&sTemp=" + encodeURIComponent(selectedClasses) + "&orgID=" + encodeURIComponent(selectedValue("selUnit")));
			if (xhr.status === 0 || (xhr.status >= 200 && xhr.status < 300)) {
				if (trim(xhr.responseText)) {
					loadXml("OutData", xhr.responseText);
					window.popItmDisplay("N");
				} else {
					alert("No Item found for Classification selected.");
					window.clearXML();
				}
			} else {
				alert("No Item found for Classification selected.");
				window.clearXML();
			}
		});
		return false;
	};

	window.popItmDisplay = function (mode) {
		var root = xmlRoot(xmlIsland("OutData"));
		var from = field("selFrombox");
		clearSelectionLists();
		elementChildren(root).forEach(function (node) {
			var text = mode === "M" ? (getAttrAt(node, 4) + " / " + getAttrAt(node, 6)) : getAttrAt(node, 6);
			var value = getAttrAt(node, 5) + ":" + getAttrAt(node, 4) + ":" + getAttrAt(node, 0) + ":" + getAttrAt(node, 2);
			addOption(from, text, value);
		});
	};

	window.DisplayMRSItem = function () {
		var root = xmlRoot(xmlIsland("OutData"));
		var from = field("selFrombox");
		clearSelectionLists();
		elementChildren(root).forEach(function (node) {
			var text;
			var value;
			if (getAttrAt(node, 4) !== "0") {
				text = (getAttrAt(node, 9) !== "0" ? getAttrAt(node, 9) : getAttrAt(node, 4)) + " / " + getAttrAt(node, 6);
				value = "M:" + getAttrAt(node, 5) + ":" + getAttrAt(node, 4) + ":" + getAttrAt(node, 0) + ":" + getAttrAt(node, 2);
			} else {
				text = (getAttrAt(node, 8) !== "0" ? getAttrAt(node, 8) : getAttrAt(node, 7)) + " / " + getAttrAt(node, 6);
				value = "D:" + getAttrAt(node, 5) + ":" + getAttrAt(node, 7) + ":" + getAttrAt(node, 0) + ":" + getAttrAt(node, 2);
			}
			addOption(from, text, value);
		});
	};

	window.clearXML = function () {
		var root = xmlRoot(xmlIsland("OutData"));
		while (root && root.firstChild) {
			root.removeChild(root.firstChild);
		}
		clearSelectionLists();
		if (form()) {
			form().reset();
		}
	};

	window.resetAll = function (mode) {
		if (mode === "U" && field("selDepart")) {
			field("selDepart").selectedIndex = 0;
		}
	};

	window.AddClass = function (callback) {
		var itemType = field("selItmType");
		var orgId;
		var itemTypeName;
		var url;
		if (itemType && itemType.selectedIndex === 0) {
			alert("Select Item Type");
			itemType.focus();
			if (callback) {
				callback("");
			}
			return false;
		}
		orgId = selectedValue("selUnit") || (field("hOrgID") && field("hOrgID").value) || "";
		itemTypeName = selectedText("selItmType");
		url = "../../include/ClassificationSelectPop.asp?sIType=" + encodeURIComponent(selectedValue("selItmType")) + "&sOrgID=" + encodeURIComponent(orgId) + "&sITypename=" + encodeURIComponent(itemTypeName);
		openModal(url, "Classification", "dialogHeight:460px;dialogWidth:625px;center:Yes;help:No;resizable:No;status:No", function (returnValue) {
			var selected = trim(String(returnValue == null ? "" : returnValue)).split("*****")[0];
			if (!selected || selected === "-1") {
				if (callback) {
					callback("");
				}
				return;
			}
			loadClassXml(selected, function (classCodes) {
				if (callback) {
					callback(classCodes);
				}
			});
		});
		return false;
	};
}(window, document));
