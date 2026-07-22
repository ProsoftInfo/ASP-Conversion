(function (window, document) {
	"use strict";

	var parentDoc = null;
	var parentRoot = null;

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
		return frm && frm.elements ? frm.elements[name] || document.getElementById(name) : document.getElementById(name);
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function xmlObject(id) {
		return document.getElementById(id) || window[id] || null;
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

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
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

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function selectedText(name) {
		var select = field(name);
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function selectedValue(name) {
		var select = field(name);
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].value : select && select.value || "";
	}

	function selectByValue(name, value) {
		var select = field(name);
		var wanted = trim(value);
		if (!select || !select.options) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value) === wanted) {
				select.selectedIndex = i;
				return;
			}
		}
	}

	function resolveParent() {
		ensureCompat();
		var args = modalArgs();
		parentDoc = xmlDocument(args) || xmlDocument(xmlObject("Data"));
		parentRoot = xmlRoot(args) || xmlRoot(xmlObject("Data"));
	}

	function returnRoot() {
		if (!parentRoot) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(parentRoot);
		} else {
			window["return" + "Value"] = parentRoot;
			window.returnvalue = parentRoot;
		}
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
		}
		alert("Modern browser compatibility script is still loading. Please try again.");
		return null;
	}

	function popupSize() {
		var value = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup("1") : "";
		var parts = String(value || "").split(":");
		return {
			program: parts[0] || "ItemSelectCommon.asp",
			height: parts[1] || "500",
			width: parts[2] || "750"
		};
	}

	function applyItemSelection(root, target) {
		var names = [];
		var codes = [];
		var attrs = [];
		elementChildren(root).forEach(function (node) {
			if (/^item$/i.test(node.nodeName)) {
				names.push(getAttr(node, "ItemName"));
				codes.push(getAttr(node, "ItemCode") + ":" + getAttr(node, "ClassCode"));
				attrs.push(getAttr(node, "AttributeList"));
			}
		});
		if (target === "add") {
			if (field("SpnAdditionalMaterials")) {
				field("SpnAdditionalMaterials").innerHTML = names.join(",");
			}
			if (field("hAddMatAs")) {
				field("hAddMatAs").value = codes.join(",");
			}
			if (field("hAddAttribute")) {
				field("hAddAttribute").value = attrs.join(",");
			}
			return;
		}
		if (field("SpnMaterialToBeReceived")) {
			field("SpnMaterialToBeReceived").innerHTML = names.join(",");
		}
		if (field("cmbMatRecdAs")) {
			field("cmbMatRecdAs").value = codes.join(",");
		}
		if (field("hAttribute")) {
			field("hAttribute").value = attrs.join(",");
		}
	}

	function selectItems(target, selectMode) {
		var size = popupSize();
		var island = target === "add" ? xmlObject("ItemAddData") : xmlObject("Data");
		var unit = field("hUnit") ? field("hUnit").value : "";
		var baseUrl = "../../Common/" + size.program + "?orgID=" + encodeURIComponent(unit) + "&sIType=&Stock=Y&hSelectMode=" + encodeURIComponent(selectMode) + "&Flag=1&hDispButt=Y&PartyType=CR&CallFrom=PUR";
		function handle(result) {
			var root = xmlRoot(result) || xmlRoot(island);
			var action = getAttr(root, "Action").toUpperCase();
			var query = getAttr(root, "PassQuery");
			if (action && action !== "DONE" && action !== "CLOSE" && query) {
				openDialog("../../Common/" + size.program + "?" + query, island, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", handle);
				return;
			}
			if (action === "CLOSE" || !root || !elementChildren(root).length) {
				return;
			}
			applyItemSelection(root, target);
		}
		openDialog(baseUrl, island, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", handle);
	}

	function isNumeric(value) {
		return /^\d*(\.\d+)?$/.test(trim(value));
	}

	function removeMatchingChildren(root, name, predicate) {
		elementChildren(root, name).forEach(function (node) {
			if (!predicate || predicate(node)) {
				root.removeChild(node);
			}
		});
	}

	function materialType() {
		var checked = document.querySelector('input[name="radType"]:checked');
		return checked ? checked.value : "P";
	}

	function validateProcessFields(requireMaterial) {
		if (!isNumeric(field("txtLabourCharge") ? field("txtLabourCharge").value : "")) {
			alert("Enter Only Numeric");
			if (field("txtLabourCharge")) {
				field("txtLabourCharge").select();
			}
			return false;
		}
		if (field("cmbSCProcess") && trim(selectedValue("cmbSCProcess")) === "0") {
			alert("Select Subcontracting Process");
			field("cmbSCProcess").focus();
			return false;
		}
		if (field("txtInstruct") && !trim(field("txtInstruct").value)) {
			alert("Enter Instruction");
			field("txtInstruct").focus();
			return false;
		}
		if (requireMaterial && field("cmbMatRecdAs") && trim(field("cmbMatRecdAs").value) === "0") {
			alert("Select Material Received As");
			return false;
		}
		return true;
	}

	function appendReceivedDetails(parentNode, materialCodes, materialNames, attributes) {
		var codes = trim(materialCodes) ? materialCodes.split(",") : [];
		var names = trim(materialNames) ? materialNames.split(",") : [];
		var attrList = trim(attributes) ? attributes.split(",") : [];
		var detail;
		var parts;
		var attr;
		for (var i = 0; i < codes.length; i += 1) {
			parts = codes[i].split(":");
			attr = attrList[i] || "";
			if (attr.indexOf(":") !== -1) {
				attr = attr.split(":")[0].split("#").pop();
			}
			if (attr === "0") {
				attr = "";
			}
			detail = parentDoc.createElement("Details");
			setAttr(detail, "MatRecdAsItem", parts[0] || "");
			setAttr(detail, "MatRecdAsCode", parts[1] || "");
			setAttr(detail, "MatRecdAsDescr", names[i] || "");
			setAttr(detail, "AttributeList", attr);
			parentNode.appendChild(detail);
		}
	}

	function saveSelectionPopup() {
		var node;
		var hardWaste;
		var invWaste;
		if (!validateProcessFields(true)) {
			return false;
		}
		resolveParent();
		removeMatchingChildren(parentRoot, "SubContract");
		hardWaste = trim(field("txtHardWaste") && field("txtHardWaste").value) || "0";
		invWaste = trim(field("txtInvWaste") && field("txtInvWaste").value) || "0";
		node = parentDoc.createElement("SubContract");
		setAttr(node, "SCProcess", selectedValue("cmbSCProcess"));
		setAttr(node, "Instruct", field("txtInstruct") ? field("txtInstruct").value : "");
		setAttr(node, "LabourCharge", field("txtLabourCharge") ? field("txtLabourCharge").value : "");
		setAttr(node, "Currency", selectedValue("cmbCurrency"));
		setAttr(node, "HardWaste", hardWaste);
		setAttr(node, "InvWaste", invWaste);
		setAttr(node, "ProcessName", selectedText("cmbSCProcess"));
		appendReceivedDetails(node, field("cmbMatRecdAs") ? field("cmbMatRecdAs").value : "", field("SpnMaterialToBeReceived") ? field("SpnMaterialToBeReceived").innerHTML : "", "");
		parentRoot.appendChild(node);
		returnRoot();
		window.close();
		return false;
	}

	function saveDetailsPopup() {
		var matType = materialType();
		var itemCode = field("hItemCode") ? field("hItemCode").value : "";
		var classCode = field("hClassCode") ? field("hClassCode").value : "";
		var entryNo = field("hEntryNo") ? field("hEntryNo").value : "";
		var materialCodes = field("cmbMatRecdAs") ? field("cmbMatRecdAs").value : "";
		var materialNames = field("SpnMaterialToBeReceived") ? field("SpnMaterialToBeReceived").innerHTML : "";
		var attrs = field("hAttribute") ? field("hAttribute").value : "";
		var node;
		resolveParent();
		if (matType === "P" && !validateProcessFields(field("hOrderFor") && field("hOrderFor").value === "C")) {
			return false;
		}
		removeMatchingChildren(parentRoot, "PRIMARYADDITIONALDET", function (child) {
			return getAttr(child, "PItemCode") === itemCode && getAttr(child, "EntryNo") === entryNo;
		});
		if (matType !== "P") {
			node = parentDoc.createElement("PRIMARYADDITIONALDET");
			setAttr(node, "PItemCode", itemCode);
			setAttr(node, "PClassCode", classCode);
			setAttr(node, "MatType", matType);
			setAttr(node, "EntryNo", entryNo);
			parentRoot.appendChild(node);
		} else {
			(materialCodes ? materialCodes.split(",") : [""]).forEach(function (code, index) {
				var parts = code.split(":");
				node = parentDoc.createElement("PRIMARYADDITIONALDET");
				setAttr(node, "MatRecdAT", "");
				setAttr(node, "SCProcess", selectedValue("cmbSCProcess"));
				setAttr(node, "MatRecdAsItem", parts[0] || "");
				setAttr(node, "MatRecdAsCode", parts[1] || "");
				setAttr(node, "Instruct", field("txtInstruct") ? field("txtInstruct").value : "");
				setAttr(node, "LabourCharge", field("txtLabourCharge") ? field("txtLabourCharge").value : "");
				setAttr(node, "Currency", selectedValue("cmbCurrency"));
				setAttr(node, "MatRecdAsDescr", materialNames.split(",")[index] || "");
				setAttr(node, "MatRecdAsItemType", "");
				setAttr(node, "AttributeList", (attrs.split(",")[index] || "").split(":")[0].split("#").pop().replace(/^0$/, ""));
				setAttr(node, "PItemCode", itemCode);
				setAttr(node, "PClassCode", classCode);
				setAttr(node, "MatType", matType);
				setAttr(node, "EntryNo", entryNo);
				parentRoot.appendChild(node);
			});
		}
		returnRoot();
		window.close();
		return false;
	}

	window.checkNumbers = isNumeric;

	window.SelectItem = function () {
		selectItems("main", field("hItemCode") ? "S" : "M");
		return false;
	};

	window.SelectAddItem = function () {
		selectItems("add", "M");
		return false;
	};

	window.ShowAdd = function () {
		var div = field("tblAddDet");
		if (div) {
			div.style.display = materialType() === "P" ? "block" : "none";
		}
		return false;
	};

	window.saveXML = function () {
		return field("hItemCode") ? saveDetailsPopup() : saveSelectionPopup();
	};

	window.Init = function () {
		var itemCode = field("hItemCode") ? field("hItemCode").value : "";
		var classCode = field("hClassCode") ? field("hClassCode").value : "";
		var entryNo = field("hEntryNo") ? field("hEntryNo").value : "";
		resolveParent();
		if (!parentRoot) {
			return false;
		}
		if (field("hItemCode")) {
			elementChildren(parentRoot, "PRIMARYADDITIONALDET").some(function (node) {
				var matType;
				if (getAttr(node, "PItemCode") !== itemCode || getAttr(node, "PClassCode") !== classCode || getAttr(node, "EntryNo") !== entryNo) {
					return false;
				}
				if (field("SpnMaterialToBeReceived")) {
					field("SpnMaterialToBeReceived").innerHTML = getAttr(node, "MatRecdAsDescr");
				}
				if (field("cmbMatRecdAs")) {
					field("cmbMatRecdAs").value = getAttr(node, "MatRecdAsItem") + ":" + getAttr(node, "MatRecdAsCode");
				}
				matType = getAttr(node, "MatType") || "P";
				Array.prototype.forEach.call(document.querySelectorAll('input[name="radType"]'), function (radio) {
					radio.checked = radio.value === matType;
				});
				selectByValue("cmbSCProcess", getAttr(node, "SCProcess"));
				if (field("txtInstruct")) {
					field("txtInstruct").value = getAttr(node, "Instruct");
				}
				if (field("txtLabourCharge")) {
					field("txtLabourCharge").value = getAttr(node, "LabourCharge");
				}
				window.ShowAdd();
				return true;
			});
			return false;
		}
		elementChildren(parentRoot, "SubContract").some(function (node) {
			var names = [];
			var codes = [];
			selectByValue("cmbSCProcess", getAttr(node, "SCProcess"));
			if (field("txtInstruct")) {
				field("txtInstruct").value = getAttr(node, "Instruct");
			}
			if (field("txtLabourCharge")) {
				field("txtLabourCharge").value = getAttr(node, "LabourCharge");
			}
			if (field("txtHardWaste")) {
				field("txtHardWaste").value = getAttr(node, "HardWaste");
			}
			if (field("txtInvWaste")) {
				field("txtInvWaste").value = getAttr(node, "InvWaste");
			}
			selectByValue("cmbCurrency", getAttr(node, "Currency"));
			elementChildren(node, "Details").forEach(function (detail) {
				names.push(getAttr(detail, "MatRecdAsDescr"));
				codes.push(getAttr(detail, "MatRecdAsItem") + ":" + getAttr(detail, "MatRecdAsCode"));
			});
			if (field("SpnMaterialToBeReceived")) {
				field("SpnMaterialToBeReceived").innerHTML = names.join(",");
			}
			if (field("cmbMatRecdAs")) {
				field("cmbMatRecdAs").value = codes.join(",");
			}
			return true;
		});
		return false;
	};

	window.addEventListener("beforeunload", returnRoot);
}(window, document));
