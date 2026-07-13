(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return document.getElementById(name) || (frm && frm.elements ? frm.elements[name] : null);
	}

	function valueOf(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function selectedText(name) {
		var select = field(name);
		return select && select.selectedIndex >= 0 && select.options[select.selectedIndex] ? select.options[select.selectedIndex].text : "";
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlDoc(name) {
		var object = xmlObject(name);
		return object && object.XMLDocument || object && object._doc || null;
	}

	function xmlRoot(name) {
		var object = xmlObject(name);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || null;
	}

	function serializeXml(nodeOrDoc) {
		var target = nodeOrDoc && nodeOrDoc.nodeType === 9 ? nodeOrDoc : nodeOrDoc && nodeOrDoc.ownerDocument || nodeOrDoc;
		return target ? new XMLSerializer().serializeToString(target) : "";
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function treeControl() {
		return field("ctlCategoryTree");
	}

	function Init() {
		var tree = treeControl();
		if (tree) {
			tree.IType = "NAP:NO:NO";
		}
		return false;
	}

	function classSelected() {
		var tree = treeControl();
		return tree && tree.classification || "";
	}

	function DisplayItemCode() {
		var type = valueOf("selItmType");
		var tempValues;
		if (type === "select") {
			alert("Select Item Type");
			field("selItmType").focus();
			return false;
		}
		tempValues = type + ":" + selectedText("selItmType");
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog("ExistingItemCodePop.asp?sTemp=" + encodeURIComponent(tempValues), "", "dialogHeight:330px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No");
		}
		return false;
	}

	function GetDetails() {
		var tempValues = valueOf("hOrgID") + ":" + valueOf("txtItmDesc") + ":" + valueOf("txtItmDesc");
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog("AssetBoMItmEntry.asp?sTemp=" + encodeURIComponent(tempValues), xmlObject("OutDataO"), "dialogHeight:450px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No");
		}
		return false;
	}

	function itemCodeCheck(itemCode, description) {
		var xhr = new XMLHttpRequest();
		var xml;
		var root;
		var nodes;
		xhr.open("GET", "itmCodeXMLSelect.asp", false);
		xhr.send();
		if (!trim(xhr.responseText) && !xhr.responseXML) {
			return false;
		}
		xml = xhr.responseXML && xhr.responseXML.documentElement ? xhr.responseXML : new DOMParser().parseFromString(xhr.responseText || "<root/>", "application/xml");
		root = xml.documentElement;
		nodes = root ? root.childNodes : [];
		for (var i = 0; i < nodes.length; i += 1) {
			if (nodes[i].nodeType !== 1) {
				continue;
			}
			if (nodes[i].attributes[0] && trim(itemCode).toLowerCase() === String(nodes[i].attributes[0].nodeValue).toLowerCase()) {
				return true;
			}
			if (nodes[i].attributes[2] && trim(description).toLowerCase() === String(nodes[i].attributes[2].nodeValue).toLowerCase()) {
				return true;
			}
		}
		return false;
	}

	function selectedStorageValues() {
		var select = field("selStorage");
		var values = [];
		if (!select) {
			return values;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (select.options[i].selected) {
				values.push(select.options[i].value);
			}
		}
		return values;
	}

	function appendElement(doc, parent, name, attrs) {
		var node = doc.createElement(name);
		Object.keys(attrs || {}).forEach(function (key) {
			node.setAttribute(key, attrs[key] == null ? "" : String(attrs[key]));
		});
		parent.appendChild(node);
		return node;
	}

	function CreateXML() {
		var doc = xmlDoc("OutData");
		var root = xmlRoot("OutData");
		var rootO = xmlRoot("OutDataO");
		var classValue = classSelected();
		var classParts = classValue.split(":");
		var itemTypeId = "6";
		var categoryCode = classParts[1] || "";
		var classCode = classParts[3] || "";
		var bomNode;
		var assetNodes;
		var itemNodes;
		var storageValues = selectedStorageValues();
		var unitId = trim(valueOf("hOrgID"));
		var lastLocation = "";
		var storeName = "";
		var storageNode;
		var lotNode;
		var xhr;
		if (!doc || !root) {
			return false;
		}
		clearChildren(root);
		appendElement(doc, root, "CLASSIFICATION", {
			CODE: classCode,
			NAME: "ASSET",
			CATEGORY: categoryCode
		});
		appendElement(doc, root, "UOMDETAILS", {
			PUR: valueOf("selUoMStores"),
			MAN: valueOf("selUoMStores"),
			SAL: valueOf("selUoMStores"),
			PURFAC: "1",
			PUROPE: "0",
			SALFAC: "1",
			SALOPE: "0",
			MANFAC: "0",
			MANOPE: "Select"
		});
		appendElement(doc, root, "DETAILS", {
			ITYPE: itemTypeId,
			ICODE: trim(valueOf("hItemCode")),
			COMPITEMCODE: trim(valueOf("txtItmCode")),
			DESC: trim(valueOf("txtItmDesc")),
			SHDESC: "",
			CATALOUGE: "",
			DRAWVER: "",
			VARIANT: "",
			ADDDESC: "",
			UOM: valueOf("selUoMStores"),
			CATEGORY: categoryCode,
			ATTRIBUTES: "NULL",
			GROUP: "",
			LEVEL: "",
			UNIT: unitId,
			OPSTOCKUNIT: unitId,
			DESCRCODE: "",
			MODVAT: "",
			PURTAX: "",
			SALTAX: ""
		});
		bomNode = appendElement(doc, root, "BOM", {});
		assetNodes = rootO && rootO.selectNodes ? rootO.selectNodes("//ASSET") : [];
		itemNodes = rootO && rootO.selectNodes ? rootO.selectNodes("//ASSET/ITMDET") : [];
		for (var i = 0; i < itemNodes.length; i += 1) {
			appendElement(doc, bomNode, "Item", {
				ItemCode: itemNodes[i].getAttribute("ITMCODE") || "",
				ClassCode: assetNodes.length ? trim(assetNodes[0].getAttribute("CLACODE") || "") : "",
				Qty: itemNodes[i].getAttribute("QTY") || "",
				UoM: itemNodes[i].getAttribute("SUOM") || "",
				Type: itemNodes[i].getAttribute("ITYPE") || "",
				Consumable: "Y"
			});
		}
		appendElement(doc, root, "CONTROLS", {
			RECNUM: "N",
			ROUTING: "S",
			ACCOUNTING: "W",
			MODVAT: "0",
			REORDERLEVEL: "0",
			REORDERQTY: "0",
			ECOORDERQTY: "0",
			BOMAPPLICABLE: itemNodes.length > 0 ? "1" : "0"
		});
		storageValues.forEach(function (value) {
			var parts = String(value).split("~");
			lastLocation = parts[0] || "";
			storeName = parts[3] || "";
			appendElement(doc, root, "STOREDET", {
				UNITSTORE: trim(unitId + "-" + lastLocation + "-0"),
				STORE: trim(storeName)
			});
		});
		storageNode = appendElement(doc, root, "STORAGE", {
			STORE: trim(lastLocation),
			BIN: "0",
			MONTHYEAR: valueOf("hMonthYr"),
			QTY: trim(valueOf("hQty")),
			STORAGEVALUE: trim(valueOf("hValue")),
			CLASSIFICATION: trim(classCode),
			UNIT: unitId
		});
		lotNode = appendElement(doc, storageNode, "LotSerial", {
			QTYIN: "N",
			TARE: "0",
			LOT: "",
			SERIALFROM: "",
			SERIALTO: "",
			TAREWEIGHT: "U",
			IVALUE: valueOf("hValue"),
			QTY: trim(valueOf("hQty")),
			COUNTER: "1",
			STAGE: "Select",
			ALTGROSS: "0",
			ALTNETT: "0",
			ALTUOM: "Select",
			AUTOGEN: "AUTO",
			ATTRIBUTE: "NULL"
		});
		appendElement(doc, lotNode, "LotSerialDetails", {
			LOTSERIAL: "1",
			QTYREC: trim(valueOf("hQty")),
			TAREREC: "0",
			SELLINGTYPE: "0",
			WEIGHTSTYPE: "0",
			PACKINGTYPE: "0",
			LOT: "",
			SELLINGFORM: "0",
			PACKNUMBER: "0",
			IVALUE: "0",
			ATTRIBUTELIST: "NULL",
			SERIALQTY: trim(valueOf("hQty")),
			BARCODE: ""
		});
		xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSave.asp?SessionFlag=False&Value=AssetItem&Folder=Master", false);
		xhr.send(serializeXml(doc));
		return true;
	}

	function CheckSubmit() {
		var classValue = classSelected();
		var button = field("B1");
		var resetButton = field("B2");
		var xhr;
		var passValue;
		if (classValue === "") {
			alert("Select Classification");
			return false;
		}
		if (classValue.split("|").length > 1) {
			alert("Select only ONE Classification");
			return false;
		}
		if (trim(valueOf("txtItmCode")) === "") {
			alert("Enter Item Code");
			field("txtItmCode").select();
			return false;
		}
		if (itemCodeCheck(trim(valueOf("txtItmCode")), trim(valueOf("txtItmDesc")))) {
			alert("Code or Description for Item or for Temporary Item Already Exists");
			field("txtItmCode").select();
			return false;
		}
		if (trim(valueOf("txtItmDesc")) === "") {
			alert("Enter Item Description");
			field("txtItmDesc").select();
			return false;
		}
		if (valueOf("selUoMStores") === "select") {
			alert("Select Stores UOM");
			field("selUoMStores").focus();
			return false;
		}
		if (selectedStorageValues().length === 0) {
			alert("Select Storage");
			field("selStorage").focus();
			return false;
		}
		if (!(field("chkAppP").checked || field("chkAppM").checked)) {
			alert("Select Applicable For");
			field("chkAppP").focus();
			return false;
		}
		if (button) {
			button.disabled = true;
		}
		xmlRoot("OutDataO").setAttribute("ORG", trim(valueOf("hOrgID")));
		xmlRoot("OutDataO").setAttribute("ASSETCODE", trim(valueOf("hAssetCode")));
		xmlRoot("OutDataO").setAttribute("ITMCODE", trim(valueOf("txtItmCode")));
		xmlRoot("OutDataO").setAttribute("CLACODE", trim(classValue));
		xmlRoot("OutDataO").setAttribute("DESC", trim(valueOf("txtItmDesc")));
		xmlRoot("OutDataO").setAttribute("SUOM", valueOf("selUoMStores"));
		xmlRoot("OutDataO").setAttribute("STORAGE", selectedStorageValues().join("|") + "|");
		if (field("chkAppP").checked && field("chkAppM").checked) {
			xmlRoot("OutDataO").setAttribute("APPLI", "PRD,MAT");
		} else if (field("chkAppP").checked) {
			xmlRoot("OutDataO").setAttribute("APPLI", "PRD");
		} else if (field("chkAppM").checked) {
			xmlRoot("OutDataO").setAttribute("APPLI", "MAT");
		}
		CreateXML();
		setValue("hClassSelected", classValue);
		passValue = trim(valueOf("hCallFrom")) + ":" + trim(valueOf("hAssetCode"));
		xhr = new XMLHttpRequest();
		xhr.open("POST", "ItmCreationDefinitionInsert.asp?sPassValue=" + encodeURIComponent(passValue), false);
		xhr.send(serializeXml(xmlDoc("OutData")));
		if ((xhr.responseText || "").substring(0, 13) === "ItemClassCode") {
			alert("Asset Item Created Successfully");
			if (valueOf("hCallFrom") === "INV") {
				window.location.href = "../welcome_Inventory.asp";
			} else {
				window.location.href = "../../Fixedassets/TRANSACTION/FIXEDASSETS.ASP";
			}
		} else if (trim(xhr.responseText) === "N") {
			alert("Item Name or Code Already Exists");
			if (resetButton) {
				resetButton.disabled = false;
			}
		} else {
			alert(xhr.responseText);
			if (button) {
				button.disabled = false;
			}
		}
		return false;
	}

	window.Init = Init;
	window.classSelected = classSelected;
	window.DisplayItemCode = DisplayItemCode;
	window.GetDetails = GetDetails;
	window.itemCodeCheck = itemCodeCheck;
	window.CheckSubmit = CheckSubmit;
	window.CreateXML = CreateXML;
}(window, document));
