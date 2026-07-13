var iType,sCheck

function trimTrue(val){
	var ltrim = /^\s+/g;
	var rtrim = /\s+$/g;
	return val.replace(ltrim,'').replace(rtrim,'');
}

function LetIType(obj){
	var frm = window.frames;

	iType = obj.value;

	document.forms[0].txtItmDesc.value = "";
	document.forms[0].txtitmCode.value = "";

	/*if (DisableEnable(iType)) {
	    	document.formname.txtitmCode.readOnly = true;
		document.formname.btnYrnCode.disabled = false;
	}
	else {
		document.formname.txtitmCode.readOnly = false;
		document.formname.btnYrnCode.disabled = true;
	}*/

}

function CheckSubmit(){
	var frm = window.frames;
	var classSelected = frm(0).ctlCategoryTree.classification;
	if (document.forms[0].selItmType.value == "select") {
		alert("Select Item Type");
		document.forms[0].selItmType.focus();
		return false;
	}
	else if(classSelected == "") {
		alert("Select Classification");
		return false;
	}
	else if (iType == "YRN" && trimTrue(document.forms[0].itmCode.value) == "") {
		alert("Create Code")
		CreateItemCode(document.formname.btnYrnCode.value);
	}

	else if (iType == "PLA" && document.forms[0].hWMCCode.value == "") {
		alert("Select Machine Center")
		CreateItemCode(document.formname.btnYrnCode.value);
	}

	else if (iType != "YRN" && trimTrue(document.forms[0].itmCode.value) == "") {
		alert("Enter Item Code");
		document.forms[0].itmCode.select();
		return false;
	}
	else if (iType == "YRN" && itemCodeCheck(trimTrue(document.forms[0].itmCode.value),trimTrue(document.forms[0].txtItmDesc.value),sCheck)) {
		alert("Code or Description for Item or for Temporary Item Already Exists");
		CreateItemCode(document.formname.btnYrnCode.value);
	}
	else if (iType != "YRN" && itemCodeCheck(trimTrue(document.forms[0].itmCode.value),trimTrue(document.forms[0].txtItmDesc.value),sCheck)) {
		alert("Code or Description for Item or for Temporary Item Already Exists");
		document.forms[0].itmCode.select();
		return false;
	}

	else if (iType == "PLA" && trimTrue(document.forms[0].itmCode.value) == "") {
		alert("Enter Item Code");
		document.forms[0].itmCode.select();
		return false;
	}
	else if (iType == "PLA" && itemCodeCheck(trimTrue(document.forms[0].itmCode.value),trimTrue(document.forms[0].txtItmDesc.value),sCheck)) {
		alert("Code or Description for Item or for Temporary Item Already Exists");
		CreateItemCode(document.formname.btnYrnCode.value);
	}
	else if (trimTrue(document.forms[0].txtItmDesc.value) == "") {
		alert("Enter Item Description");
		document.forms[0].txtItmDesc.select();
		return false;
	}
	else if (iType == "YRN" && trimTrue(document.forms[0].txtItmCat.value) == "") {
		alert("Enter Dia");
		document.forms[0].txtItmCat.select();
		return false;
	}
	else if (iType == "FAB" && trimTrue(document.forms[0].txtItmCat.value) == "") {
		alert("Enter Dia");
		document.forms[0].txtItmCat.select();
		return false;
	}
	else if (iType == "YRN" && trimTrue(document.forms[0].txtItmDrw.value) == "") {
		alert("Enter GSM");
		document.forms[0].txtItmDrw.select();
		return false;
	}
	else if (iType == "FAB" && trimTrue(document.forms[0].txtItmDrw.value) == "") {
		alert("Enter GSM");
		document.forms[0].txtItmDrw.select();
		return false;
	}
	/*else if (trimTrue(document.forms[0].txtItmShDesc.value) == "") {
		alert("Enter Item Short Description");
		document.forms[0].txtItmShDesc.select();
		return false;
	}
	else if (trimTrue(document.forms[0].txtItmAddDesc.value) == "") {
		alert("Enter Item Additional Description");
		document.forms[0].txtItmAddDesc.select();
		return false;
	}
	else if (iType != "PLA" && document.forms[0].selItmUsage.value == "") {
		alert("Select Usage of Item");
		document.forms[0].selItmUsage.focus();
		return false;
	}*/
	else if (document.forms[0].selItmController[document.forms[0].selItmController.selectedIndex].value == "select") {
		alert("Select Item Controller");
		document.forms[0].selItmController.focus();
		return false;
	}
	else if (document.forms[0].selUoMStores[document.forms[0].selUoMStores.selectedIndex].value == "select") {
		alert("Select Stores UOM");
		document.forms[0].selUoMStores.focus();
		return false;
	}
	else if (iType != "PLA" && document.forms[0].selUoMPurchase.selectedIndex == "0" && document.forms[0].selUoMManu.selectedIndex == "0" && document.forms[0].selUoMSales.selectedIndex == "0") {
		alert("Select any other UOM");
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMPurchase.selectedIndex != "0") && (document.forms[0].selUoMPurchase.selectedIndex != document.forms[0].selUoMStores.selectedIndex) && (trimTrue(document.forms[0].txtStToPur.value) == "")) {
		alert("Enter Stores To Purchase Conversion");
		document.forms[0].txtStToPur.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMPurchase.selectedIndex != "0") && (document.forms[0].selUoMPurchase.selectedIndex != document.forms[0].selUoMStores.selectedIndex) && (trimTrue(document.forms[0].txtStToPur.value) != "") && (isNaN(document.forms[0].txtStToPur.value))) {
		alert("Only Numerals are Allowed");
		document.forms[0].txtStToPur.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMPurchase.selectedIndex != "0") && (document.forms[0].selUoMPurchase.selectedIndex != document.forms[0].selUoMStores.selectedIndex) && (trimTrue(document.forms[0].txtStToPur.value) != "") && (document.forms[0].selStToPur.selectedIndex == "0")) {
		alert("Select Stores To Purchase Operator");
		document.forms[0].selStToPur.focus();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMManu.selectedIndex != "0") && (document.forms[0].selUoMManu.selectedIndex != document.forms[0].selUoMStores.selectedIndex) && (trimTrue(document.forms[0].txtStToManu.value) == "")) {
		alert("Enter Stores To Manufacturing Conversion");
		document.forms[0].txtStToManu.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMManu.selectedIndex != "0") && (document.forms[0].selUoMManu.selectedIndex != document.forms[0].selUoMStores.selectedIndex) && (trimTrue(document.forms[0].txtStToManu.value) != "") && (isNaN(document.forms[0].txtStToManu.value))) {
		alert("Only Numerals are Allowed");
		document.forms[0].txtStToManu.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMManu.selectedIndex != "0") && (document.forms[0].selUoMManu.selectedIndex != document.forms[0].selUoMStores.selectedIndex) && (trimTrue(document.forms[0].txtStToManu.value) != "") && (document.forms[0].selStToManu.selectedIndex == "0")) {
		alert("Select Stores To Manufacturing Operator");
		document.forms[0].selStToManu.focus();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMSales.selectedIndex != "0") && (document.forms[0].selUoMSales.selectedIndex != document.forms[0].selUoMStores.selectedIndex) && (trimTrue(document.forms[0].txtStToSales.value) == "")) {
		alert("Enter Stores To Sales Conversion");
		document.forms[0].txtStToSales.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMSales.selectedIndex != "0") && (document.forms[0].selUoMSales.selectedIndex != document.forms[0].selUoMStores.selectedIndex) && (trimTrue(document.forms[0].txtStToSales.value) != "") && (isNaN(document.forms[0].txtStToSales.value))) {
		alert("Only Numerals are Allowed");
		document.forms[0].txtStToSales.select();
		return false;
	}
	else if ((iType != "PLA") && (document.forms[0].selUoMSales.selectedIndex != "0") && (document.forms[0].selUoMSales.selectedIndex != document.forms[0].selUoMStores.selectedIndex) && (trimTrue(document.forms[0].txtStToSales.value) != "") && (document.forms[0].selStToSales.selectedIndex == "0")) {
		alert("Select Stores To Sales Operator");
		document.forms[0].selStToSales.focus();
		return false;
	}
	else {
		if (CheckName()){
			document.forms[0].B1.disabled = 1;
			document.forms[0].action = "itmCreationInsert.asp"
			document.forms[0].hClassSelected.value = classSelected;
			document.forms[0].hIType.value = document.forms[0].selItmType.value;
			document.forms[0].submit();
		}
	}
}

(function (window, document) {
	"use strict";

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function trimValue(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name);
	}

	function responseXmlText(xhr) {
		if (xhr.responseXML && xhr.responseXML.documentElement && window.XMLSerializer) {
			return new XMLSerializer().serializeToString(xhr.responseXML);
		}
		return xhr.responseText || "";
	}

	function loadDataXml(xhr) {
		var data = xmlIsland("Data");
		var xml = responseXmlText(xhr);
		if (data && data.loadXML && trimValue(xml) !== "") {
			data.loadXML(xml);
		}
		return data && data.documentElement;
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
		}
		return window.open(url, "_blank", "width=600,height=500,resizable=yes,scrollbars=yes");
	}

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function firstAttribute(node) {
		return node && node.attributes && node.attributes.length ? node.attributes[0] : null;
	}

	function DisplayItemCode() {
		var frm = form();
		var sTempValues;
		if (frm.selIType.value === "select") {
			alert("Select Item Type");
			frm.selIType.focus();
			return false;
		}
		sTempValues = frm.selIType.value + ":" + selectedText(frm.selIType);
		openDialog("ExistingItemCodePop.asp?sTemp=" + encodeURIComponent(sTempValues), "", "dialogHeight:330px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function () {});
		return true;
	}

	function ChangeLabel(obj) {
		if (!obj || obj.value === "FAB" || obj.value === "GAR") {
			return;
		}
		if (window.idCat) {
			window.idCat.innerHTML = "Catalogue No.";
		}
		if (window.idDrw) {
			window.idDrw.innerHTML = "Draw. Ver";
		}
	}

	function applyCreatedItemCode(returnedValue) {
		var frm = form();
		var arrTemp;
		var arrValue;
		var parts;
		var iLevel = "";
		var iGroup = "";
		var iCode = "";
		if (returnedValue == null || returnedValue === "") {
			return;
		}
		arrTemp = String(returnedValue).split("``");
		frm.txtItmDesc.value = arrTemp[0] || "";
		if (frm.selIType.value === "GAR") {
			arrValue = String(arrTemp[1] || "").split(",");
			for (var i = 0; i < arrValue.length; i += 1) {
				parts = arrValue[i].split("|");
				if (parts.length < 2) {
					continue;
				}
				iLevel += "," + parts[0];
				iGroup += "," + parts[1];
				iCode += parts[1];
			}
			frm.txtitmCode.value = iCode;
			if (frm.hGroup) {
				frm.hGroup.value = iGroup.substring(1);
			}
			if (frm.hLevel) {
				frm.hLevel.value = iLevel.substring(1);
			}
		} else {
			frm.txtitmCode.value = arrTemp[1] || "";
		}
		frm.txtitmCode.readOnly = true;
	}

	function CreateItemCode(obj) {
		var frm = form();
		var prefix = String(obj || "").charAt(0);
		var baseUrl;
		var features;
		if (prefix === "C") {
			baseUrl = frm.selIType.value === "GAR" ? "itmGarCodeCreate.asp" : "itmCodeCreate.asp";
			features = frm.selIType.value === "FAB" || frm.selIType.value === "GAR" ?
				"dialogHeight:490px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No" :
				"dialogHeight:285px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No";
			openDialog(
				baseUrl + "?sTemp=" + encodeURIComponent(frm.selIType.value) + "&ItemCode=" + encodeURIComponent(frm.txtitmCode.value) + "&ItemDesc=" + encodeURIComponent(frm.txtItmDesc.value),
				"",
				features,
				applyCreatedItemCode
			);
		} else if (prefix === "M") {
			openDialog("MachineCenterCreate.asp", frm.hWMCCode ? frm.hWMCCode.value : "", "dialogHeight:270px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function (returnedValue) {
				if (frm.hWMCCode && returnedValue != null) {
					frm.hWMCCode.value = returnedValue;
				}
			});
		}
		return true;
	}

	function itemCodeCheck(iValue, sDesc, str) {
		var root = loadDataXml(syncGet("itmCodeXMLSelect.asp"));
		var children;
		var attr;
		var value = trimValue(iValue).toLowerCase();
		void sDesc;
		if (!root || !root.hasChildNodes()) {
			return false;
		}
		children = root.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType !== 1) {
				continue;
			}
			attr = firstAttribute(children[i]);
			if (!attr) {
				continue;
			}
			if (str === "T" && String(attr.nodeName || attr.name || "").charAt(0).toUpperCase() === "T") {
				continue;
			}
			if (value === trimValue(attr.nodeValue || attr.value).toLowerCase()) {
				return true;
			}
		}
		return false;
	}

	function DisableEnable(iValue) {
		var root = loadDataXml(syncGet("XMLCodeMaster.asp?sItmType=" + encodeURIComponent(iValue)));
		var children = root ? root.childNodes || [] : [];
		var attr;
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType !== 1) {
				continue;
			}
			attr = firstAttribute(children[i]);
			return !!attr && String(attr.nodeValue || attr.value) === "Y";
		}
		return false;
	}

	function CheckName() {
		var root = loadDataXml(syncGet("XMLCheckItemName.asp?ItemName=" + encodeURIComponent(form().txtItmDesc.value)));
		var status = root && root.getAttribute ? root.getAttribute("STATUS") : "";
		if (status === "Y") {
			return confirm("Item already exists with this name.\nDo you want to create the Item.");
		}
		return true;
	}

	window.DisplayItemCode = DisplayItemCode;
	window.ChangeLabel = ChangeLabel;
	window.CreateItemCode = CreateItemCode;
	window.itemCodeCheck = itemCodeCheck;
	window.DisableEnable = DisableEnable;
	window.CheckName = CheckName;
}(window, document));
