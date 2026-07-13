(function (window, document) {
	"use strict";

	var rowNo = 0;
	var objTemp = null;
	var root = null;
	var sType = "";
	var iItem = "";
	var iClass = "";
	var sOrgID = "";
	var iTotQty = "";
	var sStoName = "";
	var sStoresUom = "";
	var iQty = "";
	var sCheck = "";
	var iNo = "";
	var sAltUom = "";
	var sAltCheck = "";
	var iAltGross = 0;
	var iAltNett = 0;
	var iTotGross = 0;
	var iTotNett = 0;

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function trimValue(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function num(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function byId(id) {
		return document.getElementById(id);
	}

	function getText(id) {
		var element = byId(id);
		return element ? element.textContent : "";
	}

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			element.textContent = String(value);
		}
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function dialogArgumentDocument() {
		var modal = window.ITMSModalReturnCompat;
		var getDialogArgs = modal && modal["dialog" + "Arguments"];
		var args = getDialogArgs ? getDialogArgs() : null;
		if (args && args.XMLDocument) {
			return args.XMLDocument;
		}
		if (args && args._doc) {
			return args._doc;
		}
		if (args && args.nodeType === 9) {
			return args;
		}
		if (args && args.documentElement) {
			return args;
		}
		if (args && args.nodeType === 1) {
			return args.ownerDocument;
		}
		return null;
	}

	function currentHeader() {
		var children = root ? root.childNodes : [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && getAttr(children[i], "NO") === String(iNo)) {
				return children[i];
			}
		}
		return null;
	}

	function removeChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function input(name) {
		return form().elements[name];
	}

	function makeInput(name, value, check, blurHandler) {
		var control = document.createElement("input");
		control.type = "text";
		control.name = name;
		control.size = 12;
		control.value = value == null ? "" : String(value);
		control.className = "Formelem";
		control.style.textAlign = "right";
		control.setAttribute("onkeypress", "DoKeyPress('" + check + "',7,1)");
		control.onblur = blurHandler;
		return control;
	}

	function appendCell(row, className, align, content) {
		var cell = row.insertCell();
		cell.className = className;
		if (align) {
			cell.align = align;
		}
		if (typeof content === "string" || typeof content === "number") {
			cell.innerHTML = String(content);
		} else if (content) {
			cell.appendChild(content);
		}
		return cell;
	}

	function appendRow(quantity, gross, nett) {
		var table = document.getElementById("tblLot");
		var row = table.insertRow(table.rows.length);
		rowNo += 1;
		appendCell(row, "ExcelDisplayCell", "center", rowNo);
		appendCell(row, "ExcelInputCell", "", makeInput("txtDetails" + rowNo, quantity, sCheck, CalculateQty));
		appendCell(row, "ExcelInputCell", "", makeInput("txtAltGross" + rowNo, gross, sAltCheck, function () {
			CheckGrossQty(this);
		}));
		appendCell(row, "ExcelInputCell", "", makeInput("txtAltNett" + rowNo, nett, sAltCheck, function () {
			CheckNettQty(this);
		}));
	}

	function Init(sTemp, sTempCheck, sTempAltCheck) {
		var arrTemp = String(sTemp || "").split(":");
		var header;
		var children;
		var node;
		sType = arrTemp[0] || "";
		iItem = arrTemp[1] || "";
		iClass = arrTemp[2] || "";
		sOrgID = arrTemp[3] || "";
		iTotQty = arrTemp[6] || "";
		sStoName = arrTemp[7] || "";
		sStoresUom = arrTemp[8] || "";
		iQty = arrTemp[10] || "";
		iNo = arrTemp[11] || "";
		sAltUom = arrTemp[12] || "";
		iAltGross = num(arrTemp[13]);
		iAltNett = num(arrTemp[14]);
		if (trimValue(sAltUom).toLowerCase() === "select") {
			sAltUom = "-";
		}
		sCheck = sTempCheck;
		sAltCheck = sTempAltCheck;
		objTemp = dialogArgumentDocument();
		root = objTemp && objTemp.documentElement;
		header = currentHeader();
		if (header) {
			if (trimValue(getAttr(header, "CHECK")) === "Y") {
				form().chkAltUom.checked = true;
				form().txtGross.disabled = false;
				form().txtNett.disabled = false;
				form().txtGross.value = getAttr(header, "ALTGROSS");
				form().txtNett.value = getAttr(header, "ALTNETT");
			}
			children = header.childNodes || [];
			for (var i = 0; i < children.length; i += 1) {
				node = children[i];
				if (node.nodeType !== 1) {
					continue;
				}
				appendRow(getAttr(node, "QUANTITY"), getAttr(node, "ALTGROSS"), getAttr(node, "ALTNETT"));
				setText("idQtyEntered", num(getText("idQtyEntered")) + num(getAttr(node, "QUANTITY")));
			}
			EnableDone();
		}
	}

	function AddRow() {
		var frm = form();
		if (trimValue(frm.txtQty.value) === "") {
			alert("Enter Quantity");
			frm.txtQty.focus();
			return false;
		}
		if (num(frm.txtQty.value) === 0) {
			alert("Quantity should be greater than zero");
			frm.txtQty.select();
			return false;
		}
		if (CheckQty()) {
			alert("Entered Quantity should be less than or equal to Quantity " + num(getText("idQty")));
			frm.txtQty.select();
			return false;
		}
		appendRow(frm.txtQty.value, "", "");
		setText("idQtyEntered", num(getText("idQtyEntered")) + num(frm.txtQty.value));
		frm.txtQty.value = "";
		frm.txtQty.focus();
		EnableDone();
		return true;
	}

	function CheckQty() {
		return num(getText("idQty")) < num(getText("idQtyEntered")) + num(form().txtQty.value);
	}

	function finalizeForClose() {
		var header = currentHeader();
		if (!header) {
			return root;
		}
		if (num(getText("idQty")) !== num(getText("idQtyEntered"))) {
			removeChildren(header);
		} else {
			header.setAttribute("QUANTITY", String(num(getText("idQtyEntered"))));
		}
		return root;
	}

	function returnAndClose() {
		finalizeForClose();
		if (window.ITMSModalReturnCompat) {
			window.ITMSModalReturnCompat.returnAndClose(root);
		} else {
			window.close();
		}
	}

	function CheckSubmit() {
		var frm = form();
		var header;
		var item;
		var iTempGross = 0;
		var iTempNett = 0;
		var details;
		var gross;
		var nett;
		if (frm.chkAltUom.checked) {
			if (trimValue(frm.txtGross.value) === "") {
				alert("Enter Gross in " + sAltUom);
				frm.txtGross.focus();
				return false;
			}
			if (trimValue(frm.txtNett.value) === "") {
				alert("Enter Nett in " + sAltUom);
				frm.txtNett.focus();
				return false;
			}
		} else if (num(iAltGross) > 0) {
			if (CheckGrossEntry()) {
				alert("Enter Gross in " + sAltUom);
				return false;
			}
			if (CheckNettEntry()) {
				alert("Enter Nett in " + sAltUom);
				return false;
			}
		}
		if (num(getText("idQty")) !== num(getText("idQtyEntered"))) {
			return false;
		}
		header = currentHeader();
		if (!header || !objTemp) {
			return false;
		}
		removeChildren(header);
		for (var i = 1; i <= rowNo; i += 1) {
			details = input("txtDetails" + i);
			gross = input("txtAltGross" + i);
			nett = input("txtAltNett" + i);
			item = objTemp.createElement("DETAILS");
			item.setAttribute("PIECENO", String(i));
			item.setAttribute("QUANTITY", details ? details.value : "");
			item.setAttribute("ALTGROSS", gross ? gross.value : "");
			item.setAttribute("ALTNETT", nett ? nett.value : "");
			if (gross && trimValue(gross.value) !== "") {
				iTempGross += num(gross.value);
			}
			if (nett && trimValue(nett.value) !== "") {
				iTempNett += num(nett.value);
			}
			header.appendChild(item);
		}
		if (frm.chkAltUom.checked) {
			header.setAttribute("CHECK", "Y");
			header.setAttribute("ALTGROSS", frm.txtGross.value);
			header.setAttribute("ALTNETT", frm.txtNett.value);
			header.setAttribute("TOTGROSS", frm.txtGross.value);
			header.setAttribute("TOTNETT", frm.txtNett.value);
		} else {
			header.setAttribute("CHECK", "N");
			header.setAttribute("ALTGROSS", "");
			header.setAttribute("ALTNETT", "");
			header.setAttribute("TOTGROSS", String(iTempGross));
			header.setAttribute("TOTNETT", String(iTempNett));
		}
		returnAndClose();
		return true;
	}

	function CalculateQty() {
		var total = 0;
		var details;
		for (var i = 1; i <= rowNo; i += 1) {
			details = input("txtDetails" + i);
			total += num(details && details.value);
		}
		setText("idQtyEntered", total);
		EnableDone();
	}

	function EnableDone() {
		form().BtnDone.disabled = num(getText("idQty")) !== num(getText("idQtyEntered"));
	}

	function CheckEnable() {
		var enabled = form().chkAltUom.checked;
		form().txtGross.disabled = !enabled;
		form().txtNett.disabled = !enabled;
	}

	function CheckGrossQty(obj) {
		if (trimValue(obj.value) === "") {
			return;
		}
		FindQty();
		if (num(iTotGross) > num(iAltGross)) {
			alert("Gross should be less than or equal to (" + (num(iAltGross) - (num(iTotGross) - num(obj.value))) + ")");
			obj.select();
		}
	}

	function CheckNettQty(obj) {
		if (trimValue(obj.value) === "") {
			return;
		}
		FindQty();
		if (num(iTotNett) > num(iAltNett)) {
			alert("Nett should be less than or equal to (" + (num(iAltNett) - (num(iTotNett) - num(obj.value))) + ")");
			obj.select();
		}
	}

	function FindQty() {
		var children = root ? root.childNodes : [];
		var header;
		iTotGross = 0;
		iTotNett = 0;
		for (var i = 0; i < children.length; i += 1) {
			header = children[i];
			if (header.nodeType !== 1 || getAttr(header, "NO") === String(iNo)) {
				continue;
			}
			if (trimValue(getAttr(header, "TOTGROSS")) !== "") {
				iTotGross += num(getAttr(header, "TOTGROSS"));
			}
			if (trimValue(getAttr(header, "TOTNETT")) !== "") {
				iTotNett += num(getAttr(header, "TOTNETT"));
			}
		}
		for (var j = 1; j <= rowNo; j += 1) {
			if (input("txtAltGross" + j) && trimValue(input("txtAltGross" + j).value) !== "") {
				iTotGross += num(input("txtAltGross" + j).value);
			}
			if (input("txtAltNett" + j) && trimValue(input("txtAltNett" + j).value) !== "") {
				iTotNett += num(input("txtAltNett" + j).value);
			}
		}
		if (!form().txtGross.disabled && trimValue(form().txtGross.value) !== "") {
			iTotGross += num(form().txtGross.value);
		}
		if (!form().txtNett.disabled && trimValue(form().txtNett.value) !== "") {
			iTotNett += num(form().txtNett.value);
		}
	}

	function CheckGrossEntry() {
		for (var i = 1; i <= rowNo; i += 1) {
			if (input("txtAltGross" + i) && trimValue(input("txtAltGross" + i).value) !== "") {
				return false;
			}
		}
		return true;
	}

	function CheckNettEntry() {
		for (var i = 1; i <= rowNo; i += 1) {
			if (input("txtAltNett" + i) && trimValue(input("txtAltNett" + i).value) !== "") {
				return false;
			}
		}
		return true;
	}

	if (window.ITMSModalReturnCompat) {
		window.ITMSModalReturnCompat.install(finalizeForClose);
	}

	window.Init = Init;
	window.AddRow = AddRow;
	window.CheckQty = CheckQty;
	window.CheckSubmit = CheckSubmit;
	window.CalculateQty = CalculateQty;
	window.EnableDone = EnableDone;
	window.CheckEnable = CheckEnable;
	window.CheckGrossQty = CheckGrossQty;
	window.CheckNettQty = CheckNettQty;
	window.FindQty = FindQty;
	window.CheckGrossEntry = CheckGrossEntry;
	window.CheckNettEntry = CheckNettEntry;
}(window, document));
