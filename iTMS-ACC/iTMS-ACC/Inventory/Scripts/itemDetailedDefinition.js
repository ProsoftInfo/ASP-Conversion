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

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setText(id, value) {
		var item = document.getElementById(id) || window[id];
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function selectedValue(name) {
		var item = field(name);
		return item && item.selectedIndex >= 0 && item.options[item.selectedIndex] ? item.options[item.selectedIndex].value : "";
	}

	function selectedText(name) {
		var item = field(name);
		return item && item.selectedIndex >= 0 && item.options[item.selectedIndex] ? item.options[item.selectedIndex].text : "";
	}

	function setSelectValue(name, value) {
		var item = field(name);
		if (!item || !item.options) {
			return;
		}
		for (var i = 0; i < item.options.length; i += 1) {
			if (trim(item.options[i].value) === trim(value)) {
				item.selectedIndex = i;
				return;
			}
		}
	}

	function setRadioValue(name, value) {
		var group = field(name);
		if (!group) {
			return;
		}
		if (group.length === undefined) {
			group.checked = trim(group.value) === trim(value);
			return;
		}
		for (var i = 0; i < group.length; i += 1) {
			group[i].checked = trim(group[i].value) === trim(value);
		}
	}

	function radioValue(name) {
		var group = field(name);
		if (!group) {
			return "";
		}
		if (group.length === undefined) {
			return group.checked ? group.value : "";
		}
		for (var i = 0; i < group.length; i += 1) {
			if (group[i].checked) {
				return group[i].value;
			}
		}
		return "";
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

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function xmlDoc(name) {
		var object = xmlObject(name);
		var root = xmlRoot(object);
		return object && object.XMLDocument || object && object._doc || root && root.ownerDocument || null;
	}

	function serializeXml(nameOrDoc) {
		var doc = typeof nameOrDoc === "string" ? xmlDoc(nameOrDoc) : nameOrDoc;
		return doc ? new XMLSerializer().serializeToString(doc) : "";
	}

	function elementChildren(node, name) {
		var result = [];
		var wanted = String(name || "").toUpperCase();
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toUpperCase() === wanted)) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function ensureChild(parent, name) {
		var child = elementChildren(parent, name)[0];
		if (!child) {
			child = parent.ownerDocument.createElement(name);
			parent.appendChild(child);
		}
		return child;
	}

	function removeChildren(parent, names) {
		var wanted = {};
		names.forEach(function (name) {
			wanted[String(name).toUpperCase()] = true;
		});
		elementChildren(parent).forEach(function (child) {
			if (wanted[String(child.nodeName).toUpperCase()]) {
				parent.removeChild(child);
			}
		});
	}

	function appendElement(parent, name, attrs) {
		var node = parent.ownerDocument.createElement(name);
		Object.keys(attrs || {}).forEach(function (key) {
			node.setAttribute(key, attrs[key] == null ? "" : String(attrs[key]));
		});
		parent.appendChild(node);
		return node;
	}

	function postXml(url, xmlText) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.send(xmlText || null);
		return xhr;
	}

	function dateValue(name) {
		var item = field(name);
		if (!item) {
			return "";
		}
		if (typeof item.getDate === "function") {
			return item.getDate();
		}
		if (item.GetDate) {
			return item.GetDate();
		}
		return item.value || "";
	}

	function setDateValue(name, value) {
		var item = field(name);
		if (!item) {
			return;
		}
		if (typeof item.setDate === "function") {
			item.setDate(value);
		} else if ("setDate" in item) {
			item.setDate = value;
		} else {
			item.value = value || "";
		}
	}

	function requireField(name, message) {
		if (trim(valueOf(name)) === "") {
			alert(message);
			if (field(name) && field(name).select) {
				field(name).select();
			} else if (field(name) && field(name).focus) {
				field(name).focus();
			}
			return false;
		}
		return true;
	}

	function requireSelect(name, message) {
		if (!field(name) || field(name).selectedIndex === 0) {
			alert(message);
			if (field(name) && field(name).focus) {
				field(name).focus();
			}
			return false;
		}
		return true;
	}

	function PurRecSubmit() {
		var root = xmlRoot("OutData");
		var purchase;
		var xhr;
		if (trim(valueOf("hItmCode")) === "") {
			alert("Item is Not Selected");
			window.history.back();
			return false;
		}
		if (field("chkVendor") && field("chkVendor").checked) {
			if (!requireField("hSuppCode", "Selet Supplier") ||
					!requireField("txtSupItmDesc", "Enter Supplier Item Description") ||
					!requireField("txtSupItmNo", "Enter the Item Code") ||
					!requireField("txtSuppDrawingNo", "Enter the Drawing No") ||
					!requireField("txtTrLTime", "Enter the Transit Lead Time") ||
					!requireField("txtPuLTime", "Enter Purchase Lead Time") ||
					!requireField("txtPurWarranty", "Enter Warrenty") ||
					!requireField("txtSuLTime", "Enter the Supplier Lead Time") ||
					!requireField("txtMarketPrice", "Enter the Market Price") ||
					!requireField("txtPrLTime", "Enter the Preorder lead Time") ||
					!requireField("txtMarketDate", "Enter the Market Date") ||
					!requireField("txtPreMinQty", "Enter the Preferred Min Qty") ||
					!requireField("txtPreMaxQty", "Enter the Preferred Max Qty")) {
				return false;
			}
		}
		purchase = ensureChild(root, "Purchase");
		removeChildren(purchase, ["Basic", "Vendor"]);
		appendElement(purchase, "Basic", {
			Buyer: selectedValue("selBuyer"),
			ModVat: "",
			InvMatch: selectedValue("selInvMat"),
			SubCont: radioValue("radSub"),
			SubReceipts: radioValue("radSubRec"),
			EnforceShipTo: radioValue("radShip"),
			RecDateAction: radioValue("radReDate"),
			RecDaysEarly: valueOf("txtRecDaysE"),
			RecDaysLate: valueOf("txtRecDaysL"),
			UnRecLow: valueOf("txtUnLow"),
			UnRecHigh: valueOf("txtUnHigh"),
			OverRecLow: valueOf("txtOvLow"),
			OverRecHigh: valueOf("txtOvHigh"),
			UnOrdRecLow: valueOf("txtUnOrLow"),
			UnOrdRecHigh: valueOf("txtUnOrHigh")
		});
		if (field("chkVendor") && field("chkVendor").checked) {
			appendElement(purchase, "Vendor", {
				Warrenty: valueOf("txtPurWarranty"),
				TransitLeadTime: valueOf("txtTrLTime"),
				PurLeadTime: valueOf("txtPuLTime"),
				SuppItemNo: valueOf("txtSupItmNo"),
				SuppLeadTime: valueOf("txtSuLTime"),
				MarketPrice: valueOf("txtMarketPrice"),
				PreOrdLeadTime: valueOf("txtPrLTime"),
				MarketDate: valueOf("txtMarketDate"),
				PreMinOrdQty: valueOf("txtPreMinQty"),
				PreMaxOrdQty: valueOf("txtPreMaxQty"),
				SuppCode: valueOf("hSuppCode"),
				SuppName: valueOf("hSuppName"),
				SuppType: valueOf("hSuppType"),
				SuppSubType: valueOf("hSuppSubType"),
				SuppItemDesc: valueOf("txtSupItmDesc"),
				SuppDrawingNo: valueOf("txtSuppDrawingNo"),
				SuppUOM: selectedValue("selSuppUOM")
			});
		}
		postXml(window.ITMS_DETAIL_PUR_SAVE_URL || "XMLSave.asp?SessionFlag=true&Value=PurchaseDet&Folder=Master", serializeXml("OutData"));
		xhr = postXml(window.ITMS_DETAIL_PUR_POST_URL || "ItmDetailedPurRecInsert.asp", null);
		alert(trim(xhr.responseText) === "" ? "Purchase & Receving Details Stored Sucessfully" : xhr.responseText);
		return false;
	}

	function SalSubmit() {
		var root = xmlRoot("OutData");
		var sales;
		var xhr;
		if (trim(valueOf("hItmCode")) === "") {
			alert("Item is Not Selected");
			return false;
		}
		if (window.ITMS_DETAIL_AMEND) {
			if (!requireField("txtSalWarranty", "Enter the Warrenty") ||
					!requireField("txtMinSale", "Enter the Minimum Sales Quantity")) {
				return false;
			}
			sales = ensureChild(root, "Sales");
			removeChildren(sales, ["Basic"]);
			appendElement(sales, "Basic", {
				WarrPeriod: valueOf("txtSalWarranty"),
				MinSalQty: valueOf("txtMinSale"),
				PurRate: valueOf("txtPurRate"),
				PurRatePer: valueOf("txtPurRatePer"),
				CharPer: valueOf("txtCharPer"),
				CharValue: valueOf("txtCharValue"),
				MarPer: valueOf("txtMarPer"),
				MarValue: valueOf("txtMarValue"),
				TotPrice: valueOf("txtTotalPrice"),
				EffectiveFrom: dateValue("ctlEffFrom")
			});
			postXml(window.ITMS_DETAIL_SAL_SAVE_URL || "XMLSave.asp?SessionFlag=true&Value=SalesDetUpdate&Folder=Master", serializeXml("OutData"));
			xhr = postXml(window.ITMS_DETAIL_SAL_POST_URL || "ItmDetailedSalesUpdate.asp", null);
			alert(trim(xhr.responseText) === "" ? "Sales Details Stored Sucessfully" : xhr.responseText);
			return false;
		}
		if (!requireField("txtMarketRate", "Enter the Market Rate") ||
				!requireField("txtSalWarranty", "Enter the Warrenty") ||
				!requireField("txtMinSale", "Enter the Minimum Sales Quantity") ||
				!requireField("txtActual", "Enter the Actual Rate") ||
				!requireField("txtUnitSize", "Enter the Unit Size") ||
				!requireSelect("selUoMUnit", "Select Unit UOM") ||
				!requireField("txtMin", "Enter Minimum Rate") ||
				!requireField("txtVolume", "Enter the Volume ") ||
				!requireSelect("selUoMVolume", "Select Volume UOM") ||
				!requireField("txtPreffered", "Enter Preferred Rate")) {
			return false;
		}
		sales = ensureChild(root, "Sales");
		removeChildren(sales, ["Basic"]);
		appendElement(sales, "Basic", {
			MarketRate: valueOf("txtMarketRate"),
			WarrPeriod: valueOf("txtSalWarranty"),
			MinSalQty: valueOf("txtMinSale"),
			Actual: valueOf("txtActual"),
			UnitSize: valueOf("txtUnitSize"),
			UnitUOM: selectedValue("selUoMUnit"),
			Minimum: valueOf("txtMin"),
			Volume: valueOf("txtVolume"),
			VolumeUOM: selectedValue("selUoMVolume"),
			Preferred: valueOf("txtPreffered"),
			Commodity: selectedValue("selCommodity")
		});
		postXml(window.ITMS_DETAIL_SAL_SAVE_URL || "XMLSave.asp?SessionFlag=true&Value=SalesDet&Folder=Master", serializeXml("OutData"));
		xhr = postXml(window.ITMS_DETAIL_SAL_POST_URL || "ItmDetailedSalesInsert.asp", null);
		alert(trim(xhr.responseText) === "" ? "Sales Details Stored Sucessfully" : xhr.responseText);
		return false;
	}

	function resetVendor(obj) {
		if (obj && !obj.checked) {
			setText("txtParty", " ");
		}
		return false;
	}

	function partySelectionUrl(query) {
		return query ? "../../Common/PartySelection.asp?" + query : "../../Common/PartySelection.asp?orgID=" + encodeURIComponent(valueOf("hOrgCode")) + "&Party=CR";
	}

	function applySupplierSelection(root) {
		var children = elementChildren(root);
		children.forEach(function (child) {
			setValue("hSuppName", child.getAttribute("RetField0") || "");
			setValue("hSuppCode", child.getAttribute("RetField1") || "");
			setValue("hSuppType", child.getAttribute("RetField3") || "");
			setValue("hSuppSubType", child.getAttribute("RetField4") || "");
			setText("txtParty", valueOf("hSuppName"));
		});
	}

	function openSupplier(query) {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			return;
		}
		window.ITMSModernCompat.openModalDialog(partySelectionUrl(query), xmlObject("PartyData"), "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (outValue) {
			var root = xmlRoot(outValue) || xmlRoot("PartyData");
			var action = root && root.getAttribute ? root.getAttribute("Action") : "";
			var passQuery = root && root.getAttribute ? root.getAttribute("PassQuery") : "";
			if (action === "CLOSE") {
				return;
			}
			if (action && action !== "Done") {
				openSupplier(passQuery);
				return;
			}
			if (root) {
				applySupplierSelection(root);
			}
		});
	}

	function PopulateSupplier() {
		if (!field("chkVendor") || !field("chkVendor").checked) {
			return false;
		}
		openSupplier("");
		return false;
	}

	function openModal(url, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, xmlObject("OutData"), "dialogHeight:370px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", callback);
		}
	}

	function OpenAlter() {
		var button = field("btnCheck");
		var isSelect = button && (trim(button.value) === "Select");
		var flag = isSelect ? "S" : "V";
		openModal("ItmAlterPoPEntry.asp?Flag=" + flag + "&iItmCode=" + encodeURIComponent(valueOf("hItmCode")), function (out) {
			if (isSelect && out === "OK") {
				button.value = "View";
			}
		});
		return false;
	}

	function OpenUoM(value) {
		var isPur = trim(value).toUpperCase() === "PUR";
		var button = field(isPur ? "btnUoMPur" : "btnUoMSal");
		var callFrom = isPur ? "PUR" : "SAL";
		var isSelect = button && trim(button.value) === "Select";
		var flag = isSelect ? "S" : "V";
		openModal("ItmOpUoMSalPoPEntry.asp?Flag=" + flag + "&iItmCode=" + encodeURIComponent(valueOf("hItmCode")) + "&CallFrom=" + callFrom, function (out) {
			if (isSelect && out === "OK") {
				button.value = "View";
			}
		});
		return false;
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function addCell(row, text, className, align) {
		var cell = row.insertCell(-1);
		cell.textContent = text == null ? "" : String(text);
		cell.className = className || "ExcelDisplayCell";
		if (align) {
			cell.align = align;
		}
		return cell;
	}

	function clearDisplayRows(tableId) {
		var table = document.getElementById(tableId);
		while (table && table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function validateNumberField(name, emptyMessage, numberMessage, rangeMessage) {
		var value = trim(valueOf(name));
		if (value === "") {
			alert(emptyMessage);
			field(name).select();
			return false;
		}
		if (!checkNumbers(value)) {
			alert(numberMessage || "Enter Numerals Only");
			field(name).select();
			return false;
		}
		if (rangeMessage && (Number(value) < 1 || Number(value) > 100)) {
			alert(rangeMessage);
			field(name).select();
			return false;
		}
		return true;
	}

	function nextNodeIndex(name) {
		var root = xmlRoot("OutData");
		var nodes = root && root.selectNodes ? root.selectNodes("//" + name) : [];
		return nodes.length + 1;
	}

	function CheckEntryDis() {
		var root = xmlRoot("OutData");
		var row;
		var index;
		var qtyValue;
		var appValue;
		if (!field("chkDiscount") || !field("chkDiscount").checked) {
			return false;
		}
		if (!validateNumberField("txtQtyFrom", "Enter Quantity From") ||
				!validateNumberField("txtQtyTo", "Enter Quantity To") ||
				!validateNumberField("txtQtyDis", "Enter Quantity Discount", "Enter Numerals Only", "Enter Quantity Discount between 1 and 100")) {
			return false;
		}
		index = nextNodeIndex("DisEntry");
		row = document.getElementById("tblDataQty").insertRow(-1);
		addCell(row, index, "ExcelSerial", "center");
		addCell(row, trim(valueOf("txtQtyFrom")), "ExcelDisplayCell", "right");
		addCell(row, trim(valueOf("txtQtyTo")), "ExcelDisplayCell", "right");
		addCell(row, trim(selectedText("selUoMQty")), "ExcelDisplayCell", "center");
		addCell(row, trim(valueOf("txtQtyDis")), "ExcelDisplayCell", "right");
		addCell(row, "", "ExcelDisplayCell", "center");
		qtyValue = radioValue("radQV");
		appValue = radioValue("radApplicable");
		appendElement(root, "DisEntry", {
			SNO: index,
			QTYFROM: trim(valueOf("txtQtyFrom")),
			QTYTO: trim(valueOf("txtQtyTo")),
			QTYDIS: trim(valueOf("txtQtyDis")),
			QTYUOM: trim(valueOf("selUoMQty")),
			QTYVAL: qtyValue,
			APPIN: appValue
		});
		setValue("txtQtyFrom", "");
		setValue("txtQtyTo", "");
		setValue("txtQtyDis", "");
		if (field("selUoMQty")) {
			field("selUoMQty").selectedIndex = 0;
		}
		return false;
	}

	function CheckEntryVal() {
		var root = xmlRoot("OutData");
		var row;
		var index;
		if (!field("chkDiscount") || !field("chkDiscount").checked) {
			return false;
		}
		if (!validateNumberField("txtValFrom", "Enter Value From") ||
				!validateNumberField("txtValTo", "Enter Value To") ||
				!validateNumberField("txtValDis", "Enter Value Discount", "Enter Numerals Only", "Enter Value Discount between 1 and 100")) {
			return false;
		}
		index = nextNodeIndex("ValEntry");
		row = document.getElementById("tblDataVal").insertRow(-1);
		addCell(row, index, "ExcelSerial", "center");
		addCell(row, trim(valueOf("txtValFrom")), "ExcelDisplayCell", "right");
		addCell(row, trim(valueOf("txtValTo")), "ExcelDisplayCell", "right");
		addCell(row, trim(valueOf("txtValDis")), "ExcelDisplayCell", "right");
		addCell(row, "", "ExcelDisplayCell", "center");
		appendElement(root, "ValEntry", {
			SNO: index,
			VALFROM: trim(valueOf("txtValFrom")),
			VALTO: trim(valueOf("txtValTo")),
			VALDIS: trim(valueOf("txtValDis")),
			QTYVAL: radioValue("radQV"),
			APPIN: radioValue("radApplicable")
		});
		setValue("txtValFrom", "");
		setValue("txtValTo", "");
		setValue("txtValDis", "");
		return false;
	}

	function resetDiscount(obj) {
		if (obj && !obj.checked) {
			setValue("txtQtyFrom", "");
			setValue("txtQtyTo", "");
			setValue("txtQtyDis", "");
			if (field("selUoMQty")) {
				field("selUoMQty").selectedIndex = 0;
			}
			setValue("txtValFrom", "");
			setValue("txtValTo", "");
			setValue("txtValDis", "");
		}
		return false;
	}

	function AssaignValue() {
		setValue("hItemRate", valueOf("txtPurRate"));
		CalcValue("MP");
		CalcValue("OP");
		return false;
	}

	function formatNumber(value, places) {
		var number = Number(value);
		return isNaN(number) ? "" : number.toFixed(places == null ? 2 : places);
	}

	function CalcValue(callFrom) {
		var rate = Number(valueOf("hItemRate")) || 0;
		var percentage;
		if (rate === 0) {
			return false;
		}
		if (callFrom === "MP") {
			percentage = valueOf("txtMarPer");
			if (percentage === "") {
				return false;
			}
			if (Number(percentage) < 0) {
				alert("Enter Numerals Only");
				setValue("txtMarPer", 0);
				return false;
			}
			setValue("txtMarValue", formatNumber(rate * (Number(percentage) / 100), 2));
		} else if (callFrom === "MV") {
			setValue("txtMarPer", formatNumber((Number(valueOf("txtMarValue")) * 100) / rate, 1));
		} else if (callFrom === "OP") {
			percentage = valueOf("txtCharPer");
			if (percentage === "") {
				return false;
			}
			if (Number(percentage) < 0) {
				alert("Enter Numerals Only");
				setValue("txtCharValue", 0);
				return false;
			}
			setValue("txtCharValue", formatNumber(rate * (Number(percentage) / 100), 2));
		} else if (callFrom === "OV") {
			setValue("txtCharPer", formatNumber((Number(valueOf("txtCharValue")) * 100) / rate, 1));
		}
		setValue("txtTotalPrice", (Number(rate) + Number(valueOf("txtMarValue") || 0) + Number(valueOf("txtCharValue") || 0)).toString());
		return false;
	}

	function renderDiscountRows(root) {
		var disEntries = root && root.selectNodes ? root.selectNodes("//DisEntry") : [];
		var valEntries = root && root.selectNodes ? root.selectNodes("//ValEntry") : [];
		var row;
		clearDisplayRows("tblDataQty");
		clearDisplayRows("tblDataVal");
		if (disEntries.length || valEntries.length) {
			field("chkDiscount").checked = true;
		}
		for (var i = 0; i < disEntries.length; i += 1) {
			row = document.getElementById("tblDataQty").insertRow(-1);
			addCell(row, i + 1, "ExcelSerial", "center");
			addCell(row, trim(attr(disEntries[i], "QTYFROM")), "ExcelDisplayCell", "right");
			addCell(row, trim(attr(disEntries[i], "QTYTO")), "ExcelDisplayCell", "right");
			addCell(row, trim(attr(disEntries[i], "QTYUOM")), "ExcelDisplayCell", "center");
			addCell(row, trim(attr(disEntries[i], "QTYDIS")), "ExcelDisplayCell", "right");
			addCell(row, "", "ExcelDisplayCell", "center");
			setRadioValue("radQV", attr(disEntries[i], "QTYVAL") || "Q");
			setRadioValue("radApplicable", attr(disEntries[i], "APPIN") || "B");
		}
		for (var j = 0; j < valEntries.length; j += 1) {
			row = document.getElementById("tblDataVal").insertRow(-1);
			addCell(row, j + 1, "ExcelSerial", "center");
			addCell(row, trim(attr(valEntries[j], "VALFROM")), "ExcelDisplayCell", "right");
			addCell(row, trim(attr(valEntries[j], "VALTO")), "ExcelDisplayCell", "right");
			addCell(row, trim(attr(valEntries[j], "VALDIS")), "ExcelDisplayCell", "right");
			addCell(row, "", "ExcelDisplayCell", "center");
			setRadioValue("radQV", attr(valEntries[j], "QTYVAL") || "Q");
			setRadioValue("radApplicable", attr(valEntries[j], "APPIN") || "B");
		}
	}

	function DisplayData() {
		var root = xmlRoot("OutData");
		var purchase = root && root.selectNodes ? root.selectNodes("//Purchase/Basic") : [];
		var sales = root && root.selectNodes ? root.selectNodes("//Sales/Basic") : [];
		var basic;
		if (purchase.length) {
			basic = purchase[0];
			setSelectValue("selBuyer", attr(basic, "Buyer"));
			setSelectValue("selInvMat", attr(basic, "InvMatch"));
			setRadioValue("radSub", attr(basic, "SubCont") || "0");
			setRadioValue("radSubRec", attr(basic, "SubReceipts") || "0");
			setRadioValue("radShip", attr(basic, "EnforceShipTo") || "0");
			setRadioValue("radReDate", attr(basic, "RecDateAction") || "R");
			setValue("txtRecDaysE", attr(basic, "RecDaysEarly"));
			setValue("txtRecDaysL", attr(basic, "RecDaysLate"));
			setValue("txtUnLow", attr(basic, "UnRecLow"));
			setValue("txtUnHigh", attr(basic, "UnRecHigh"));
			setValue("txtOvLow", attr(basic, "OverRecLow"));
			setValue("txtOvHigh", attr(basic, "OverRecHigh"));
			setValue("txtUnOrLow", attr(basic, "UnOrdRecLow"));
			setValue("txtUnOrHigh", attr(basic, "UnOrdRecHigh"));
		}
		if (sales.length) {
			basic = sales[0];
			if (attr(basic, "EffectiveFrom")) {
				setDateValue("ctlEffFrom", attr(basic, "EffectiveFrom"));
			}
			setValue("txtSalWarranty", attr(basic, "WarrPeriod"));
			setValue("txtMinSale", attr(basic, "MinSalQty"));
			setValue("txtPurRate", attr(basic, "PurRate"));
			setValue("txtPurRatePer", attr(basic, "PurRatePer"));
			setValue("txtCharPer", attr(basic, "CharPer"));
			setValue("txtCharValue", attr(basic, "CharValue"));
			setValue("txtMarPer", attr(basic, "MarPer"));
			setValue("txtMarValue", attr(basic, "MarValue"));
			setValue("txtTotalPrice", attr(basic, "TotPrice"));
		}
		renderDiscountRows(root);
		return false;
	}

	window.PurRecSubmit = PurRecSubmit;
	window.SalSubmit = SalSubmit;
	window.resetVendor = resetVendor;
	window.PopulateSupplier = PopulateSupplier;
	window.OpenAlter = OpenAlter;
	window.OpenUoM = OpenUoM;
	window.CheckEntryDis = CheckEntryDis;
	window.CheckEntryVal = CheckEntryVal;
	window.resetDiscount = resetDiscount;
	window.checkNumbers = checkNumbers;
	window.AssaignValue = AssaignValue;
	window.CalcValue = CalcValue;
	window.DisplayData = DisplayData;
}(window, document));
