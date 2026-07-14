(function () {
	"use strict";

	function loadCompat() {
		var currentScript;
		var src;
		var loader;
		if (window.ITMSModernCompat || document.querySelector('script[src*="itms-modern-compat.js"]')) {
			return;
		}
		currentScript = document.currentScript;
		src = currentScript ? currentScript.getAttribute("src") || "" : "";
		loader = document.createElement("script");
		loader.type = "text/javascript";
		loader.src = "/Scripts/itms-modern-compat.js";
		(document.head || document.documentElement).appendChild(loader);
	}

	loadCompat();

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value, decimals) {
		return toNumber(value).toFixed(decimals == null ? 2 : decimals);
	}

	function isNumeric(value) {
		return trim(value) !== "" && !isNaN(Number(String(value).replace(/,/g, "")));
	}

	function form() {
		return document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
	}

	function formField(name) {
		var frm = form();
		return frm && (frm.elements[name] || frm[name]) || null;
	}

	function fieldValue(name, fallback) {
		var field = formField(name);
		return field ? field.value : fallback;
	}

	function setFieldValue(name, value) {
		var field = formField(name);
		if (field) {
			field.value = value;
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || null;
	}

	function setVisible(id, visible, height, width) {
		var element = byId(id) || window[id];
		if (!element || !element.style) {
			return;
		}
		element.style.height = visible ? height : "1px";
		if (width) {
			element.style.width = visible ? width : "1px";
		}
		element.style.visibility = visible ? "visible" : "hidden";
	}

	function getXmlObject(name) {
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
		return window[name] || document[name] || null;
	}

	function getDocumentElement(xmlObject) {
		return xmlObject && xmlObject.documentElement || xmlObject && xmlObject.XMLDocument && xmlObject.XMLDocument.documentElement || null;
	}

	function getXmlRoot(name) {
		return getDocumentElement(getXmlObject(name));
	}

	function createXmlElement(xmlName, nodeName) {
		var xmlObject = getXmlObject(xmlName);
		if (xmlObject && typeof xmlObject.createElement === "function") {
			return xmlObject.createElement(nodeName);
		}
		if (xmlObject && xmlObject.XMLDocument) {
			return xmlObject.XMLDocument.createElement(nodeName);
		}
		return document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function childElements(node, nodeName) {
		var result = [];
		var name = nodeName && String(nodeName).toLowerCase();
		if (!node || !node.childNodes) {
			return result;
		}
		for (var i = 0; i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && (!name || String(node.childNodes[i].nodeName).toLowerCase() === name)) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function hasChildNodes(node) {
		return childElements(node).length > 0;
	}

	function attrValue(node, nameOrIndex) {
		var attr;
		if (!node || !node.attributes) {
			return "";
		}
		if (typeof nameOrIndex === "number") {
			attr = node.attributes.item(nameOrIndex);
			return attr ? attr.nodeValue : "";
		}
		return node.getAttribute(nameOrIndex) || "";
	}

	function setAttrValue(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value);
		}
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var nodes = [];
		if (!context) {
			return nodes;
		}
		if (context.selectNodes) {
			return context.selectNodes(expression);
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return nodes;
		}
		found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (var i = 0; i < found.snapshotLength; i += 1) {
			nodes.push(found.snapshotItem(i));
		}
		nodes.Item = function (index) {
			return this[index];
		};
		return nodes;
	}

	function getEntryRoot() {
		if (window.EntryRoot && window.EntryRoot.nodeType) {
			return window.EntryRoot;
		}
		return getXmlRoot("EntryData");
	}

	function setEntryRoot(node) {
		window.EntryRoot = node;
		return node;
	}

	function insertLegacyCell(row, iType, sName, sValue, sClass, sAlign, sValign, iSize, iMaxlen, iColspan, iRowspan, sOptions) {
		if (window.InsertCell && window.InsertCell !== insertLegacyCell) {
			return window.InsertCell(row, iType, sName, sValue, sClass, sAlign, sValign, iSize, iMaxlen, iColspan, iRowspan, sOptions);
		}
		var cell = row.insertCell();
		var input;
		if (Number(iType) === 1) {
			cell.innerHTML = sValue == null ? "" : String(sValue);
		} else {
			input = document.createElement(Number(iType) === 3 ? "input" : "input");
			input.type = Number(iType) === 3 ? "checkbox" : "text";
			input.name = sName || "";
			input.value = sValue == null ? "" : String(sValue);
			if (iSize) {
				input.size = Number(iSize);
			}
			if (iMaxlen) {
				input.maxLength = Number(iMaxlen);
			}
			input.className = Number(iType) === 4 ? "FormelemRead" : "Formelem";
			if (sOptions && String(sOptions).toLowerCase().indexOf("text-align:right") !== -1) {
				input.style.textAlign = "right";
			}
			cell.appendChild(input);
		}
		cell.className = sClass || "";
		if (sAlign) {
			cell.align = sAlign;
		}
		if (sValign) {
			cell.vAlign = sValign;
		}
		if (Number(iColspan)) {
			cell.colSpan = Number(iColspan);
		}
		if (Number(iRowspan)) {
			cell.rowSpan = Number(iRowspan);
		}
		return cell;
	}

	function clearTable(tableName, startIndex, keepCount) {
		if (window.ClearTable && window.ClearTable !== clearTable) {
			window.ClearTable(tableName, startIndex, keepCount);
			return;
		}
		var table = typeof tableName === "string" ? byId(tableName) : tableName;
		var start = Number(startIndex) || 0;
		var keep = Number(keepCount) || 0;
		if (!table || !table.rows) {
			return;
		}
		while (table.rows.length > start + keep) {
			table.deleteRow(start);
		}
	}

	window.SelectHead = function (sAccHead, sType, objHead, iHeadCount) {
		var valueParts;
		var target;
		var optionParts;
		if (String(sType) === "G") {
			for (var i = 0; i < Number(objHead.length); i += 1) {
				optionParts = String(objHead.options[i].value).split("?");
				if (optionParts[0] === String(sAccHead) || objHead.options[i].value === sAccHead) {
					objHead.selectedIndex = i;
					return;
				}
			}
			objHead.selectedIndex = Number(iHeadCount) + 1;
			return;
		}
		valueParts = String(sAccHead).split("?");
		target = trim(valueParts[0]) + "?" + trim(valueParts[1]);
		for (var j = 0; j < Number(objHead.length); j += 1) {
			if (objHead.options[j].value === target) {
				objHead.selectedIndex = j;
				return;
			}
		}
		objHead.selectedIndex = Number(iHeadCount) + 1;
	};

	window.CheckAccHead = function (nodRoot, sAccHead) {
		return selectNodes(nodRoot, "//AccHead[@No='" + sAccHead + "']").length > 0;
	};

	window.popAddAmount = function () {
		var entryRoot = getEntryRoot();
		var amountField = formField("txtAmount");
		var amount = toNumber(amountField && amountField.value);
		var headers = childElements(entryRoot);
		if (typeof window.checkFileds === "function" && !window.checkFileds()) {
			if (amountField) {
				amountField.value = "";
			}
			return;
		}
		headers.forEach(function (header) {
			var nodes = childElements(header);
			var total = amount;
			var ratioTotal = 0;
			var ratio;
			if (header.nodeName === "CostCenter" && nodes.length) {
				ratio = Math.round((100 / nodes.length) * 100) / 100;
				nodes.forEach(function (node, index) {
					var code = attrValue(node, "No");
					var nodeAmount = index < nodes.length - 1 ? Math.round(((ratio * amount) / 100) * 100) / 100 : total;
					var nodeRatio = index < nodes.length - 1 ? ratio : 100 - ratioTotal;
					setFieldValue("txtCCRatio" + code, nodeRatio);
					setFieldValue("txtCCAmount" + code, nodeAmount);
					setAttrValue(node, "Ratio", nodeRatio);
					setAttrValue(node, "Amount", nodeAmount);
					total -= nodeAmount;
					ratioTotal += nodeRatio;
				});
			}
			if (header.nodeName === "Analytical" && nodes.length) {
				total = amount;
				ratioTotal = 0;
				ratio = Math.round((100 / nodes.length) * 100) / 100;
				nodes.forEach(function (node, index) {
					var code = attrValue(node, "No");
					var groupCode = attrValue(node, "GroupCode");
					var nodeAmount = index < nodes.length - 1 ? Math.round(((ratio * amount) / 100) * 100) / 100 : total;
					var nodeRatio = index < nodes.length - 1 ? ratio : 100 - ratioTotal;
					setFieldValue("txtANALRatio" + code + "Z" + groupCode, nodeRatio);
					setFieldValue("txtANALAmount" + code + "Z" + groupCode, nodeAmount);
					setAttrValue(node, "Ratio", nodeRatio);
					setAttrValue(node, "Amount", nodeAmount);
					total -= nodeAmount;
					ratioTotal += nodeRatio;
				});
			}
			if (header.nodeName === "PayRec" && nodes.length) {
				nodes.forEach(function (node, index) {
					var code = attrValue(node, "No");
					var payNo = trim(attrValue(node, "PayableNo"));
					var adjust = toNumber(attrValue(node, "TransAmount")) - toNumber(attrValue(node, "AmtAdjusted")) - toNumber(attrValue(node, "AmtToAccount"));
					setFieldValue("txtDocAmount" + code + "Z" + payNo + "Z" + (index + 1), formatNumber(adjust, 2));
				});
			}
		});
	};

	window.clearXML = function () {
		var entryData = getXmlObject("EntryData");
		var root = entryData && entryData.createElement ? entryData.createElement("Entry") : createXmlElement("EntryData", "Entry");
		setAttrValue(root, "No", window.iEntryNo || "");
		setAttrValue(root, "CRDR", "0");
		setAttrValue(root, "Payto", "0");
		setAttrValue(root, "Amount", "0");
		setAttrValue(root, "AccUnit", "0");
		setAttrValue(root, "AccName", "");
		setAttrValue(root, "TdsAmount", "0");
		setAttrValue(root, "TDSElgi", "0");
		setAttrValue(root, "TdsPercentage", "0");
		setAttrValue(root, "PayRecAmount", "0");
		setEntryRoot(root);
	};

	window.ValidateAmount = function (dAmount, sName, dFrom, dTo) {
		var name = sName || "Amount";
		var min = dFrom == null ? 0 : Number(dFrom);
		var max = dTo == null ? 9999999999.99 : Number(dTo);
		if (trim(dAmount) === "") {
			alert(name + " Cannot be blank");
			return false;
		}
		if (!isNumeric(dAmount)) {
			alert("Enter Numeric values for " + name);
			return false;
		}
		if (toNumber(dAmount) <= min || toNumber(dAmount) > max) {
			alert(name + " should be >" + min + " and < " + max);
			return false;
		}
		return true;
	};

	window.showNarration = function (sBookCode) {
		var bookParts = String(fieldValue("selBook", "")).split("-");
		var orgId = fieldValue("hOrgId", fieldValue("horgID", ""));
		var bookNo = sBookCode + "?" + trim(bookParts[0]);
		var url = "NarrationSelection.asp?orgId=" + encodeURIComponent(orgId) + "&BookCode=" + encodeURIComponent(bookNo);
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", "dialogHeight:300px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function (sNarration) {
				if (sNarration) {
					setFieldValue("txtNarration", sNarration);
				}
			});
		}
	};

	window.popCostCenter = function (headerNode) {
		var nodes = childElements(headerNode);
		var table;
		if (!nodes.length) {
			window.setAnalDisplay("C", 0);
			return;
		}
		window.setAnalDisplay("C", 1);
		clearTable("tblCost", 1, 1);
		table = byId("tblCost");
		nodes.forEach(function (node, index) {
			var row = table && table.insertRow(index + 1);
			if (!row) {
				return;
			}
			insertLegacyCell(row, 1, "", index + 1, "ExcelSerial", "Center", "", 0, 0, 0, 0, "");
			insertLegacyCell(row, 1, "", attrValue(node, "ShortName") || attrValue(node, 2), "ExcelDisplayCell", "", "", 0, 0, 0, 0, "");
			insertLegacyCell(row, 2, "txtCCRatio" + (attrValue(node, "No") || attrValue(node, 0)), attrValue(node, "Ratio") || attrValue(node, 3), "ExcelInputCell", "", "", 6, 5, 0, 0, "");
			insertLegacyCell(row, 2, "txtCCAmount" + (attrValue(node, "No") || attrValue(node, 0)), attrValue(node, "Amount") || attrValue(node, 4), "ExcelInputCell", "", "", 12, 10, 0, 0, "");
		});
	};

	window.popAnalytical = function (headerNode) {
		var nodes = childElements(headerNode);
		var table;
		if (!nodes.length) {
			window.setAnalDisplay("A", 0);
			return;
		}
		window.setAnalDisplay("A", 1);
		clearTable("tblAnal", 1, 1);
		table = byId("tblAnal");
		nodes.forEach(function (node, index) {
			var code = attrValue(node, "No") || attrValue(node, 0);
			var groupCode = attrValue(node, "GroupCode");
			var row = table && table.insertRow(index + 1);
			if (!row) {
				return;
			}
			insertLegacyCell(row, 1, "", index + 1, "ExcelSerial", "Center", "", 0, 0, 0, 0, "");
			insertLegacyCell(row, 1, "", attrValue(node, "ShortName") || attrValue(node, 2), "ExcelDisplayCell", "", "", 0, 0, 0, 0, "");
			insertLegacyCell(row, 2, "txtANALRatio" + code + "Z" + groupCode, attrValue(node, "Ratio") || attrValue(node, 3), "ExcelInputCell", "", "", 6, 5, 0, 0, "");
			insertLegacyCell(row, 2, "txtANALAmount" + code + "Z" + groupCode, attrValue(node, "Amount") || attrValue(node, 4), "ExcelInputCell", "", "", 12, 10, 0, 0, "");
		});
	};

	window.showCCAnal = function (sOrgId, iAccCode, bCostCenter, bAnal) {
		var transNo = fieldValue("hTransNo", "");
		var entNo = fieldValue("hEntryNo", "");
		var entryRoot = getEntryRoot();
		var url;
		if (Number(bCostCenter) !== 1 && Number(bAnal) !== 1) {
			window.setADDDisplay(0);
			return;
		}
		url = "CCAnalysisSelection.asp?orgId=" + encodeURIComponent(sOrgId) + "&AccCode=" + encodeURIComponent(iAccCode) +
			"&TransNo=" + encodeURIComponent(transNo) + "&EntNo=" + encodeURIComponent(entNo);
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, entryRoot, "dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", function (node) {
				var children;
				if (!node || toNumber(attrValue(node, 0)) !== 1) {
					window.setADDDisplay(0);
					return;
				}
				window.setADDDisplay(1);
				children = childElements(node);
				children.forEach(function (child) {
					var imported = child;
					if (entryRoot && entryRoot.ownerDocument && child.ownerDocument !== entryRoot.ownerDocument) {
						imported = entryRoot.ownerDocument.importNode(child, true);
					}
					if (entryRoot) {
						entryRoot.appendChild(imported);
					}
					if (child.nodeName === "CostCenter") {
						window.popCostCenter(child);
					}
					if (child.nodeName === "Analytical") {
						window.popAnalytical(child);
					}
				});
			});
		}
	};

	window.popPayRec = function (headerNode) {
		renderPayRec(headerNode, false);
	};

	window.popPayRecAmd = function (headerNode) {
		renderPayRec(headerNode, true);
	};

	function renderPayRec(headerNode, amended) {
		var nodes = childElements(headerNode);
		var table;
		if (!nodes.length) {
			return;
		}
		if (amended) {
			window.setPayableDisplayAmd(1);
		} else {
			window.setPayableDisplay(1);
		}
		clearTable("tblPayable", 2, 1);
		table = byId("tblPayable");
		nodes.forEach(function (node, index) {
			var docNo = attrValue(node, "No");
			var invNo = attrValue(node, "InvNo");
			var invDate = attrValue(node, "InvDate");
			var transAmount = toNumber(attrValue(node, "TransAmount"));
			var adjusted = toNumber(attrValue(node, "AmtAdjusted"));
			var toAccount = toNumber(attrValue(node, "AmtToAccount"));
			var toAdjust = formatNumber(attrValue(node, "AmtToAdjust"), 2);
			var payNo = trim(attrValue(node, "PayableNo"));
			var total = formatNumber(transAmount - adjusted - toAccount, 2);
			var row = table && table.insertRow(index + 2);
			if (!row) {
				return;
			}
			if (amended) {
				invNo = invNo.replace(/Credit Notes/g, "CR.NO").replace(/Debit Notes/g, "DR.NO").replace(/Advance Receipts/g, "ADV REC").replace(/Advance Payaments/g, "ADV PAY");
				if (invNo.indexOf("SALE") === -1) {
					invNo += " DT:" + invDate;
				}
				invNo = invNo.replace(/SALE INV/g, "SAL");
			}
			insertLegacyCell(row, 1, "", index + 1, "ExcelSerial", "Center", "", 0, 0, 0, 0, "");
			insertLegacyCell(row, 1, "", invNo, "ExcelDisplayCell", "", "", 0, 0, 0, 0, "");
			if (!amended) {
				insertLegacyCell(row, 1, "", invDate, "ExcelDisplayCell", "", "", 0, 0, 0, 0, "");
			}
			insertLegacyCell(row, 1, "", formatNumber(transAmount, 2), "ExcelDisplayCell", "Right", "", 0, 0, 0, 0, "");
			insertLegacyCell(row, 1, "", formatNumber(adjusted, 2), "ExcelDisplayCell", "Right", "", 0, 0, 0, 0, "");
			insertLegacyCell(row, 1, "", formatNumber(toAccount, 2), "ExcelDisplayCell", "Right", "", 0, 0, 0, 0, "");
			insertLegacyCell(row, 1, "", total, "ExcelDisplayCell", "Right", "", 0, 0, 0, 0, "");
			insertLegacyCell(row, 2, "txtDocAmount" + docNo + "Z" + payNo + "Z" + (index + 1), toAdjust, "ExcelInputCell", "right", "", 12, 10, 0, 0, 'style="text-align:right"');
		});
	}

	window.setPayableDisplay = function (iFlag) {
		var visible = Number(iFlag) !== 0;
		setVisible("Disaddtional", visible, "115px");
		setVisible("DisPayable", visible, "110px");
	};

	window.setAnalDisplay = function (sDisplay, iFlag) {
		var visible = Number(iFlag) !== 0;
		setVisible(String(sDisplay) === "A" ? "DisAnal" : "DisCost", visible, "100px", "280px");
	};

	window.setADDDisplay = function (iFlag) {
		var flag = Number(iFlag);
		setVisible("Disaddtional", flag !== 0, "115px");
		if (flag === 1) {
			setVisible("DisCCANL", true, "114px");
		} else if (flag === 0) {
			setVisible("DisCCANL", false, "1px");
		}
	};

	window.GetGlHeadXml = function (sValue) {
		var parts = String(sValue).split(":");
		var root = getXmlRoot("AccHeadData");
		var elem = createXmlElement("AccHeadData", "AccHead");
		setAttrValue(elem, "No", trim(parts[0]));
		setAttrValue(elem, "CostCenter", trim(parts[2]));
		setAttrValue(elem, "Analytical", trim(parts[3]));
		setAttrValue(elem, "Name", trim(parts[5]));
		setAttrValue(elem, "Type", "G");
		setAttrValue(elem, "TransFlag", trim(parts[4]));
		if (root) {
			root.appendChild(elem);
		}
	};

	window.GetGlHeadXmlForSalAcc = function () {
		var root = getXmlRoot("AccHeadData");
		var elem = createXmlElement("AccHeadData", "AccHead");
		setAttrValue(elem, "No", fieldValue("hSalAccCode", ""));
		setAttrValue(elem, "CostCenter", "0");
		setAttrValue(elem, "Analytical", "0");
		setAttrValue(elem, "Name", fieldValue("hSalAccName", ""));
		setAttrValue(elem, "Type", "G");
		setAttrValue(elem, "TransFlag", "A");
		if (root) {
			root.appendChild(elem);
		}
	};

	window.GetPartyHeadXml = function (sCode, sName, sValue2) {
		var parts = String(sValue2).split(":");
		var root = getXmlRoot("AccHeadData");
		var elem = createXmlElement("AccHeadData", "AccHead");
		setAttrValue(elem, "No", trim(sCode));
		setAttrValue(elem, "Pay", trim(parts[0]));
		setAttrValue(elem, "Rec", trim(parts[1]));
		setAttrValue(elem, "Name", sName);
		setAttrValue(elem, "Type", "P");
		setAttrValue(elem, "Adv", trim(parts[2]));
		if (root) {
			root.appendChild(elem);
		}
	};

	window.SetApp = function (sType) {
		var user = formField("selUserId");
		if (user) {
			user.disabled = String(sType) !== "Y";
		}
	};

	window.CheckApp = function () {
		var frm = form();
		if (!frm) {
			return true;
		}
		if (frm.optApprove && frm.optApprove[0] && frm.optApprove[0].checked === true && frm.selUserId && frm.selUserId.selectedIndex === 0) {
			alert("Select Approver ");
			frm.selUserId.focus();
			return false;
		}
		if (frm.txtNarration && frm.txtNarration.value.length > 300) {
			alert("Narration Should be Less than 300 Characters ");
			frm.txtNarration.focus();
			return false;
		}
		if (frm.selAccUnitId && frm.selUnitId && String(frm.selAccUnitId.value) !== String(frm.selUnitId.value)) {
			alert("Created Unit and Accounting Unit is different!!");
			return false;
		}
		return true;
	};

	window.CheckFinDate = function () {
		var finFrom = toNumber(fieldValue("hFinFrm", 0));
		var finTo = toNumber(fieldValue("hFinTo", 0));
		var ctlDate = formField("ctlDate") || byId("ctlDate");
		var dateText = ctlDate && (typeof ctlDate.GetDate === "function" ? ctlDate.GetDate() : ctlDate.value);
		var parts = String(dateText || "").split("/");
		var current = toNumber((parts[2] || "") + (parts[1] || ""));
		if (current < finFrom || current > finTo) {
			alert("Voucher Date Should Be Between 01/04/" + String(finFrom).substring(0, 4) + " To 31/03/" + String(finTo).substring(0, 4));
			return false;
		}
		return true;
	};

	window.setPayableDisplayAmd = function (iFlag) {
		var visible = Number(iFlag) !== 0;
		setVisible("Disaddtional", visible, "115px");
		setVisible("DisPayable", visible, "90px");
	};

	window.ChkEnter = function (evt) {
		evt = evt || null;
		if (evt && evt.key === "Enter") {
			if (evt.preventDefault) {
				evt.preventDefault();
			}
			return false;
		}
		return true;
	};
}());
