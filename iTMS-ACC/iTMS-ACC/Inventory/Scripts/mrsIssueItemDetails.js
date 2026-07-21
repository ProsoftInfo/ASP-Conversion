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
		return frm && frm.elements && frm.elements[name] || document.getElementById(name) || window[name] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function checkNumbers(value) {
		return /^([0-9]+(\.[0-9]*)?|\.[0-9]+)$/.test(trim(value));
	}

	function pad2(value) {
		return value < 10 ? "0" + value : String(value);
	}

	function parseDate(value) {
		var text = trim(value);
		var match;
		var year;
		var date;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			year = Number(match[3]);
			if (year < 100) {
				year += 2000;
			}
			date = new Date(year, Number(match[2]) - 1, Number(match[1]));
			return date.getFullYear() === year && date.getMonth() === Number(match[2]) - 1 && date.getDate() === Number(match[1]) ? date : null;
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
		if (match) {
			date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
			return date.getFullYear() === Number(match[1]) && date.getMonth() === Number(match[2]) - 1 && date.getDate() === Number(match[3]) ? date : null;
		}
		return null;
	}

	function formatDate(value) {
		var date = parseDate(value);
		return date ? pad2(date.getDate()) + "/" + pad2(date.getMonth() + 1) + "/" + date.getFullYear() : trim(value);
	}

	function dateDiffDays(start, end) {
		var startDate = parseDate(start);
		var endDate = parseDate(end);
		if (!startDate || !endDate) {
			return 0;
		}
		return Math.floor((endDate.getTime() - startDate.getTime()) / 86400000);
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

	function loadXml(islandName, text) {
		var island = xmlIsland(islandName);
		if (island && typeof island.loadXML === "function") {
			return island.loadXML(text);
		}
		return false;
	}

	function serializeXml(value) {
		var doc = value && value.nodeType === 9 ? value : xmlDocument(value);
		return doc ? new XMLSerializer().serializeToString(doc) : "";
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

	function importForDocument(doc, node) {
		if (!doc || !node) {
			return null;
		}
		return doc.importNode ? doc.importNode(node, true) : node.cloneNode(true);
	}

	function createElement(doc, name) {
		return doc.createElement(name);
	}

	function clearChildrenByName(node, name) {
		var children = elementChildren(node, name);
		for (var i = 0; i < children.length; i += 1) {
			node.removeChild(children[i]);
		}
	}

	function resetIslandRoot(islandName, rootName) {
		var doc = xmlDocument(xmlIsland(islandName));
		while (doc && doc.firstChild) {
			doc.removeChild(doc.firstChild);
		}
		if (doc) {
			doc.appendChild(doc.createElement(rootName));
		}
		return xmlRoot(xmlIsland(islandName));
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

	function fieldByPrefix(prefix) {
		var frm = form();
		var controls = frm && frm.elements || [];
		for (var i = 0; i < controls.length; i += 1) {
			if (String(controls[i].name || "").indexOf(prefix) === 0) {
				return controls[i];
			}
		}
		return null;
	}

	function controlValue(control) {
		if (!control) {
			return "";
		}
		if (control.options && control.selectedIndex >= 0) {
			return control.options[control.selectedIndex].value;
		}
		return control.value || "";
	}

	function issueTypeValue() {
		var control = field("selIssType");
		if (!control) {
			return "";
		}
		if (control.options) {
			return selectedValue("selIssType");
		}
		return control.checked ? "M" : "F";
	}

	function checkedValue(name, checkedValue, uncheckedValue) {
		var control = field(name);
		return control && control.checked ? checkedValue : uncheckedValue;
	}

	function islandExists(name) {
		ensureCompat();
		return !!(document.getElementById(name) || window[name] || document[name]);
	}

	function detailsIslandName() {
		return islandExists("OutData2") ? "OutData2" : "OutData";
	}

	function detailsIsland() {
		return xmlIsland(detailsIslandName());
	}

	function detailsRoot() {
		return xmlRoot(detailsIsland());
	}

	function getDatePicker() {
		ensureCompat();
		return field("ctlIssDate");
	}

	function getDatePickerValue() {
		var control = getDatePicker();
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		return formatDate(control.value || "");
	}

	function setDatePickerValue(value) {
		var control = getDatePicker();
		if (!control || !value) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else {
			control.value = value;
		}
	}

	function openModal(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		if (typeof window.showModalDialog === "function") {
			if (callback) {
				callback(window.showModalDialog(url, args, features));
			}
			return null;
		}
		window.open(url, "_blank");
		return null;
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function syncPost(url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignoreHeader) {}
		xhr.send(body == null ? null : body);
		return xhr;
	}

	function issueQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyPPX" + classCode + "X" + itemCode + "X" + entryNo);
	}

	function returnTypeField(itemCode, classCode, entryNo) {
		return field("selReturnZ" + itemCode + "Z" + classCode + "Z" + entryNo);
	}

	function approvedQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyA" + classCode + "A" + itemCode + "A" + entryNo);
	}

	function pendingQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyPenX" + classCode + "X" + itemCode + "X" + entryNo);
	}

	function transferQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyTraX" + classCode + "X" + itemCode + "X" + entryNo);
	}

	function prQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyPrX" + classCode + "X" + itemCode + "X" + entryNo);
	}

	function totalQtyField(classCode, itemCode, entryNo) {
		return field("txtQtyTotX" + classCode + "X" + itemCode + "X" + entryNo);
	}

	function selectedIndex(name) {
		var control = field(name);
		return control && typeof control.selectedIndex === "number" ? control.selectedIndex : -1;
	}

	function storeValueFor(classCode, itemCode, entryNo, approvedControl) {
		var frm = form();
		var controls = frm && frm.elements || [];
		var prefix = "txtStockZ" + classCode + "Z" + itemCode + "Z" + entryNo + "Z";
		var parts;
		for (var i = 0; i < controls.length; i += 1) {
			if (controls[i] === approvedControl && i >= 2 && /^hStoA/i.test(controls[i - 2].name || "")) {
				return controls[i - 2].value || "";
			}
		}
		for (var j = 0; j < controls.length; j += 1) {
			if (String(controls[j].name || "").indexOf(prefix) === 0) {
				parts = String(controls[j].name).split("Z");
				return (parts[5] || "") + ":" + (parts[6] || "");
			}
		}
		return "";
	}

	function firstRootWithItems() {
		var root1 = xmlRoot(xmlIsland("OutData1"));
		var root2 = xmlRoot(xmlIsland("OutData2"));
		if (root1 && (elementChildren(root1, "ITEMDETAILS").length || elementChildren(root1, "ITEM").length)) {
			return root1;
		}
		if (root2 && (elementChildren(root2, "ITEMDETAILS").length || elementChildren(root2, "ITEM").length)) {
			return root2;
		}
		return root1 || root2;
	}

	function matchingItem(root, entryNo, itemCode, classCode) {
		var nodes = elementChildren(root);
		for (var i = 0; i < nodes.length; i += 1) {
			if ((String(nodes[i].nodeName).toLowerCase() === "itemdetails" || String(nodes[i].nodeName).toLowerCase() === "item") &&
					(getAttr(nodes[i], "ENTRYNO") === trim(entryNo) || getAttr(nodes[i], "ItemEntNo") === trim(entryNo)) &&
					(getAttr(nodes[i], "ITEMCODE") === trim(itemCode) || getAttr(nodes[i], "ITMCODE") === trim(itemCode)) &&
					(getAttr(nodes[i], "CLASSCODE") === trim(classCode) || getAttr(nodes[i], "CLACODE") === trim(classCode))) {
				return nodes[i];
			}
		}
		return null;
	}

	function parseAttributeList(value) {
		var text = trim(value).replace(/#/g, "$");
		var parts = text ? text.split(",") : [];
		var attList = [];
		var attId = [];
		var entry;
		var left;
		var attrParts;
		for (var i = 0; i < parts.length; i += 1) {
			entry = parts[i].split("@");
			left = entry[0] || "";
			attrParts = left.split("$");
			if (entry.length > 1) {
				if (trim(attrParts[1]) && trim(attrParts[1]) !== "0") {
					attList.push(attrParts[0]);
					attId.push(attrParts[1]);
				}
			} else if (trim(left)) {
				attList.push(left);
			}
		}
		return {
			attributeList: attList.join(","),
			attributeId: attId.join(",")
		};
	}

	function updatePickQuantity(itemCode, classCode, entryNo) {
		var root = xmlRoot(xmlIsland("OutData1"));
		var item = matchingItem(root, entryNo, itemCode, classCode);
		var picks = elementChildren(item, "Pick");
		var qty = issueQtyField(classCode, itemCode, entryNo);
		var total;
		if (!item) {
			return;
		}
		if (getAttr(root, "DONE") === "NO") {
			return;
		}
		if (picks.length) {
			total = getAttr(picks[0], "TOT") || "0";
			if (qty) {
				qty.value = total;
				qty.readOnly = toNumber(total) !== 0;
				qty.disabled = false;
				qty.className = toNumber(total) !== 0 ? "FormElemRead" : "FormElem";
			}
			setAttr(item, "ONLYLOT", getAttr(picks[0], "ONLYLOT"));
		} else if (qty) {
			qty.value = "0";
			qty.readOnly = false;
			qty.disabled = false;
			qty.className = "FormElem";
		}
	}

	function appendDefaultPick(itemNode, outDoc, issueQty) {
		var itemCode = getAttr(itemNode, "ITMCODE");
		var classCode = getAttr(itemNode, "CLACODE");
		var receiptNumbering = trim(syncGet("../../Common/GetItemRcptNumbering.asp?ItemCode=" + encodeURIComponent(itemCode)).responseText || "");
		var storeResponse = syncGet("../../Common/GetStoreDetailsForItem.asp?ItemCode=" + encodeURIComponent(itemCode) + "&ClassCode=" + encodeURIComponent(classCode));
		var storeDoc = new DOMParser().parseFromString(storeResponse.responseText || "<Root/>", "application/xml");
		var stores = elementChildren(storeDoc.documentElement);
		var pick;
		var detail;
		if (!stores.length || stores.length > 1) {
			alert("Multiple Stores available please select the Issue Quantity for specify store");
			return false;
		}
		pick = createElement(outDoc, "Pick");
		setAttr(pick, "TOT", issueQty);
		setAttr(pick, "NoofPack", "0");
		detail = createElement(outDoc, receiptNumbering === "N" ? "STORE" : "PICK");
		setAttr(detail, "LOC", getAttr(stores[0], "LocNo"));
		setAttr(detail, "BIN", getAttr(stores[0], "BinNo"));
		setAttr(detail, "LOTNO", "N/A");
		setAttr(detail, "INVRECNO", "");
		setAttr(detail, "QTYISS", issueQty);
		setAttr(detail, "NoofPack", "0");
		pick.appendChild(detail);
		itemNode.appendChild(pick);
		return true;
	}

	function appendImmediateSchedule(itemNode, outDoc, nodeName, selectName, scheduleDate) {
		var select = field(selectName) || fieldByPrefix(selectName);
		var schedule;
		if (!select || select.selectedIndex !== 0) {
			return;
		}
		clearChildrenByName(itemNode, nodeName);
		clearChildrenByName(itemNode, "ScheduleDetails");
		schedule = createElement(outDoc, nodeName);
		setAttr(schedule, "STYPE", controlValue(select) || "ID");
		setAttr(schedule, "SVALUE", scheduleDate || formatDate(new Date()));
		itemNode.appendChild(schedule);
	}

	function buildIssueXml(todaysDate) {
		var sourceRoot = firstRootWithItems();
		var sourceItems = elementChildren(sourceRoot);
		var targetRoot = resetIslandRoot("OutData1", "ISSTYPE");
		var targetDoc = xmlDocument(xmlIsland("OutData1"));
		var retCount = 0;
		var issueDate = getDatePickerValue();
		var refType = selectedValue("selRefName");
		var issueMode = field("hIssMode") ? field("hIssMode").value : "N";
		var parsedAttrs;
		var itemNode;
		var itemCode;
		var classCode;
		var entryNo;
		var itemName;
		var issueQty;
		var transferQty;
		var prQty;
		var totalQty;
		var approvedQty;
		var qtyControl;
		var approvedControl;
		var pendingControl;
		var transferControl;
		var prControl;
		var totalControl;
		var isMRSItem;
		var nodeName;
		var retControl;
		var retValue;
		var converted;
		var children;
		var child;
		var auxItem;
		var auxChildren;

		if (dateDiffDays(todaysDate, issueDate) > 0) {
			alert("Issue Date should be less than or equal to Today's Date");
			return null;
		}
		if (field("hAutoConsumption") && field("hAutoConsumption").value === "Y") {
			if (field("selAccHead") && field("selAccHead").selectedIndex <= 1) {
				alert("Select the Account Head");
				field("selAccHead").focus();
				return null;
			}
		}
		if (trim(field("hType") && field("hType").value) === "SUB" && !trim(field("hIssueToCode") && field("hIssueToCode").value)) {
			alert("Please Select Issued To");
			return null;
		}

		for (var i = 0; i < sourceItems.length; i += 1) {
			itemNode = sourceItems[i];
			if (String(itemNode.nodeName).toLowerCase() === "subcontract") {
				targetRoot.appendChild(importForDocument(targetDoc, itemNode));
				continue;
			}
			if (String(itemNode.nodeName).toLowerCase() !== "itemdetails" && String(itemNode.nodeName).toLowerCase() !== "item") {
				continue;
			}
			entryNo = getAttr(itemNode, "ENTRYNO");
			itemCode = getAttr(itemNode, "ITEMCODE") || getAttr(itemNode, "ITMCODE");
			classCode = getAttr(itemNode, "CLASSCODE") || getAttr(itemNode, "CLACODE");
			itemName = getAttr(itemNode, "ITEMNAME") || getAttr(itemNode, "ITMNAME");
			nodeName = String(itemNode.nodeName).toLowerCase();
			isMRSItem = nodeName === "item" && !!getAttr(itemNode, "MRSNO");
			qtyControl = issueQtyField(classCode, itemCode, entryNo);
			approvedControl = approvedQtyField(classCode, itemCode, entryNo);
			pendingControl = pendingQtyField(classCode, itemCode, entryNo);
			transferControl = transferQtyField(classCode, itemCode, entryNo);
			prControl = prQtyField(classCode, itemCode, entryNo);
			totalControl = totalQtyField(classCode, itemCode, entryNo);
			issueQty = trim(qtyControl && qtyControl.value) || "0";
			transferQty = trim(transferControl && transferControl.value) || "0";
			prQty = trim(prControl && prControl.value) || "0";
			approvedQty = trim(approvedControl && approvedControl.value) || getAttr(itemNode, "REQQTY") || "0";
			totalQty = toNumber(issueQty) + toNumber(transferQty) + toNumber(prQty);
			if (!checkNumbers(issueQty) || !checkNumbers(transferQty) || !checkNumbers(prQty)) {
				alert("Enter Numerals Only");
				if (qtyControl && !checkNumbers(issueQty)) {
					qtyControl.focus();
				} else if (transferControl && !checkNumbers(transferQty)) {
					transferControl.focus();
				} else if (prControl && !checkNumbers(prQty)) {
					prControl.focus();
				}
				return null;
			}
			if (totalControl) {
				totalControl.value = String(totalQty);
			}
			if (isMRSItem) {
				if (pendingControl && totalQty > toNumber(pendingControl.value)) {
					alert("Total Quantity should be less than or equal to Quantity Pending (" + pendingControl.value + ")");
					if (qtyControl) {
						qtyControl.focus();
					}
					return null;
				}
			} else {
				if (toNumber(issueQty) <= 0) {
					alert("Issue Quantity Should be Greater then Zero for " + itemName);
					return null;
				}
			}
			if (refType === "11" && issueMode !== "E") {
				if (toNumber(issueQty) > toNumber(field("txtRequestedZ" + itemCode + "Z" + classCode + "Z" + entryNo) && field("txtRequestedZ" + itemCode + "Z" + classCode + "Z" + entryNo).value) - toNumber(field("txtToIssueZ" + itemCode + "Z" + classCode + "Z" + entryNo) && field("txtToIssueZ" + itemCode + "Z" + classCode + "Z" + entryNo).value)) {
					alert("Issue Quantity Must be Less than or Equal to Request Quantity");
					return null;
				}
			}

			parsedAttrs = parseAttributeList(getAttr(itemNode, "ATTRIBUTELIST"));
			converted = createElement(targetDoc, "ITEM");
			setAttr(converted, "ENTRYNO", entryNo);
			setAttr(converted, "ITMCODE", itemCode);
			setAttr(converted, "CLACODE", classCode);
			setAttr(converted, "ITMNAME", itemName);
			setAttr(converted, "SSTORE", isMRSItem ? storeValueFor(classCode, itemCode, entryNo, approvedControl) : "");
			setAttr(converted, "REQQTY", isMRSItem ? approvedQty : "0");
			setAttr(converted, "REQBY", field("txtRecBy") ? field("txtRecBy").value : "");
			setAttr(converted, "REMARKS", field("Remarks") ? field("Remarks").value : "");
			setAttr(converted, "ITEMTYPE", field("hItemType") ? field("hItemType").value : field("hItmType") ? field("hItmType").value : "");
			setAttr(converted, "ISSUEDATE", issueDate);
			setAttr(converted, "ISSQTY", issueQty);
			setAttr(converted, "TRAQTY", transferQty);
			setAttr(converted, "PRQTY", prQty);
			setAttr(converted, "IVALUE", isMRSItem ? getAttr(itemNode, "IVALUE") || "0" : issueQty);
			setAttr(converted, "ORGCODE", getAttr(itemNode, "UNIT") || field("hOrgID") && field("hOrgID").value || field("hUnit") && field("hUnit").value || "");
			setAttr(converted, "MRSNO", getAttr(itemNode, "MRSNO") || (refType === "11" && field("hRefNo") ? field("hRefNo").value : ""));
			setAttr(converted, "MRSDATE", getAttr(itemNode, "MRSDATE") || (refType === "11" && field("hRefDate") ? field("hRefDate").value : ""));
			setAttr(converted, "ATTRIBUTELIST", isMRSItem ? getAttr(itemNode, "ATTRIBUTELIST") : parsedAttrs.attributeId || parsedAttrs.attributeList);
			setAttr(converted, "CREATEDBY", getAttr(itemNode, "CREATEDBY") || (field("hUserId") ? field("hUserId").value : ""));
			setAttr(converted, "CREATEDON", getAttr(itemNode, "CREATEDON") || issueDate);
			setAttr(converted, "RefNo", getAttr(itemNode, "RefNo") || getAttr(itemNode, "MRSNO"));
			setAttr(converted, "ONLYLOT", getAttr(itemNode, "ONLYLOT"));

			retControl = returnTypeField(itemCode, classCode, entryNo);
			retValue = selectedValue("selReturnZ" + itemCode + "Z" + classCode + "Z" + entryNo) || "N";
			if (retControl && retControl.selectedIndex > 0) {
				retCount += 1;
			}
			if (retValue === "N") {
				setAttr(converted, "RETURNABLE", "N");
				setAttr(converted, "RETURNITEM", "S");
			} else if (retValue === "Y") {
				setAttr(converted, "RETURNABLE", "Y");
				setAttr(converted, "RETURNITEM", "S");
			} else {
				setAttr(converted, "RETURNABLE", "Y");
				setAttr(converted, "RETURNITEM", "D");
			}
			setAttr(converted, "MatType", "");

			children = elementChildren(itemNode);
			for (var c = 0; c < children.length; c += 1) {
				child = children[c];
				if (/^(pick|adddet|stschedule|prschedule|scheduledetails)$/i.test(child.nodeName)) {
					converted.appendChild(importForDocument(targetDoc, child));
				}
			}
			auxItem = matchingItem(detailsRoot(), entryNo, itemCode, classCode);
			auxChildren = elementChildren(auxItem);
			for (var ac = 0; ac < auxChildren.length; ac += 1) {
				child = auxChildren[ac];
				if (/^(stschedule|prschedule|scheduledetails)$/i.test(child.nodeName)) {
					converted.appendChild(importForDocument(targetDoc, child));
				}
			}
			if (isMRSItem && toNumber(transferQty) > 0) {
				appendImmediateSchedule(converted, targetDoc, "STSchedule", "selSTSchZ" + classCode + "Z" + itemCode + "Z" + entryNo + "Z", todaysDate);
			}
			if (isMRSItem && toNumber(prQty) > 0) {
				appendImmediateSchedule(converted, targetDoc, "PRSchedule", "selPRSchZ" + classCode + "Z" + itemCode + "Z" + entryNo + "Z", todaysDate);
			}
			if (toNumber(issueQty) > 0 && !elementChildren(converted, "Pick").length && !appendDefaultPick(converted, targetDoc, issueQty)) {
				return null;
			}
			targetRoot.appendChild(converted);
		}

		if (trim(field("hType") && field("hType").value) === "SUB" && retCount <= 0) {
			alert("SubContract Issue Should have minimum One Returnable Item");
			return null;
		}
		if (!elementChildren(targetRoot, "ITEM").length) {
			alert("No Items are available for Issue");
			return null;
		}
		return targetRoot;
	}

	function decorateIssueRoot(root) {
		var issueType = issueTypeValue();
		var refType = selectedValue("selRefName");
		setAttr(root, "ISSTYPE", issueType);
		setAttr(root, "ISSTOTYPE", field("hIssueToType") ? field("hIssueToType").value : "");
		setAttr(root, "ISSTOCODE", field("hIssueToCode") ? field("hIssueToCode").value : "");
		setAttr(root, "ISSTOSUBCODE", field("hIssueToSubCode") ? field("hIssueToSubCode").value : "");
		setAttr(root, "ISSFORCODE", field("hUsage") ? field("hUsage").value : "");
		setAttr(root, "ISSFORTYPE", field("hIssForType") ? field("hIssForType").value : "");
		setAttr(root, "PARTYCODE", field("hPartyCode") ? field("hPartyCode").value : "");
		setAttr(root, "POConfirm", "N");
		setAttr(root, "SInvConfirm", "N");
		setAttr(root, "Invoice", "A");
		setAttr(root, "GPConfirm", "N");
		setAttr(root, "ProConfirm", "N");
		setAttr(root, "MCallFrom", "MRIssue");
		setAttr(root, "RedirectTo", "ISSUEMGMT.ASP");
		setAttr(root, "MRNo", "");
		setAttr(root, "AppRefType", refType && refType !== "N" ? refType : "");
		setAttr(root, "AppRefNo", field("hRefNo") ? field("hRefNo").value : "");
		setAttr(root, "AppRefDate", field("hRefDate") ? field("hRefDate").value : "");
		setAttr(root, "ConsumptionAccHead", selectedValue("selAccHead"));
		setAttr(root, "IssueToCode", field("hIssueToCode") ? field("hIssueToCode").value : "");
		setAttr(root, "PickPackFlag", field("hPickPackFlag") ? field("hPickPackFlag").value : "");
		setAttr(root, "IssFrom", field("hIssFrom") ? field("hIssFrom").value : "");
		setAttr(root, "Returnable", "N");
		setAttr(root, "ReturnItem", "S");
		setAttr(root, "TYPE", field("hType") ? field("hType").value : "");
	}

	window.checkNumbers = checkNumbers;

	window.SetDate = function (dateValue) {
		var min = field("hMinDate") && field("hMinDate").value;
		var max = field("hMaxDate") && field("hMaxDate").value;
		if (dateDiffDays(min, dateValue) < 0 || dateDiffDays(dateValue, max) < 0) {
			setDatePickerValue(max);
		} else {
			setDatePickerValue(dateValue);
		}
	};

	window.MinDate = function () {
		var min = field("hMinDate") && field("hMinDate").value;
		var max = field("hMaxDate") && field("hMaxDate").value;
		var value = getDatePickerValue();
		if (dateDiffDays(min, value) < 0 || dateDiffDays(value, max) < 0) {
			alert("Issue Date Should Be With in the Financial Year " + min + " and " + max);
			setDatePickerValue(max);
		}
	};

	window.Back = function () {
		form().action = "MRSMGMTLIST.ASP?hCheck=I";
		form().submit();
	};

	window.DisplayDet = function () {};
	window.LoadData = function () {};

	window.DisplayStock = function (obj) {
		var parts = String(obj && obj.name || "").split("Z");
		var tempValues = (parts[2] || "") + ":" + (parts[1] || "") + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + (parts[3] || "") + ":" + (parts[4] || "") + ":" + (parts[5] || "") + ":" + (parts[6] || "");
		openModal("itmStockPoP.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:450px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No");
	};

	window.GetAddDetails = function (orgId, classCode, itemCode, mrsNo) {
		var qty = field("txtQtyPPX" + classCode + "x" + itemCode) || field("txtQtyPPX" + classCode + "X" + itemCode);
		var tempValues = classCode + "|" + itemCode + "|" + orgId + "|" + mrsNo + "|" + (qty ? qty.value : "");
		openModal("IssueAddDetails.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:400px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No");
	};

	window.CheckSch = function (obj) {
		var parts = String(obj && obj.name || "").split(":");
		var itemCode = parts[2] || "";
		var classCode = parts[1] || "";
		var entryNo = parts[3] || "";
		var optionName = parts[4] || "";
		var qty = field("txtQtyA" + classCode + "A" + itemCode + "A" + entryNo);
		var tempValues = (qty ? qty.value : "") + ":" + itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + (field("hOrgID") ? field("hOrgID").value : "") + ":" + entryNo + ":" + optionName;
		openModal("mrsIssueSchedulePoP.asp?sTemp=" + encodeURIComponent(tempValues), xmlIsland("Data"), "dialogHeight:480px;dialogWidth:390px;center:Yes;help:No;resizable:No;status:No");
	};

	window.GetSch = function (obj) {
		var parts = String(obj && obj.name || "").split("X");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var sch = field("hSchA" + classCode + "A" + itemCode + "A" + entryNo);
		var tempValues;
		if (!sch || sch.value !== "S" || trim(obj.value) === "" || toNumber(obj.value) === 0) {
			return;
		}
		tempValues = obj.value + ":" + itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + (field("hOrgID") ? field("hOrgID").value : "") + ":" + entryNo;
		openModal("mrsIssueScheduleEntryPoP.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:490px;dialogWidth:330px;center:Yes;help:No;resizable:No;status:No");
	};

	window.CheckQty = function (obj) {
		var parts = String(obj && obj.name || "").split(":");
		var tempValues = (parts[2] || "") + ":" + (parts[1] || "") + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + (parts[3] || "") + ":" + (parts[4] || "");
		openModal("mrsIssueQtyParaPoP.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:385px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No");
	};

	window.DisplayItem = function (obj) {
		var parts = String(obj || "").split("A");
		var tempValues = parts.length > 3 ? parts[0] + "A" + parts[2] + "A" + parts[1] + "A" + parts[3] : obj;
		openModal("itmDetailsPop.asp?sTemp=" + encodeURIComponent(tempValues), "", "dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No");
	};

	window.CheckLot = function (obj, issueQty, attributeList) {
		var parts = String(obj && obj.name || "").split("Z");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var optionName = parts[4] || "";
		var hasStoreBin = parts.length > 7;
		var locCode = hasStoreBin ? parts[5] || "" : entryNo;
		var binCode = hasStoreBin ? parts[6] || "" : optionName;
		var orgId = hasStoreBin ? field("hOrgID") && field("hOrgID").value || parts[5] || "" : parts[5] || "";
		var parsedAttrs = parseAttributeList(attributeList || (hasStoreBin ? parts.slice(7).join("Z") : parts[6]) || "");
		var issueType = issueTypeValue();
		var packFlag = field("hPickPackFlag") ? field("hPickPackFlag").value : "";
		var usage = field("hUsage") ? field("hUsage").value : "";
		var popupName = islandExists("OutData2") ? "mrsPickDetailPoP.asp" : "MRPickDetPoP.asp";
		var tempValues = itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + entryNo + ":" + optionName + ":" + locCode + ":" + binCode + ":" + (issueQty || "") + ":" + issueType + ":" + usage + ":" + orgId + ":" + parsedAttrs.attributeList + ":" + parsedAttrs.attributeId;
		openModal(popupName + "?sTemp=" + encodeURIComponent(tempValues) + "&AttributeList=" + encodeURIComponent(attributeList || "") + "&PickPackFlag=" + encodeURIComponent(packFlag), xmlIsland("OutData1"), "dialogHeight:350px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No", function () {
			updatePickQuantity(itemCode, classCode, entryNo);
		});
		return false;
	};

	window.CheckST = function (obj) {
		var parts = String(obj && obj.name || "").split(":");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var optionName = parts[4] || "";
		var orgId = parts[5] || "";
		var store = field("hStoName" + classCode + "A" + itemCode + "A" + entryNo);
		var tempValues = itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + (store ? store.value : "") + ":" + entryNo + ":" + optionName + ":" + orgId;
		openModal("mrsIssueSTPoP.asp?sTemp=" + encodeURIComponent(tempValues), detailsIsland(), "dialogHeight:330px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function () {
			var root = detailsRoot();
			var item = matchingItem(root, entryNo, itemCode, classCode);
			var qty = field("txtQtyTraX" + classCode + "X" + itemCode + "X" + entryNo) || field("txtQtyTraX" + classCode + "X" + itemCode);
			if (item && qty) {
				qty.value = getAttr(item, "TRAQTY") || getAttr(item, "ISSQTY");
			}
		});
	};

	window.CheckSTPRSch = function (obj, todaysDate, target) {
		var parts = String(obj && obj.name || "").split("Z");
		var classCode = parts[1] || "";
		var itemCode = parts[2] || "";
		var entryNo = parts[3] || "";
		var optionName = parts[4] || "";
		var qty = target === "ST" ? field("txtQtyTraX" + classCode + "X" + itemCode + "X" + entryNo) : field("txtQtyPrX" + classCode + "X" + itemCode + "X" + entryNo);
		var root = detailsRoot();
		var item = matchingItem(root, entryNo, itemCode, classCode);
		var doc = xmlDocument(detailsIsland());
		var scheduleName = target === "ST" ? "STSchedule" : "PRSchedule";
		var value;
		var node;
		if (!qty || trim(qty.value) === "") {
			alert("Enter Quantity");
			if (qty) {
				qty.focus();
			}
			obj.selectedIndex = 0;
			return false;
		}
		if (!checkNumbers(qty.value)) {
			alert("Enter Numerals Only");
			qty.focus();
			obj.selectedIndex = 0;
			return false;
		}
		if (toNumber(qty.value) === 0) {
			obj.selectedIndex = 0;
			return false;
		}
		if (!item && root && doc) {
			item = createElement(doc, "ITEM");
			setAttr(item, "ENTRYNO", entryNo);
			setAttr(item, "ITEMCODE", itemCode);
			setAttr(item, "ITMCODE", itemCode);
			setAttr(item, "CLASSCODE", classCode);
			setAttr(item, "CLACODE", classCode);
			root.appendChild(item);
		}
		if (!item) {
			return false;
		}
		clearChildrenByName(item, "STSchedule");
		clearChildrenByName(item, "PRSchedule");
		clearChildrenByName(item, "ScheduleDetails");
		if (obj.selectedIndex === 1 || obj.selectedIndex === 2) {
			value = prompt(obj.selectedIndex === 1 ? "Enter No of Days" : "Enter the Date", obj.selectedIndex === 1 ? "0" : "");
			if (value == null || trim(value) === "") {
				obj.selectedIndex = 0;
				return false;
			}
			if (obj.selectedIndex === 1 && !checkNumbers(value)) {
				alert("Enter Numerals Only");
				obj.selectedIndex = 0;
				return false;
			}
			if (obj.selectedIndex === 2 && (!parseDate(value) || dateDiffDays(todaysDate, value) < 0)) {
				alert("Invalid Date");
				obj.selectedIndex = 0;
				return false;
			}
			node = createElement(doc, scheduleName);
			setAttr(node, "STYPE", selectedValue(obj.name));
			setAttr(node, "SVALUE", value);
			item.appendChild(node);
		} else if (obj.selectedIndex === 3) {
			node = createElement(doc, scheduleName);
			setAttr(node, "STYPE", selectedValue(obj.name));
			setAttr(node, "SVALUE", "");
			item.appendChild(node);
			openModal((target === "ST" ? "mrsSTSchedulePoP.asp" : "mrsPRSchedulePoP.asp") + "?sTemp=" + encodeURIComponent(qty.value + ":" + itemCode + ":" + classCode + ":" + (field("hMRSNo") ? field("hMRSNo").value : "") + ":" + entryNo + ":" + optionName), detailsIsland(), "dialogHeight:510px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No");
		}
		return false;
	};

	window.RemoveXML = function () {
		var root = detailsRoot();
		var items = elementChildren(root, "ITEMDETAILS");
		for (var i = 0; i < items.length; i += 1) {
			while (items[i].firstChild) {
				items[i].removeChild(items[i].firstChild);
			}
		}
	};

	window.CheckSubmit = function (todaysDate) {
		var root = buildIssueXml(todaysDate);
		var issueMode = field("hIssMode") ? field("hIssMode").value : "N";
		var issueEntryNo = field("hIssEntryNo") ? field("hIssEntryNo").value : "";
		var callFrom = field("hCallFrom") ? field("hCallFrom").value : "";
		if (!root) {
			return false;
		}
		decorateIssueRoot(root);
		syncPost(issueMode === "E" ? "XMLSave.asp?Name=mrsIssueDataEdit&SessionFlag=true" : "XMLSave.asp?Name=mrsIssueData&SessionFlag=true", serializeXml(xmlIsland("OutData1")));
		form().action = issueMode === "E" ? "mrsIssueUpdate.asp?hCallFrom=" + encodeURIComponent(callFrom) + "&IssEntNo=" + encodeURIComponent(issueEntryNo) : "mrsIssueInsert.asp?hCallFrom=" + encodeURIComponent(callFrom);
		form().submit();
		return false;
	};
}(window, document));
