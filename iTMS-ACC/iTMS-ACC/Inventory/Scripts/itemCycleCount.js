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
		var item;
		var wanted;
		var i;
		if (!frm || !frm.elements) {
			return null;
		}
		item = frm.elements[name];
		if (item) {
			return item;
		}
		wanted = String(name).toLowerCase();
		for (i = 0; i < frm.elements.length; i += 1) {
			if (String(frm.elements[i].name || "").toLowerCase() === wanted) {
				return frm.elements[i];
			}
		}
		return null;
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

	function xmlDocument(name) {
		var object = xmlObject(name);
		var root = xmlRoot(object);
		return object && object.XMLDocument || object && object._doc || root && root.ownerDocument || null;
	}

	function childElements(node, name) {
		var result = [];
		var wanted = name ? String(name).toUpperCase() : "";
		var child;
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			child = node.childNodes[i];
			if (child.nodeType === 1 && (!wanted || String(child.nodeName || "").toUpperCase() === wanted)) {
				result.push(child);
			}
		}
		return result;
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function serializeXml(rootOrDoc) {
		var doc = rootOrDoc && rootOrDoc.nodeType === 9 ? rootOrDoc : rootOrDoc && rootOrDoc.ownerDocument;
		return new XMLSerializer().serializeToString(doc || rootOrDoc);
	}

	function postXml(url, rootOrDoc) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(serializeXml(rootOrDoc));
		return xhr.responseText || "";
	}

	function selectorSize() {
		var value = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup("1") : "ItemSelectRelPartyCommon.asp:500:850";
		var parts = String(value || "ItemSelectRelPartyCommon.asp:500:850").split(":");
		return {
			program: parts[0] || "ItemSelectRelPartyCommon.asp",
			height: parts[1] || "500",
			width: parts[2] || "850"
		};
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		window.open(url, "_blank");
		return null;
	}

	function selectedItemCodes() {
		var root = xmlRoot("ItemDetails");
		var codes = [];
		childElements(root, "Item").forEach(function (node) {
			if (attr(node, "ItemCode")) {
				codes.push(attr(node, "ItemCode"));
			}
		});
		return codes.join("?");
	}

	function continueItemSelection(returnedValue, afterDone) {
		var root = xmlRoot(returnedValue);
		var action = String(attr(root, "Action")).toUpperCase();
		var query;
		var size;
		if (!root || action === "CLOSE") {
			return;
		}
		if (action && action !== "DONE") {
			query = attr(root, "PassQuery");
			if (query) {
				size = selectorSize();
				openDialog("../../Common/" + size.program + "?" + query, xmlObject("ItemDetails"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (value) {
					continueItemSelection(value, afterDone);
				});
			}
			return;
		}
		afterDone();
	}

	function openItemSelector(afterDone, stockFlag) {
		var size = selectorSize();
		var url = "../../Common/" + size.program +
			"?orgID=" + encodeURIComponent(valueOf("hOrgCode")) +
			"&sIType=&Stock=" + encodeURIComponent(stockFlag || "N") +
			"&hSelectMode=M&Flag=1&hDispButt=Y&CallFrom=PUR";
		openDialog(url, xmlObject("ItemDetails"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (value) {
			continueItemSelection(value, afterDone);
		});
	}

	function AddItem() {
		openItemSelector(function () {
			setValue("hSelectItem", selectedItemCodes());
			form().submit();
		}, "N");
		return false;
	}

	function PopulateItem() {
		openItemSelector(function () {
			var codes = selectedItemCodes();
			var root = xmlRoot("ItemDetails");
			if (!codes) {
				return;
			}
			postXml("XMLSave.asp?SessionFlag=true&Value=ItemCycleCount&Folder=Master", root);
			form().action = "ItemCycleCount.asp?hSelectItem=" + encodeURIComponent(codes);
			form().submit();
		}, "Y");
		return false;
	}

	function ViewCycleCount(rowIndex, date, orgCode) {
		var value = field("Chkbox" + trim(rowIndex)) ? field("Chkbox" + trim(rowIndex)).value : "";
		var parts = value ? value.split(":") : [];
		var info;
		if (date != null) {
			form().action = "ItemCycleCount.asp?sTemp=" + encodeURIComponent(trim(rowIndex) + ":" + trim(date) + ":" + trim(orgCode) + ":E");
			form().submit();
			return false;
		}
		if (!parts.length) {
			return false;
		}
		info = [parts[0], parts[1], parts[2]].join(":");
		openDialog("ItemCycleCountHistoryPop.asp?Info=" + encodeURIComponent(info), "", "dialogHeight:500px;dialogWidth:400px;Status:No", function () {});
		return false;
	}

	function CheckAvailability(itemCode) {
		return childElements(xmlRoot("ItemDetails"), "Item").some(function (node) {
			return attr(node, "ItemCode") === itemCode;
		});
	}

	function ClearTable() {
		var table = document.getElementById("tblItem");
		if (!table) {
			return;
		}
		while (table.rows.length > 2) {
			table.deleteRow(2);
		}
	}

	function addTextCell(row, text, className, align) {
		var cell = row.insertCell(-1);
		cell.className = className || "ExcelDisplayCell";
		cell.align = align || "left";
		cell.textContent = text == null ? "" : String(text);
		return cell;
	}

	function encodedItemName(value) {
		return String(value || "").replace(/ /g, "-").replace(/"/g, "~~").replace(/'/g, "~");
	}

	function rowValue(node, unitId) {
		return [
			attr(node, "ItemCode"),
			attr(node, "ClassCode"),
			unitId,
			attr(node, "CompanyItemCode"),
			encodedItemName(attr(node, "ItemName")),
			"Y",
			"0",
			attr(node, "ItemStock"),
			attr(node, "ItemValue")
		].join(":");
	}

	function addCheckboxCell(row, name, value) {
		var cell = addTextCell(row, "", "ExcelDisplayCell", "center");
		var checkbox = document.createElement("input");
		checkbox.type = "checkbox";
		checkbox.name = name;
		checkbox.value = value;
		cell.appendChild(checkbox);
		return checkbox;
	}

	function addQuantityCell(row, name, value) {
		var cell = addTextCell(row, "", "ExcelDisplayCell", "right");
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.className = "FormElem";
		input.style.textAlign = "right";
		input.size = 12;
		input.value = trim(value) || "0";
		cell.appendChild(input);
	}

	function addViewCell(row, index) {
		var cell = addTextCell(row, "", "ExcelDisplayCell", "center");
		var img = document.createElement("img");
		img.border = "0";
		img.src = "../../assets/images/iTMS%20Icons/DetailsIcon.gif";
		img.width = 10;
		img.height = 10;
		img.style.cursor = "pointer";
		img.onclick = function () {
			ViewCycleCount(index);
		};
		cell.appendChild(img);
	}

	function DisplayDetails() {
		var root = xmlRoot("ItemDetails");
		var index = 0;
		var unitId = valueOf("hOrgCode");
		ClearTable();
		childElements(root, "Item").forEach(function (node) {
			var row = document.getElementById("tblItem").insertRow(-1);
			index += 1;
			addTextCell(row, index, "ExcelSerial", "center");
			addCheckboxCell(row, "Chkbox" + index, rowValue(node, unitId));
			addTextCell(row, attr(node, "CompanyItemCode"), "ExcelDisplayCell", "left");
			addTextCell(row, attr(node, "ItemName"), "ExcelDisplayCell", "left");
			addTextCell(row, attr(node, "ClassName"), "ExcelDisplayCell", "left");
			addTextCell(row, attr(node, "ItemStock"), "ExcelDisplayCell", "right");
			addTextCell(row, attr(node, "ItemValue"), "ExcelDisplayCell", "right");
			addQuantityCell(row, "txtCQty" + index, attr(node, "CCStock"));
			addViewCell(row, index);
		});
		setValue("hCtr", index);
	}

	function DeleteItem() {
		var root = xmlRoot("ItemDetails");
		var selected = {};
		var count = Number(valueOf("hCtr")) || 0;
		var item;
		var parts;
		for (var i = 1; i <= count; i += 1) {
			item = field("Chkbox" + i);
			if (item && item.checked) {
				parts = String(item.value || "").split(":");
				selected[parts[0]] = true;
			}
		}
		childElements(root, "Item").forEach(function (node) {
			if (selected[attr(node, "ItemCode")]) {
				root.removeChild(node);
			}
		});
		DisplayDetails();
		return false;
	}

	function cycleDateValue() {
		var item = field("ctlCCDate");
		if (!item) {
			return "";
		}
		if (typeof item.getdate === "function") {
			return item.getdate();
		}
		if (typeof item.getDate === "function") {
			return item.getDate();
		}
		return item.value;
	}

	function CheckSubmit() {
		var root = xmlRoot("CycleCount");
		var doc = xmlDocument("CycleCount");
		var dateValue = cycleDateValue();
		var entryNo = valueOf("hCycleCountEntryNo");
		var count = Number(valueOf("hCtr")) || 0;
		var item;
		var parts;
		var node;
		if (!root || !doc) {
			return false;
		}
		clearChildren(root);
		root.setAttribute("OrgCode", valueOf("hOrgCode"));
		root.setAttribute("CycleCountEntryNo", entryNo);
		root.setAttribute("CCDate", dateValue);
		root.setAttribute("Mode", trim(entryNo) ? "E" : "");
		for (var i = 1; i <= count; i += 1) {
			item = field("Chkbox" + i);
			if (!item || !trim(item.value)) {
				continue;
			}
			parts = String(item.value).split(":");
			node = doc.createElement("Child");
			node.setAttribute("ItemCode", parts[0] || "");
			node.setAttribute("ClassCode", parts[1] || "");
			node.setAttribute("OrgCode", parts[2] || "");
			node.setAttribute("CStock", parts[7] || "");
			node.setAttribute("CValue", parts[8] || "");
			node.setAttribute("CCQty", valueOf("txtCQty" + i));
			node.setAttribute("CCDate", dateValue);
			node.setAttribute("CCVal", "");
			root.appendChild(node);
		}
		postXml("XMLSave.asp?SessionFlag=True&Value=Inv_CycleCount_&Folder=Master", root);
		form().action = "ItemCycleCountInsert.asp";
		form().submit();
		return false;
	}

	function ShowCycleCountDet() {
		openDialog("CycleCountDetPop.asp", "", "dialogWidth:400px;dialogHeight:500px;", function () {});
		return false;
	}

	function Search() {
		setValue("hFromDate", valueOf("ctlFromDate"));
		setValue("hToDate", valueOf("ctlToDate"));
		form().submit();
		return false;
	}

	function Init() {
		if (field("ctlCCDate")) {
			setValue("ctlCCDate", valueOf("hCycleCountDate"));
		}
		if (field("ctlFromDate")) {
			setValue("ctlFromDate", valueOf("hFromDate"));
		}
		if (field("ctlToDate")) {
			setValue("ctlToDate", valueOf("hToDate"));
		}
	}

	function Help() {
		window.open("../HelpFiles/List.htm", "", "toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=No,scrollbars=yes,resizable=no,width=800px,height=500px,left=10,top=10");
		return false;
	}

	window.Search = Search;
	window.Init = Init;
	window.populateItem = PopulateItem;
	window.PopulateItem = PopulateItem;
	window.AddItem = AddItem;
	window.ViewCycleCount = ViewCycleCount;
	window.CheckAvailability = CheckAvailability;
	window.DeleteItem = DeleteItem;
	window.CheckSubmit = CheckSubmit;
	window.ShowCycleCountDet = ShowCycleCountDet;
	window.ClearTable = ClearTable;
	window.DisplayDetails = DisplayDetails;
	window.Help = Help;
}(window, document));
