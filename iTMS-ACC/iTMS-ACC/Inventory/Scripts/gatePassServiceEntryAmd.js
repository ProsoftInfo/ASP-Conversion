(function (window, document) {
	"use strict";

	var supplierCode = "";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function upper(value) {
		return trim(value).toUpperCase();
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		if (!frm) {
			return null;
		}
		return frm.elements[name] || frm.elements[name.toLowerCase()] || frm.elements[name.toUpperCase()] || null;
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function focusField(name, selectText) {
		var item = field(name);
		if (!item) {
			return;
		}
		if (selectText && item.select) {
			item.select();
		} else if (item.focus) {
			item.focus();
		}
	}

	function xmlRoot(value, fallback) {
		var candidate = value || fallback;
		if (!candidate) {
			return null;
		}
		if (candidate.documentElement) {
			return candidate.documentElement;
		}
		if (candidate.XMLDocument && candidate.XMLDocument.documentElement) {
			return candidate.XMLDocument.documentElement;
		}
		if (candidate._doc && candidate._doc.documentElement) {
			return candidate._doc.documentElement;
		}
		if (candidate.nodeType === 1) {
			return candidate;
		}
		return null;
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
		if (value.nodeType === 1 && value.ownerDocument) {
			return value.ownerDocument;
		}
		return null;
	}

	function childElements(node, name) {
		var wanted = name && String(name).toLowerCase();
		return Array.prototype.filter.call(node && node.childNodes || [], function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function getXmlText(xmlObject) {
		var target = xmlObject && (xmlObject.XMLDocument || xmlObject._doc || xmlObject);
		if (!target) {
			return "";
		}
		if (typeof target.xml === "string") {
			return target.xml;
		}
		return new XMLSerializer().serializeToString(target);
	}

	function openDialog(url, args, features, callback) {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function popupSize(kind, fallbackProgram, fallbackHeight, fallbackWidth) {
		var value = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup(kind) : "";
		var parts = String(value || "").split(":");
		return {
			program: parts[0] || fallbackProgram,
			height: parts[1] || fallbackHeight,
			width: parts[2] || fallbackWidth
		};
	}

	function tableElement(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function clearTableRows() {
		var table = tableElement("tblDetails");
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function appendCell(row, text, className, align) {
		var cell = row.insertCell();
		cell.className = className || "ExcelDisplayCell";
		cell.align = align || "left";
		cell.innerHTML = text == null ? "" : String(text);
		return cell;
	}

	function detailsNodes(root) {
		return childElements(root, "DETAILS");
	}

	function normalizeDescription(value) {
		return trim(value).replace(/\s+/g, "");
	}

	function hasDetailWithDescription(root, description) {
		var wanted = trim(description).replace(/'/g, " ");
		return detailsNodes(root).some(function (node) {
			return trim(getAttr(node, "DESC")) === wanted;
		});
	}

	function resetEntryFields() {
		setValue("txtDesc", "");
		setValue("txtQty", "");
		setValue("txtItemDesc", "");
		setValue("hItem", "");
		setValue("hClass", "");
		if (field("selUOM")) {
			field("selUOM").value = "select";
			field("selUOM").disabled = false;
		}
		setValue("txtReason", "SENT FOR REPAIRS - TO BE RETURNED");
		setValue("txtValue", "");
		if (field("ChkFormJJ")) {
			field("ChkFormJJ").checked = false;
		}
	}

	function renderDetails() {
		var root = window.OutData && window.OutData.documentElement;
		var table = tableElement("tblDetails");
		var rowIndex = 0;
		clearTableRows();
		if (!root || !table) {
			setValue("hRows", "0");
			return;
		}
		childElements(root).forEach(function (node) {
			var row;
			var deleteCheckbox;
			var formJjCheckbox;
			var description;
			if (upper(node.nodeName) === "SUPPLIER") {
				setValue("txtRefName", getAttr(node, "SuppName"));
				return;
			}
			if (upper(node.nodeName) !== "DETAILS") {
				return;
			}
			rowIndex += 1;
			description = getAttr(node, "DESC");
			row = table.insertRow(rowIndex);
			appendCell(row, rowIndex, "ExcelSerial", "center");
			deleteCheckbox = document.createElement("input");
			deleteCheckbox.type = "checkbox";
			deleteCheckbox.value = [getAttr(node, "ITEMCODE"), getAttr(node, "CLASSCODE"), normalizeDescription(description)].join(":");
			deleteCheckbox.name = "chkboxDel" + rowIndex;
			deleteCheckbox.className = "FormElem";
			appendCell(row, "", "ExcelDisplayCell", "left").appendChild(deleteCheckbox);
			appendCell(row, description, "ExcelDisplayCell", "left");
			appendCell(row, getAttr(node, "QTY"), "ExcelDisplayCell", "right");
			appendCell(row, getAttr(node, "UOM"), "ExcelDisplayCell", "right");
			appendCell(row, getAttr(node, "VALUE"), "ExcelDisplayCell", "right");
			formJjCheckbox = document.createElement("input");
			formJjCheckbox.type = "checkbox";
			formJjCheckbox.value = "Y";
			formJjCheckbox.name = "chkbox" + rowIndex;
			formJjCheckbox.className = "FormElem";
			formJjCheckbox.checked = getAttr(node, "FORMJJ") === "Y";
			appendCell(row, "", "ExcelDisplayCell", "right").appendChild(formJjCheckbox);
			appendCell(row, getAttr(node, "REASON"), "ExcelDisplayCell", "left");
		});
		setValue("hRows", rowIndex);
	}

	function processEntrySupplierSelection(root) {
		if (!root || !root.hasChildNodes()) {
			return;
		}
		childElements(root).forEach(function (node) {
			if (upper(node.nodeName) === "ENTRY") {
				setValue("txtRefName", getAttr(node, "RetField0"));
				supplierCode = getAttr(node, "RetField1");
			}
		});
	}

	function continueSelectionDialog(result, xmlIsland, size, processor) {
		var root = xmlRoot(result, xmlIsland);
		var action = upper(getAttr(root, "Action"));
		var passQuery = getAttr(root, "PassQuery");
		if (action && action !== "DONE" && action !== "CLOSE") {
			openDialog("../../Common/" + size.program + "?" + passQuery, xmlIsland, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (nextResult) {
				continueSelectionDialog(nextResult, xmlIsland, size, processor);
			});
			return;
		}
		if (action !== "CLOSE") {
			processor(root);
		}
	}

	function processItemSelection(root) {
		if (!root || !root.hasChildNodes()) {
			return;
		}
		childElements(root).some(function (node) {
			setValue("txtItemDesc", getAttr(node, "ItemName"));
			setValue("hItem", getAttr(node, "ItemCode"));
			setValue("hClass", getAttr(node, "ClassCode"));
			if (field("selUOM")) {
				field("selUOM").value = getAttr(node, "StoresUoM");
				field("selUOM").disabled = true;
			}
			return true;
		});
	}

	window.DelItem = function () {
		var root = window.OutData && window.OutData.documentElement;
		var rows = parseInt(field("hRows") && field("hRows").value, 10) || 0;
		var removed = 0;
		var i;
		if (!root) {
			return;
		}
		for (i = rows; i >= 1; i -= 1) {
			var checkbox = field("chkboxDel" + i);
			var parts;
			if (!checkbox || !checkbox.checked) {
				continue;
			}
			parts = String(checkbox.value || "").split(":");
			childElements(root, "DETAILS").some(function (node) {
				var normalized = normalizeDescription(getAttr(node, "DESC"));
				var other = normalizeDescription(getAttr(node, "OTHERDESC"));
				var itemDesc = normalizeDescription(getAttr(node, "ITEMDESC"));
				if (
					getAttr(node, "ITEMCODE") === trim(parts[0]) &&
					getAttr(node, "CLASSCODE") === trim(parts[1]) &&
					(normalized === trim(parts[2]) || other === trim(parts[2]) || itemDesc === trim(parts[2]))
				) {
					root.removeChild(node);
					removed += 1;
					return true;
				}
				return false;
			});
		}
		if (removed) {
			renderDetails();
		}
	};

	window.popSuppAgentOLD = function () {
		window.popSuppAgent();
	};

	window.popSuppAgent = function () {
		var size = popupSize("2", "PartySelectionCommon.asp", "500", "750");
		var unit = trim(field("hForUnit") && field("hForUnit").value) + ":O";
		openDialog("../../Common/" + size.program + "?orgID=" + encodeURIComponent(unit), window.OutData, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (result) {
			continueSelectionDialog(result, window.OutData, size, processEntrySupplierSelection);
		});
	};

	window.ClearTable = function () {
		clearTableRows();
	};

	window.ClearTableItem = function () {
		clearTableRows();
	};

	window.AddDetails = function () {
		var root = window.OutData && window.OutData.documentElement;
		var itemDescription = trim(field("txtItemDesc") && field("txtItemDesc").value);
		var otherDescription = trim(field("txtDesc") && field("txtDesc").value);
		var description = "";
		var uom = field("selUOM");
		var node;
		if (!root) {
			return;
		}
		if (!otherDescription && !itemDescription) {
			alert("Enter Description or Select Item...!");
			focusField("txtDesc", true);
			return;
		}
		if (!trim(field("txtQty") && field("txtQty").value)) {
			alert("Enter Quantity");
			focusField("txtQty", true);
			return;
		}
		if (uom && !uom.disabled && (uom.value === "select" || !trim(uom.value))) {
			alert("Select UOM");
			focusField("selUOM");
			return;
		}
		if (itemDescription) {
			description = upper(itemDescription);
		}
		if (otherDescription) {
			description = description ? description + " - " + upper(otherDescription) : upper(otherDescription);
		}
		if (hasDetailWithDescription(root, description)) {
			alert("Description Already entered");
			focusField("txtDesc", true);
			return;
		}
		node = window.OutData.createElement("DETAILS");
		setAttr(node, "DESC", description);
		setAttr(node, "QTY", field("txtQty") && field("txtQty").value);
		setAttr(node, "ITEMCODE", field("hItem") && field("hItem").value);
		setAttr(node, "CLASSCODE", field("hClass") && field("hClass").value);
		setAttr(node, "ITEMDESC", upper(itemDescription));
		setAttr(node, "OTHERDESC", upper(otherDescription));
		setAttr(node, "UOM", upper(uom && uom.value));
		setAttr(node, "VALUE", field("txtValue") && field("txtValue").value);
		setAttr(node, "FORMJJ", field("ChkFormJJ") && field("ChkFormJJ").checked ? "Y" : "N");
		setAttr(node, "REASON", trim(field("txtReason") && field("txtReason").value));
		root.appendChild(node);
		renderDetails();
		resetEntryFields();
	};

	window.CheckSubmit = function () {
		var root = window.OutData && window.OutData.documentElement;
		var nodes = detailsNodes(root);
		var header;
		var xhr;
		if (!trim(field("txtRefName") && field("txtRefName").value)) {
			alert("Select Supplier");
			return;
		}
		if (trim(field("txtRemarks") && field("txtRemarks").value).length > 200) {
			alert("Remarks should be less than 200 characters");
			focusField("txtRemarks", true);
			return;
		}
		if (!nodes.length) {
			alert("Enter Details");
			return;
		}
		nodes.forEach(function (node, index) {
			var checkbox = field("chkbox" + (index + 1));
			setAttr(node, "FORMJJ", checkbox && checkbox.checked ? "Y" : "N");
		});
		childElements(root, "HEADER").forEach(function (node) {
			root.removeChild(node);
		});
		header = window.OutData.createElement("HEADER");
		setAttr(header, "FORUNIT", field("hForUnit") && field("hForUnit").value);
		setAttr(header, "ITEMTYPE", "");
		setAttr(header, "SUPPAGENT", supplierCode || field("hPartyCode") && field("hPartyCode").value);
		setAttr(header, "REMARKS", trim(field("txtRemarks") && field("txtRemarks").value));
		setAttr(header, "Transport", trim(field("txtTransport") && field("txtTransport").value));
		setAttr(header, "TakenBy", trim(field("txtTakenBy") && field("txtTakenBy").value));
		setAttr(header, "DeliveryBy", trim(field("txtDeliveryBy") && field("txtDeliveryBy").value));
		root.appendChild(header);

		xhr = new XMLHttpRequest();
		xhr.open("POST", "GatePassServiceUpdate.asp?GatePassNo=" + encodeURIComponent(field("hGatePassNo") && field("hGatePassNo").value), true);
		xhr.setRequestHeader("Content-Type", "text/xml; charset=UTF-8");
		xhr.onreadystatechange = function () {
			if (xhr.readyState !== 4) {
				return;
			}
			if (xhr.responseText === "") {
				if (confirm("Gate Pass for Service has been created. Do you want to create another one?")) {
					window.location.href = "../../Inventory/Transaction/GATEPASSSELECTION.ASP?SelSent=Y&InvoiceType=V";
				} else {
					window.location.href = "../../Inventory/welcome_Inventory.asp";
				}
			} else {
				alert(xhr.responseText);
			}
		};
		xhr.send(getXmlText(window.OutData));
	};

	window.SelectItem = function () {
		var size = popupSize("1", "ItemSelectRelPartyCommon.asp", "500", "850");
		var unit = trim(field("hForUnit") && field("hForUnit").value);
		var url = "../../Common/" + size.program +
			"?orgID=" + encodeURIComponent(unit) +
			"&Stock=N&iType=&hSelectMode=R&Flag=1&hDispButt=Y";
		openDialog(url, window.Data, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (result) {
			continueSelectionDialog(result, window.Data, size, processItemSelection);
		});
	};

	window.Init = function () {
		var sourceRoot = window.newData && window.newData.documentElement;
		var targetRoot = window.OutData && window.OutData.documentElement;
		var targetDoc = xmlDocument(window.OutData);
		if (!sourceRoot || !targetRoot || !targetDoc) {
			return;
		}
		while (targetRoot.firstChild) {
			targetRoot.removeChild(targetRoot.firstChild);
		}
		childElements(sourceRoot).forEach(function (sourceNode) {
			var node = targetDoc.importNode ? targetDoc.importNode(sourceNode, true) : sourceNode.cloneNode(true);
			targetRoot.appendChild(node);
			if (upper(node.nodeName) === "HEADER") {
				setValue("txtRemarks", getAttr(node, "REMARKS"));
				setValue("txtTransport", getAttr(node, "Transport"));
				setValue("txtDeliveryBy", getAttr(node, "TakenBy"));
				setValue("txtTakenBy", getAttr(node, "DeliveryBy"));
			} else if (upper(node.nodeName) === "SUPPLIER") {
				supplierCode = getAttr(node, "SuppCode") || getAttr(node, "PartyCode");
				setValue("txtRefName", getAttr(node, "SuppName"));
			}
		});
		supplierCode = supplierCode || trim(field("hPartyCode") && field("hPartyCode").value);
		childElements(targetRoot, "HEADER").forEach(function (node) {
			setAttr(node, "SUPPAGENT", supplierCode);
		});
		renderDetails();
	};

	window.dispitem = renderDetails;
	window.DispItem = renderDetails;
})(window, document);
