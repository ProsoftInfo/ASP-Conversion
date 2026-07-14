(function (window, document) {
	"use strict";

	var rowIndex = 0;
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
		return frm && (frm.elements[name] || frm.elements[name.toLowerCase()] || frm.elements[name.toUpperCase()]);
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

	function childElements(node) {
		return Array.prototype.filter.call(node && node.childNodes || [], function (child) {
			return child.nodeType === 1;
		});
	}

	function getAttr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function focusField(name, selectText) {
		var item = field(name);
		if (item) {
			if (selectText && item.select) {
				item.select();
			} else if (item.focus) {
				item.focus();
			}
		}
	}

	function openDialog(url, args, features, callback) {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
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

	function hasDetailWithDescription(root, description) {
		var wanted = trim(description).replace(/'/g, " ");
		return childElements(root).some(function (node) {
			return upper(node.nodeName) === "DETAILS" && trim(getAttr(node, "DESC")) === wanted;
		});
	}

	function detailsNodes(root) {
		return childElements(root).filter(function (node) {
			return upper(node.nodeName) === "DETAILS";
		});
	}

	function appendCell(row, text, className, align) {
		var cell = row.insertCell();
		cell.className = className || "ExcelDisplayCell";
		cell.align = align || "left";
		cell.innerHTML = text == null ? "" : String(text);
		return cell;
	}

	function tableElement(id) {
		return document.getElementById(id) || window[id];
	}

	function processSupplierSelection(result) {
		var root = xmlRoot(result, window.OutData);
		if (!root || !root.hasChildNodes()) {
			return;
		}
		childElements(root).forEach(function (node) {
			if (upper(node.nodeName) === "SUPPLIER") {
				supplierCode = getAttr(node, "SuppCode");
				setValue("txtRefName", getAttr(node, "SuppShortCode"));
			}
		});
	}

	function continueSupplierDialog(result) {
		var action = upper(getAttr(result, "Action"));
		var passQuery = getAttr(result, "PassQuery");
		if (action && action !== "DONE" && action !== "CLOSE") {
			openDialog("SupplierSelect.asp?" + passQuery, window.OutData, "status:no", continueSupplierDialog);
			return;
		}
		if (action !== "CLOSE") {
			processSupplierSelection(result);
		}
	}

	window.popSuppAgent = function () {
		var unit = trim(field("hUnitID") && field("hUnitID").value) + ":O";
		openDialog("SupplierSelect.asp?Unit=" + encodeURIComponent(unit) + "&hSelectMode=S&Flag=1", window.OutData, "status:no", continueSupplierDialog);
	};

	window.ClearTable = function () {
		var table = tableElement("tblLot") || tableElement("tblDetails");
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
		rowIndex = 0;
	};

	window.AddDetails = function () {
		var frm = form();
		var root = window.OutData && window.OutData.documentElement;
		var table = tableElement("tblDetails");
		var itemDescription = trim(field("txtItemDesc") && field("txtItemDesc").value);
		var otherDescription = trim(field("txtDesc") && field("txtDesc").value);
		var description = "";
		var uom = field("selUOM");
		var applyFormJj;
		var node;
		var row;
		var checkbox;

		if (!root || !table || !frm) {
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
		applyFormJj = field("ChkFormJJ") && field("ChkFormJJ").checked ? "Y" : "N";

		if (hasDetailWithDescription(root, description)) {
			alert("Description Already entered");
			focusField("txtDesc", true);
			return;
		}

		node = window.OutData.createElement("DETAILS");
		node.setAttribute("DESC", description);
		node.setAttribute("QTY", field("txtQty").value);
		node.setAttribute("ITEMCODE", field("hItem").value);
		node.setAttribute("CLASSCODE", field("hClass").value);
		node.setAttribute("ITEMDESC", upper(itemDescription));
		node.setAttribute("OTHERDESC", upper(otherDescription));
		node.setAttribute("UOM", upper(uom && uom.value));
		node.setAttribute("VALUE", field("txtValue").value);
		node.setAttribute("FORMJJ", applyFormJj);
		node.setAttribute("REASON", trim(field("txtReason") && field("txtReason").value));
		root.appendChild(node);

		row = table.insertRow(rowIndex + 1);
		appendCell(row, rowIndex + 1, "ExcelSerial", "center");
		appendCell(row, description, "ExcelDisplayCell", "left");
		appendCell(row, field("txtQty").value, "ExcelDisplayCell", "right");
		appendCell(row, uom && uom.value, "ExcelDisplayCell", "right");
		appendCell(row, field("txtValue").value, "ExcelDisplayCell", "right");
		checkbox = document.createElement("input");
		checkbox.type = "checkbox";
		checkbox.value = "Y";
		checkbox.name = "chkbox" + (rowIndex + 1);
		checkbox.className = "FormElem";
		checkbox.checked = applyFormJj === "Y";
		appendCell(row, "", "ExcelDisplayCell", "right").appendChild(checkbox);
		appendCell(row, trim(field("txtReason") && field("txtReason").value), "ExcelDisplayCell", "left");

		rowIndex += 1;
		setValue("txtDesc", "");
		setValue("txtQty", "");
		setValue("txtItemDesc", "");
		setValue("hItem", "");
		setValue("hClass", "");
		if (uom) {
			uom.value = "select";
		}
		setValue("txtValue", "");
		if (field("ChkFormJJ")) {
			field("ChkFormJJ").checked = false;
		}
	};

	window.CheckSubmit = function () {
		var root;
		var nodes;
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

		root = window.OutData && window.OutData.documentElement;
		nodes = detailsNodes(root);
		if (!nodes.length) {
			alert("Enter Details");
			return;
		}

		nodes.forEach(function (node, index) {
			var checkbox = field("chkbox" + (index + 1));
			node.setAttribute("FORMJJ", checkbox && checkbox.checked ? "Y" : "N");
		});

		header = window.OutData.createElement("HEADER");
		header.setAttribute("FORUNIT", field("hUnitID").value);
		header.setAttribute("ITEMTYPE", "");
		header.setAttribute("SUPPAGENT", supplierCode);
		header.setAttribute("REMARKS", trim(field("txtRemarks") && field("txtRemarks").value));
		header.setAttribute("Transport", trim(field("txtTransport") && field("txtTransport").value));
		header.setAttribute("TakenBy", trim(field("txtTakenBy") && field("txtTakenBy").value));
		header.setAttribute("DeliveryBy", trim(field("txtDeliveryBy") && field("txtDeliveryBy").value));
		root.appendChild(header);

		xhr = new XMLHttpRequest();
		xhr.open("POST", "GatePassServiceInsert.asp", true);
		xhr.setRequestHeader("Content-Type", "text/xml; charset=UTF-8");
		xhr.onreadystatechange = function () {
			if (xhr.readyState !== 4) {
				return;
			}
			if (xhr.responseText === "") {
				if (confirm("Gate Pass for Service has been created. Do you want to create another one?")) {
					window.location.href = "GatePassServiceEntry.asp";
				} else {
					window.location.href = "GATEPASSSELECTION.ASP";
				}
			} else {
				alert(xhr.responseText);
			}
		};
		xhr.send(getXmlText(window.OutData));
	};

	window.SelectItem = function () {
		var unit = trim(field("hUnitID") && field("hUnitID").value);
		var sizeText = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup("1") : "";
		var size = sizeText.split(":");
		var programName = size[0] || "ItemSelectRelPartyCommon.asp";
		var popupHeight = size[1] || "500";
		var popupWidth = size[2] || "850";
		var url = "../../Common/" + programName +
			"?orgID=" + encodeURIComponent(unit) +
			"&sIType=&Stock=&hSelectMode=S&Flag=&hDispButt=&hDispItem=&CallFrom=";

		openDialog(url, window.Data, "dialogHeight:" + popupHeight + "px;dialogWidth:" + popupWidth + "px;Status:No", function (result) {
			var root = xmlRoot(result, window.Data);
			if (!root || !root.hasChildNodes()) {
				return;
			}
			childElements(root).some(function (node) {
				setValue("txtItemDesc", trim(getAttr(node, "ItemName")));
				setValue("hItem", getAttr(node, "ItemCode"));
				setValue("hClass", getAttr(node, "ClassCode"));
				if (field("selUOM")) {
					field("selUOM").value = trim(getAttr(node, "StoresUoM"));
					field("selUOM").disabled = true;
				}
				return true;
			});
		});
	};
})(window, document);
