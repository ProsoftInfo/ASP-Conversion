(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || {};
	}

	function field(name) {
		var frm = form();
		var elements;
		var target;
		var index;
		if (!frm.elements) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		elements = frm.elements;
		target = String(name).toLowerCase();
		for (index = 0; index < elements.length; index += 1) {
			if (String(elements[index].name || "").toLowerCase() === target) {
				return elements[index];
			}
		}
		return null;
	}

	function fields(name) {
		var item = field(name);
		if (item && item.length != null && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		return item ? [item] : [];
	}

	function valueOf(name, fallback) {
		var item = field(name);
		return item && item.value != null ? item.value : fallback || "";
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

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function dialogRoot() {
		return xmlRoot(window.dialogArguments) || (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.dialogArgumentsRoot()) || xmlRoot("OutData") || xmlRoot("NewData");
	}

	function childElements(node, name) {
		var wanted = name && String(name).toLowerCase();
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttr(node, name, value) {
		if (node) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function createInstrument(root) {
		var doc = root && root.ownerDocument || document.implementation.createDocument("", "", null);
		return doc.createElement("BankInstrumentDet");
	}

	function getDateControl(name) {
		var control = field(name) || byId(name);
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate) {
			return window.ITMSModernCompat.toDisplayDate(control.value);
		}
		return control.value || "";
	}

	function setDateControl(name, value) {
		var control = field(name) || byId(name);
		if (!control || value == null || value === "") {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate) {
			control.value = window.ITMSModernCompat.toIsoDate(value);
		} else {
			control.value = value;
		}
	}

	function todayDisplay() {
		var date = new Date();
		var day = String(date.getDate()).padStart(2, "0");
		var month = String(date.getMonth() + 1).padStart(2, "0");
		return day + "/" + month + "/" + date.getFullYear();
	}

	function selectedInstrumentTypeCode() {
		var selected = fields("optInsType").filter(function (item) {
			return item.checked;
		})[0];
		return selected ? selected.value : "C";
	}

	function selectedInstrumentTypeLabel() {
		var code = selectedInstrumentTypeCode();
		if (code === "C") {
			return "Cheque";
		}
		if (code === "D") {
			return "Demand Draft";
		}
		if (code === "B") {
			return "Bankers Cheque";
		}
		if (code === "T") {
			return "RTGS";
		}
		return "Cash Deposited";
	}

	function setInstrumentType(value) {
		var label = trim(value);
		var code = label === "Cheque" ? "C" : label === "Demand Draft" ? "D" : label === "Bankers Cheque" ? "B" : label === "RTGS" ? "T" : label === "Cash Withdrawn" || label === "Cash Deposited" ? "W" : value;
		fields("optInsType").forEach(function (item) {
			item.checked = item.value === code;
		});
		setValue("hInsType", code);
	}

	function optionText(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function selectedInstrumentValue() {
		var select = field("SelInsNo");
		return select ? select.value : "";
	}

	function displayInstrumentNo(value, option) {
		if (trim(option) === "Y") {
			var select = field("SelInsNo");
			var parts = String(value || "").split("-");
			if (select) {
				if (trim(value)) {
					select.value = value;
					if (select.value !== value) {
						select.add(new Option(parts[2] || value, value));
						select.value = value;
					}
				}
				if (select.selectedIndex >= 0 && parts[2]) {
					select.options[select.selectedIndex].text = parts[2];
				}
			}
			setValue("txtInsNo", "");
			return;
		}
		setValue("txtInsNo", value);
	}

	function selectedOrEnteredInstrumentNo() {
		var text = field("txtInsNo");
		if (text && text.disabled) {
			return selectedInstrumentValue();
		}
		return valueOf("txtInsNo");
	}

	function selectedOptionFlag() {
		var text = field("txtInsNo");
		return text && text.disabled ? "Y" : "";
	}

	function resetEntryFields() {
		setValue("txtInsNo", "");
		setDateControl("ctlDate", todayDisplay());
		setValue("txtPayableAt", "");
		setValue("txtDrawnOn", "");
		setValue("txtAmount", "");
		setValue("hExists", "");
		setValue("hEditNo", "");
	}

	function clearTable() {
		var table = byId("InsTab");
		while (table && table.rows.length) {
			table.deleteRow(0);
		}
	}

	function appendCell(row, text, className, align, width) {
		var cell = row.insertCell(row.cells.length);
		cell.className = className || "ExcelDisplayCell";
		cell.align = align || "left";
		if (width) {
			cell.width = width;
		}
		cell.textContent = text == null ? "" : String(text);
		return cell;
	}

	function readonlyInput(value, name, size, align) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.value = value == null ? "" : String(value);
		input.size = size;
		input.className = "FormelemRead";
		input.style.textAlign = align || "left";
		input.readOnly = true;
		return input;
	}

	function displayTable() {
		var root = dialogRoot();
		var table = byId("InsTab");
		var next = 1;
		clearTable();
		if (!table || !root) {
			return;
		}
		var header = table.insertRow(table.rows.length);
		["Sl.No", "", "Instrument No", "Instrument Date", "Payble At", "Drawn On", "Instrument Type", "Instrument Amount"].forEach(function (title, index) {
			var cell = appendCell(header, title, "ExcelHeaderCell", "center");
			if (index === 1) {
				var link = document.createElement("a");
				var img = document.createElement("img");
				link.href = "#";
				link.onclick = function () {
					window.DeleteItems();
					return false;
				};
				img.border = "0";
				img.src = "../../assets/images/iTMS%20Icons/DeleteIcon.gif";
				img.width = 15;
				img.height = 15;
				link.appendChild(img);
				cell.textContent = "";
				cell.appendChild(link);
			}
		});
		childElements(root, "BankInstrumentDet").forEach(function (node) {
			var slNo = attr(node, "SlNo") || next;
			var rawInsNo = attr(node, "InsNo");
			var option = attr(node, "Option");
			var displayNo = rawInsNo;
			var row = table.insertRow(table.rows.length);
			var checkbox;
			if (trim(option) === "Y" && trim(rawInsNo)) {
				displayNo = rawInsNo.split("-")[2] || rawInsNo;
			}
			setAttr(node, "SlNo", slNo);
			appendCell(row, slNo, "ExcelSerial", "center", 3);
			var checkCell = appendCell(row, "", "ExcelDisplayCell", "center", 3);
			checkbox = document.createElement("input");
			checkbox.type = "checkbox";
			checkbox.name = "ChkInsNo" + slNo;
			checkbox.value = slNo + ":" + displayNo;
			checkbox.className = "Formelem";
			checkbox.onclick = function () {
				window.DispVal(slNo, rawInsNo, attr(node, "InsType"), attr(node, "InsDate"), attr(node, "PayAt"), attr(node, "DrawnOn"), option, attr(node, "InsAmt"));
			};
			checkCell.appendChild(checkbox);
			[
				[displayNo, "InsNo", 7, "left"],
				[attr(node, "InsDate"), "InsDate", 12, "center"],
				[attr(node, "PayAt"), "PayAt", 15, "left"],
				[attr(node, "DrawnOn"), "DrawnOn", 15, "left"],
				[attr(node, "InsType"), "OptType", 22, "left"],
				[attr(node, "InsAmt"), "InsAmt", 12, "right"]
			].forEach(function (part) {
				var cell = appendCell(row, "", "ExcelDisplayCell", part[3], part[2]);
				cell.appendChild(readonlyInput(part[0], part[1], part[2], part[3]));
			});
			setInstrumentType(attr(node, "InsType"));
			next = Number(slNo) + 1;
		});
		setValue("hCtr", String(next));
		resetEntryFields();
	}

	function validateInstrument() {
		var usingSelect = field("txtInsNo") && field("txtInsNo").disabled;
		if (usingSelect) {
			if (trim(selectedInstrumentValue()) === "S" || trim(selectedInstrumentValue()) === "0" || trim(selectedInstrumentValue()) === "") {
				alert("Select Instrument No");
				return false;
			}
		} else if (trim(valueOf("txtInsNo")) === "") {
			alert("Enter Instrument No");
			return false;
		}
		if (trim(valueOf("txtAmount")) === "") {
			alert("Enter Instrument Amount");
			return false;
		}
		if (toNumber(valueOf("txtAmount")) <= 0) {
			alert("Instrument Amount Should be Greater than 0");
			return false;
		}
		if (!usingSelect && isNaN(parseFloat(valueOf("txtAmount")))) {
			setValue("txtAmount", "");
			alert("Enter Numeric Values");
			return false;
		}
		return true;
	}

	function fillInstrumentNode(node) {
		setAttr(node, "InsNo", selectedOrEnteredInstrumentNo());
		setAttr(node, "InsType", selectedInstrumentTypeLabel());
		setAttr(node, "InsDate", getDateControl("ctlDate"));
		setAttr(node, "PayAt", valueOf("txtPayableAt"));
		setAttr(node, "DrawnOn", valueOf("txtDrawnOn"));
		setAttr(node, "InsAmt", valueOf("txtAmount"));
		setAttr(node, "Option", selectedOptionFlag());
		setAttr(node, "Action", trim(valueOf("hTransNo")) !== "0" ? valueOf("SelAct", "0") : "0");
	}

	window.OptFun = function (item) {
		var value = item && item.value || selectedInstrumentTypeCode();
		var flag = trim(valueOf("hFlag")) === "True";
		var select = field("SelInsNo");
		var text = field("txtInsNo");
		setValue("hInsType", value);
		if (value === "C") {
			if (flag && select && text) {
				select.disabled = false;
				text.disabled = true;
				text.value = "";
			} else if (text) {
				text.disabled = false;
			}
			return;
		}
		if (flag && select && text) {
			select.disabled = true;
			select.value = "0";
			text.disabled = false;
		}
		setValue("txtInsNo", value === "T" ? "0" : "");
	};

	window.InsType = function () {
		if (trim(valueOf("hFlag")) === "True") {
			window.OptFun(fields("optInsType").filter(function (item) {
				return item.checked;
			})[0]);
		}
	};

	window.SelAction = function () {
		var select = field("SelInsNo");
		if (trim(valueOf("SelAct")) !== "0") {
			if (select) {
				select.disabled = false;
			}
			return;
		}
		var parts = valueOf("hUsInsNo").split("-");
		if (select) {
			select.value = valueOf("hUsInsNo");
			if (select.selectedIndex >= 0 && parts[2]) {
				select.options[select.selectedIndex].text = parts[2];
			}
			select.disabled = true;
		}
	};

	window.Init = function () {
		var initialDate = valueOf("hDate") || todayDisplay();
		ensureCompat();
		setDateControl("ctlDate", initialDate);
		if (childElements(dialogRoot(), "BankInstrumentDet").length) {
			displayTable();
		}
	};

	window.DisplayDet = function () {
		var first = childElements(dialogRoot(), "BankInstrumentDet")[0];
		if (!first) {
			if (trim(valueOf("hTransNo")) !== "0" && trim(valueOf("hUsInsNo")) && trim(valueOf("hFlag")) === "True") {
				displayInstrumentNo(valueOf("hUsInsNo"), "Y");
			}
			return;
		}
		setInstrumentType(attr(first, "InsType"));
		displayInstrumentNo(attr(first, "InsNo"), attr(first, "Option"));
		setDateControl("ctlDate", attr(first, "InsDate"));
		setValue("txtPayableAt", attr(first, "PayAt"));
		setValue("txtDrawnOn", attr(first, "DrawnOn"));
		setValue("txtAmount", attr(first, "InsAmt"));
	};

	window.AddFun = function () {
		var root = dialogRoot();
		var exists = trim(valueOf("hExists")) === "Y";
		var editNo = valueOf("hEditNo");
		var node;
		if (!root || !validateInstrument()) {
			return false;
		}
		if (exists) {
			childElements(root, "BankInstrumentDet").some(function (item) {
				if (trim(attr(item, "SlNo")) === trim(editNo)) {
					fillInstrumentNode(item);
					return true;
				}
				return false;
			});
		} else {
			node = createInstrument(root);
			setAttr(node, "SlNo", valueOf("hCtr", "1") || "1");
			fillInstrumentNode(node);
			root.appendChild(node);
		}
		if (field("B2")) {
			field("B2").disabled = false;
		}
		displayTable();
		return false;
	};

	window.ClearTable = clearTable;

	window.DeleteItems = function () {
		var root = dialogRoot();
		var checked = Array.prototype.slice.call(document.querySelectorAll("input[name^='ChkInsNo']:checked"));
		var remove = {};
		var counter = 1;
		checked.forEach(function (item) {
			remove[String(item.value).split(":")[0]] = true;
		});
		childElements(root, "BankInstrumentDet").forEach(function (node) {
			if (remove[attr(node, "SlNo")]) {
				root.removeChild(node);
			}
		});
		childElements(root, "BankInstrumentDet").forEach(function (node) {
			setAttr(node, "SlNo", counter);
			counter += 1;
		});
		setValue("hCtr", counter);
		displayTable();
		return false;
	};

	window.DispVal = function (slNo, insNo, optType, insDate, payAt, drawnOn, option, insAmt) {
		setInstrumentType(optType);
		displayInstrumentNo(insNo, option);
		setDateControl("ctlDate", insDate);
		setValue("txtPayableAt", payAt);
		setValue("txtDrawnOn", drawnOn);
		setValue("txtAmount", insAmt);
		setValue("hExists", "Y");
		setValue("hEditNo", slNo);
	};

	window.DisplayTable = displayTable;

	window.CheckSubmit = function () {
		var root = dialogRoot();
		if (window.ITMSModalReturnCompat) {
			window.ITMSModalReturnCompat.returnAndClose(root);
		} else {
			window.returnValue = root;
			window.close();
		}
		return false;
	};

	window.window_onunload = function () {
		var root = dialogRoot();
		if (window.ITMSModalReturnCompat) {
			window.ITMSModalReturnCompat.returnValue(root);
		} else {
			window.returnValue = root;
		}
	};

	if (window.ITMSModalReturnCompat) {
		window.ITMSModalReturnCompat.install(function () {
			return dialogRoot();
		});
	}
}(window, document));
