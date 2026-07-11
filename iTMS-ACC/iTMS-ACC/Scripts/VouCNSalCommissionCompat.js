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
		var target = String(name).toLowerCase();
		var index;
		if (!frm || !frm.elements) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		for (index = 0; index < frm.elements.length; index += 1) {
			if (String(frm.elements[index].name || "").toLowerCase() === target) {
				return frm.elements[index];
			}
		}
		return document.getElementsByName(name)[0] || document.getElementById(name) || null;
	}

	function fields(name) {
		return Array.prototype.slice.call(document.getElementsByName(name) || []);
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

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function textOf(id) {
		var item = byId(id);
		return item ? trim(item.textContent || item.innerText || item.innerHTML || "") : "";
	}

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.innerHTML = value == null ? "" : String(value);
		}
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 && select.options[select.selectedIndex] ? select.options[select.selectedIndex].text : "";
	}

	function selectedRadioValue(name, fallback) {
		var radios = fields(name).filter(function (item) {
			return String(item.type || "").toLowerCase() === "radio";
		});
		var selected = radios.filter(function (item) {
			return item.checked;
		})[0];
		return selected ? selected.value : fallback || "";
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(nameOrObject) {
		if (typeof nameOrObject !== "string") {
			return nameOrObject;
		}
		ensureCompat();
		return window[nameOrObject] || document[nameOrObject] || byId(nameOrObject) || null;
	}

	function xmlDocument(nameOrObject) {
		var object = xmlObject(nameOrObject);
		return object && object.XMLDocument || object && object._doc || object && object.nodeType === 9 && object || null;
	}

	function xmlRoot(nameOrObject) {
		var object = xmlObject(nameOrObject);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function serializeXml(nameOrObject) {
		var doc = xmlDocument(nameOrObject);
		var root = xmlRoot(nameOrObject);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function createNode(xmlName, nodeName) {
		var doc = xmlDocument(xmlName);
		return doc && doc.createElement ? doc.createElement(nodeName) : document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function childElements(node, nodeName) {
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function attr(node, nameOrIndex) {
		var item;
		if (!node || !node.attributes) {
			return "";
		}
		if (typeof nameOrIndex === "number") {
			item = node.attributes.item(nameOrIndex);
			return item ? item.nodeValue : "";
		}
		return node.getAttribute(nameOrIndex) || "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function setAttrByIndex(node, index, value, fallbackName) {
		var item = node && node.attributes && node.attributes.item(index);
		if (item) {
			item.nodeValue = value == null ? "" : String(value);
		} else if (fallbackName) {
			setAttr(node, fallbackName, value);
		}
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var result = [];
		var i;
		if (!context) {
			return result;
		}
		if (typeof context.selectNodes === "function") {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return result;
		}
		found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (i = 0; i < found.snapshotLength; i += 1) {
			result.push(found.snapshotItem(i));
		}
		return result;
	}

	function clearDocument(xmlName) {
		var doc = xmlDocument(xmlName);
		while (doc && doc.firstChild) {
			doc.removeChild(doc.firstChild);
		}
		return doc;
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
		} catch (ignore) {}
		xhr.send(body || null);
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function rootFromDialog(value) {
		return xmlRoot(value) || value && value.nodeType === 9 && value.documentElement || value && value.nodeType === 1 && value || null;
	}

	function retFromGlRoot(root) {
		var entry = childElements(root, "Entry")[0] || childElements(root)[0];
		if (!entry) {
			return "";
		}
		return [0, 1, 2, 3, 4, 5, 6, 7].map(function (index) {
			return attr(entry, "RetField" + index);
		}).join(":");
	}

	function runGlDialog(url, callback) {
		openDialog(url, "", "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (value) {
			var root = rootFromDialog(value);
			var action = trim(attr(root, "Action")).toUpperCase();
			var passQuery = trim(attr(root, "PassQuery"));
			var text = trim(value);
			if (root && action === "CLOSE") {
				return;
			}
			if (root && action && action !== "DONE" && passQuery) {
				runGlDialog("GLHeadSelection.asp?" + passQuery, callback);
				return;
			}
			if (!root && text && text.split(":").length <= 1) {
				runGlDialog("GLHeadSelection.asp?" + text, callback);
				return;
			}
			callback(root ? retFromGlRoot(root) : text);
		});
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
		return window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate ? window.ITMSModernCompat.toDisplayDate(control.value) : control.value || "";
	}

	function setDateControl(name, value) {
		var control = field(name) || byId(name);
		if (!control || trim(value) === "") {
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

	function currentYearMonth() {
		var parts = String(getDateControl("ctlDate") || "").split("/");
		return toNumber((parts[2] || "") + (parts[1] || ""));
	}

	function validateAccountHead() {
		if (trim(valueOf("SelAccountHd")) === "0") {
			alert("Select Account Head");
			return false;
		}
		if (trim(textOf("spAccHead")) === "") {
			alert("Select Account Head Value");
			return false;
		}
		return true;
	}

	function updateApprovalAttributes(root) {
		var approval = selectedRadioValue("optApprove", "Y") === "Y" ? "Y" : "N";
		var approver = field("selUserId");
		var invoices = selectNodes(root, "//SaleInvoice");
		if (approval === "Y" && approver && approver.selectedIndex === 0) {
			alert("Select Approver ");
			approver.focus();
			return false;
		}
		invoices.forEach(function (node) {
			setAttr(node, "Approval", approval);
			setAttr(node, "Approver", valueOf("selUserId"));
			setAttr(node, "SalTrNo", valueOf("hdTransNo"));
		});
		return true;
	}

	function appendNarration(root) {
		var narration = createNode("TaxData", "Narration");
		narration.textContent = valueOf("txtNarration");
		root.appendChild(narration);
	}

	function updateCreditNoteXml(root) {
		childElements(root, "Details").forEach(function (node) {
			setAttr(node, "VouDate", getDateControl("ctlDate"));
		});
		appendNarration(root);
		if (!updateApprovalAttributes(root)) {
			return false;
		}
		selectNodes(root, "//TaxDetails").forEach(function (node) {
			setAttrByIndex(node, 0, valueOf("txtTotalInv"), "TaxAmount");
		});
		if (trim(valueOf("SelCrAgain")) === "A") {
			selectNodes(root, "//Entry").forEach(function (node, index) {
				setAttr(node, "Amount", valueOf("txtAmount" + (index + 1)));
			});
			selectNodes(root, "//Tax").forEach(function (node) {
				setAttr(node, "TaxAmount", "0.00");
			});
		}
		if (trim(valueOf("SelAccountHd")) === "G") {
			selectNodes(root, "//AccHead").forEach(function (node) {
				setAttr(node, "No", valueOf("hCrAccHead"));
				setAttr(node, "Name", textOf("spAccHead"));
			});
			selectNodes(root, "//Tax").forEach(function (node) {
				setAttr(node, "AccHead", valueOf("hCrAccHead"));
			});
		}
		return true;
	}

	function finishCreditNoteSave() {
		var xhr = syncPost("XMLSave.asp?Mod=CN&Name=Voucher%20Entry", serializeXml("TaxData"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return false;
		}
		if (trim(valueOf("hInvCallFrom")) === "SI") {
			form().action = "VouCNOthInvGenerate.asp";
		} else if (trim(valueOf("hInvCallFrom")) === "SR") {
			form().action = "VouCNSalRetAdj.asp";
		}
		form().submit();
		return true;
	}

	function appendGjEntry(doc, root, no, crdr, payTo, amount, unitNo, unitName) {
		var entry = doc.createElement("Entry");
		setAttr(entry, "No", no);
		setAttr(entry, "CRDR", crdr);
		setAttr(entry, "Payto", payTo);
		setAttr(entry, "Amount", amount);
		setAttr(entry, "AccUnit", unitNo);
		setAttr(entry, "AccName", unitName);
		setAttr(entry, "TdsAmount", "0.00");
		setAttr(entry, "TDSElgi", "0");
		setAttr(entry, "TdsPercentage", "0");
		setAttr(entry, "PayRecAmount", "0");
		root.appendChild(entry);
		return entry;
	}

	function appendNarrationToEntry(doc, entry) {
		var narration = doc.createElement("Narration");
		narration.textContent = valueOf("txtNarration");
		entry.appendChild(narration);
	}

	function agentPartyNo() {
		return [
			valueOf("hagentType") || valueOf("hAgentType"),
			valueOf("hagentsubType") || valueOf("hAgentSubType"),
			(valueOf("hagentType") || valueOf("hAgentType")) + "- Commission Agent",
			valueOf("hagentCode") || valueOf("hAgentCode")
		].join("?");
	}

	function buildGjVoucher(sourceRoot) {
		var doc = clearDocument("GJVoucher");
		var root = doc.createElement("voucher");
		var unitNo = "";
		var unitName = "";
		var bookAccHead = "";
		var firstAdded = false;
		var count = 0;
		doc.appendChild(root);
		childElements(sourceRoot).forEach(function (header) {
			if (header.nodeName === "Header") {
				childElements(header).forEach(function (node) {
					if (node.nodeName === "Organization") {
						unitNo = attr(node, "OrgId");
						unitName = node.textContent || "";
						setAttr(root, "UnitNo", unitNo);
						setAttr(root, "UnitName", unitName);
						setAttr(root, "BookNo", valueOf("hBookCode"));
						setAttr(root, "BookName", valueOf("hBookName"));
						setAttr(root, "CRDR", "");
					}
					if (node.nodeName === "Book") {
						bookAccHead = attr(node, "BKAccHead");
					}
				});
			}
			if (header.nodeName === "Details") {
				setAttr(root, "VouDate", getDateControl("ctlDate"));
				setAttr(root, "BookAcchead", bookAccHead);
				setAttr(root, "Approver", valueOf("selUserId"));
				childElements(header, "Entry").forEach(function (sourceEntry) {
					var firstEntry;
					var secondEntry;
					if (!firstAdded) {
						count += 1;
						firstEntry = appendGjEntry(doc, root, count, "C", "0", valueOf("txtTotalInv"), unitNo, unitName);
						childElements(sourceEntry, "AccHead").some(function () {
							var acc = doc.createElement("AccHead");
							setAttr(acc, "No", agentPartyNo());
							setAttr(acc, "Pay", "0");
							setAttr(acc, "Rec", "0");
							setAttr(acc, "Name", valueOf("hagentName") || valueOf("hAgentName"));
							setAttr(acc, "Type", "P");
							setAttr(acc, "Adv", "0");
							firstEntry.appendChild(acc);
							appendNarrationToEntry(doc, firstEntry);
							return true;
						});
						firstAdded = true;
					}
					count += 1;
					secondEntry = appendGjEntry(doc, root, count, "D", "", valueOf("txtTotalInv"), unitNo, unitName);
					childElements(sourceEntry, "AccHead").forEach(function (sourceAcc) {
						var acc = doc.createElement("AccHead");
						setAttr(acc, "No", valueOf("hCrAccHead"));
						setAttr(acc, "CostCenter", attr(sourceAcc, "CostCenter"));
						setAttr(acc, "Analytical", attr(sourceAcc, "Analytical"));
						setAttr(acc, "Name", textOf("spAccHead"));
						setAttr(acc, "Type", attr(sourceAcc, "Type"));
						setAttr(acc, "TransFlag", attr(sourceAcc, "No") !== "0" ? "A" : "W");
						secondEntry.appendChild(acc);
						appendNarrationToEntry(doc, secondEntry);
					});
				});
			}
		});
		return doc;
	}

	function finishGjSave(doc) {
		var xhr;
		syncPost("AccSalCommUpdate.asp", null);
		if (!window.CheckFinDate()) {
			return false;
		}
		xhr = syncPost("XMLSave.asp?Mod=GJ&Name=Voucher%20Entry", new XMLSerializer().serializeToString(doc));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return false;
		}
		form().action = "VouGenerate.asp";
		form().submit();
		return true;
	}

	window.EnbApp = function (item) {
		var approver = field("selUserId");
		if (!approver) {
			return false;
		}
		if (item && item.value === "Y") {
			approver.disabled = false;
		} else {
			approver.selectedIndex = 0;
			approver.disabled = true;
		}
		return false;
	};

	window.SetRetVal = function () {
		var invoice = textOf("InoviceNo");
		var parts = invoice.split("-");
		setValue("txtNarration", "CR Note for " + (parts[0] || "") + " " + (parts[1] || "") + " Sales Return Qty: ");
	};

	window.ResetTax = function () {
		var root = xmlRoot("TaxData");
		var taxRoot = childElements(root, "TaxDetails")[0];
		childElements(taxRoot).forEach(function (node) {
			var catCode = attr(node, 0);
			var taxCode = attr(node, 1);
			setAttrByIndex(node, 5, "0.00", "TaxAmount");
			setValue("txtTaxValue" + catCode + taxCode, "0.00");
		});
		setValue("txtInvValue", "0.00");
		setAttrByIndex(taxRoot, 0, "0.00", "BasicValue");
		setAttrByIndex(taxRoot, 1, "0.00", "Discount");
		setAttrByIndex(taxRoot, 2, "0.00", "ActualValue");
	};

	window.CheckNoSer = function () {
		var passValue;
		var response;
		if (trim(valueOf("hVouCode")) === "04") {
			passValue = [valueOf("selUnitId"), valueOf("hVouCode"), valueOf("hCallFrm"), "D", valueOf("selBook")].join(":");
		} else {
			passValue = [valueOf("hOrgid") || valueOf("hOrgId"), valueOf("hVouCode"), valueOf("hCallFrm"), valueOf("hVouCRDR") || "D", valueOf("hBookCode")].join(":");
		}
		passValue += ":" + getDateControl("ctlDate");
		response = trim(syncGet("NoSeriesCheck.asp?sValue=" + encodeURIComponent(passValue)).responseText);
		if (response === "T") {
			return true;
		}
		alert(response === "F" ? "No Series is Not Defined " : "Error ");
		return false;
	};

	window.AccHead = function (select) {
		if (select && select.selectedIndex > 0 && select.value === "G") {
			window.showGLHead(valueOf("hOrgId") || valueOf("hOrgid"));
		}
		return false;
	};

	window.showGLHead = function (orgId) {
		var bookNo = valueOf("hBookcode") || valueOf("hBookCode");
		var url = "GLHeadSelection.asp?orgId=" + encodeURIComponent(orgId || "") + "&BookId=01&BookNo=" + encodeURIComponent(bookNo) + "&AccHead=" + encodeURIComponent(valueOf("hCrAccHead", "0"));
		runGlDialog(url, function (ret) {
			var parts = String(ret || "").split(":");
			if (parts.length <= 1) {
				return;
			}
			setValue("hCrAccHead", parts[0]);
			setText("spAccHead", parts[5] || parts[3] || "");
		});
		return false;
	};

	window.ReTotalCr = function () {
		var root = xmlRoot("TaxData");
		var taxTotal = 0;
		selectNodes(root, "//Tax").forEach(function (node) {
			var taxCode = attr(node, "TaxCode");
			var catCode = attr(node, "CatCode");
			var item = field("txtTaxValue" + catCode + taxCode);
			if (item) {
				setAttr(node, "TaxAmount", item.value);
				taxTotal += toNumber(item.value);
			}
		});
		setValue("txtInvValue", (toNumber(valueOf("txtTotal")) + taxTotal).toFixed(2));
		setValue("txtTotalInv", valueOf("txtInvValue"));
	};

	window.CheckFinDate = function () {
		var finFrom = toNumber(valueOf("hFinFrm"));
		var finTo = toNumber(valueOf("hFinTo"));
		var current = currentYearMonth();
		if (current < finFrom || current > finTo) {
			alert("Voucher Date Should Be Between 01/04/" + String(finFrom).substring(0, 4) + " To 31/03/" + String(finTo).substring(0, 4));
			return false;
		}
		return true;
	};

	window.SaveXML = function () {
		var root = xmlRoot("TaxData");
		var doc;
		if (!validateAccountHead()) {
			return false;
		}
		if (trim(valueOf("hCallFromVoucher")) === "CR") {
			if (!updateCreditNoteXml(root) || !window.CheckFinDate()) {
				return false;
			}
			return finishCreditNoteSave();
		}
		doc = buildGjVoucher(root);
		return finishGjSave(doc);
	};

	window.InitVouCNSalCommission = function () {
		ensureCompat();
		setText("sPartyName", valueOf("hAgentName") || valueOf("hagentName"));
		setValue("txtNarration", valueOf("hTxtnarr"));
		setDateControl("ctlDate", getDateControl("ctlDate"));
	};
}(window, document));
