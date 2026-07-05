(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		var lower;
		var index;
		if (!frm || !frm.elements || !name) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		lower = String(name).toLowerCase();
		for (index = 0; index < frm.elements.length; index += 1) {
			if (String(frm.elements[index].name || "").toLowerCase() === lower) {
				return frm.elements[index];
			}
		}
		return null;
	}

	function valueOf(name, fallback) {
		var item = field(name);
		return item ? item.value : fallback || "";
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

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.upgradeXmlIslands) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
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

	function loadXml(name, xmlText) {
		var object = xmlObject(name);
		var doc;
		if (object && typeof object.loadXML === "function") {
			object.loadXML(xmlText || "<Root/>");
			return object;
		}
		doc = new DOMParser().parseFromString(xmlText || "<Root/>", "text/xml");
		if (object) {
			object._doc = doc;
		}
		return object || doc;
	}

	function loadResponseXml(islandName, xhr) {
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml(islandName, serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml(islandName, xhr.responseText);
		}
	}

	function childElements(node, nodeName) {
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName || "").toLowerCase() === wanted);
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

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function selectedDetails(select) {
		var selected = [];
		var index;
		if (!select || !select.options || select.options.length === 0) {
			return "";
		}
		if (select.size > 1 || select.multiple) {
			for (index = 0; index < select.options.length; index += 1) {
				if (select.options[index].selected) {
					selected.push(select.options[index].text);
				}
			}
			return selected.join(":");
		}
		return selectedText(select);
	}

	function addOption(select, text, value) {
		if (select) {
			select.add(new Option(text == null ? "" : String(text), value == null ? "" : String(value)));
		}
	}

	function setSelectLength(select, length) {
		if (select) {
			select.options.length = length;
		}
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function openDialog(url, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", features || "", callback || function () {});
			return;
		}
		window.open(url, "_blank", "height=500,width=420,resizable=no,status=no,scrollbars=yes");
	}

	function runStringDialog(page, query, features, done) {
		openDialog(page + "?" + query, features, function (outValue) {
			var text = trim(outValue);
			var parts;
			if (text === "") {
				return;
			}
			parts = text.split(":");
			if (parts.length === 1) {
				runStringDialog(page, text, features, done);
				return;
			}
			done(text, parts);
		});
	}

	function populateSelectFromXml(select, root, textIndex, valueIndex) {
		childElements(root).forEach(function (node) {
			addOption(select, attr(node, textIndex), attr(node, valueIndex));
		});
	}

	function lastAccHeadNode() {
		var nodes = childElements(xmlRoot("AccHeadData"));
		return nodes.length ? nodes[nodes.length - 1] : null;
	}

	function addPartyHeadXml(code, name, value2) {
		if (typeof window.GetPartyHeadXml === "function") {
			window.GetPartyHeadXml(code, name, value2);
		}
	}

	function setPartySelection(partyType, parts, select) {
		var node;
		if (parts.length <= 2) {
			return false;
		}
		addPartyHeadXml(parts[1] || "", parts[0] || "", "0:0:0");
		node = lastAccHeadNode();
		if (!node) {
			if (select) {
				select.selectedIndex = 0;
			}
			return false;
		}
		setValue("hPartyCode", partyType + "?" + attr(node, 0));
		setValue("txtPartyName", attr(node, 3));
		return true;
	}

	function setCommonHeaderValues() {
		var unit = field("selUnitId");
		var book = field("selBook");
		setValue("horgName", selectedText(unit));
		setValue("hBookName", selectedText(book));
		setValue("hVouDetails", selectedDetails(field("selInvoiceNo")));
	}

	function disableIfPresent(name, disabled) {
		var item = field(name);
		if (item) {
			item.disabled = !!disabled;
		}
	}

	function submitTo(action) {
		var frm = form();
		if (!frm) {
			return;
		}
		frm.action = action;
		frm.submit();
	}

	function installCredit() {
		window.popPartyType = function () {
			var select = field("selPartyType");
			var xhr;
			setSelectLength(select, 1);
			xhr = syncGet("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(valueOf("selUnitId")));
			loadResponseXml("OutData", xhr);
			childElements(xmlRoot("OutData")).forEach(function (node) {
				addOption(select, trim(node.textContent || node.text || ""), attr(node, 0));
			});
			return true;
		};

		window.DisplayBook = function () {
			var select = field("selBook");
			var orgId = valueOf("selUnitId");
			var chkGj = field("ChkGJ");
			var bookCode = chkGj && chkGj.checked ? "08" : "07";
			var xhr;
			setSelectLength(select, 1);
			if (!orgId || orgId === "0") {
				return false;
			}
			xhr = syncGet("XMLGetOrgBook.asp?BkCode=" + bookCode + "&orgID=" + encodeURIComponent(orgId));
			loadResponseXml("UnitBookData", xhr);
			populateSelectFromXml(select, xmlRoot("UnitBookData"), 1, 0);
			window.popPartyType();
			return true;
		};

		window.NewPage = function () {
			submitTo("CreditNoteNewPage.asp");
		};

		window.validate = function () {
			var unit = field("selUnitId");
			var book = field("selBook");
			var voucherType = field("selVoucherType");
			if (unit && unit.selectedIndex < 1) {
				alert("Select Unit");
				unit.focus();
				return false;
			}
			if (book && book.selectedIndex < 1) {
				alert("Select Credit Note");
				book.focus();
				return false;
			}
			if (voucherType && voucherType.selectedIndex < 1) {
				alert("Select Voucher type");
				voucherType.focus();
				return false;
			}
			setCommonHeaderValues();
			return true;
		};

		window.VouCreate = function () {
			var invoice = field("selInvoiceNo");
			var voucherType = valueOf("selVoucherType");
			var detailsParts;
			setValue("hChkVal", field("ChkGJ") && field("ChkGJ").checked ? "Y" : "N");
			if (!window.validate()) {
				return false;
			}
			if (invoice && invoice.value === "S" && voucherType !== "OT") {
				alert("Select Invoice ");
				invoice.focus();
				return false;
			}
			if (voucherType === "SC") {
				form().action = "VouCNCommisionEntry.asp";
			} else if (voucherType === "SR") {
				form().action = "VouCNSalReturnEntry.asp";
			} else if (voucherType === "OT") {
				form().action = "VouCNOthersEntry.asp";
			} else if (voucherType === "OIS") {
				form().action = "VouCNOtherInvEntry.asp";
			} else if (voucherType === "OIP") {
				form().action = "VouCNPurInvEntry.asp";
			}
			detailsParts = String(valueOf("hVouDetails")).split(":");
			if (voucherType === "OSC" || voucherType === "OPC") {
				form().action = detailsParts.length !== 1 ? "VouCNOthersEntry.asp" : "VouCNOthCommEntry.asp";
			}
			window.DisButt();
			form().submit();
			return false;
		};

		window.selParty = function (select) {
			var orgId = valueOf("selUnitId");
			var invoice = field("selInvoiceNo");
			var voucherType = field("selVoucherType");
			var partyType = select.value + "?" + selectedText(select);
			setSelectLength(voucherType, 1);
			if (invoice) {
				invoice.options.length = invoice.multiple ? 0 : 1;
				invoice.size = invoice.multiple ? 4 : 1;
			}
			setValue("txtPartyName", "");
			setValue("hPartyCode", "");
			if (!select || select.selectedIndex <= 0) {
				return false;
			}
			runStringDialog("PartySelection.asp", "orgId=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyType), "dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (outValue, parts) {
				var xhr;
				if (!setPartySelection(partyType, parts, select)) {
					return;
				}
				if (String(select.value) !== "CR?1") {
					addOption(voucherType, "Sales Return", "SR");
					addOption(voucherType, "Sales Invoices", "OIS");
					addOption(voucherType, "Others", "OT");
					addOption(voucherType, "Purchase Invoices", "OIP");
					xhr = syncGet("XMLSalInvDetails.asp?BookCode=05&OrgId=" + encodeURIComponent(orgId) + "&PartyCode=" + encodeURIComponent(valueOf("hPartyCode")));
					loadResponseXml("UnitBookData", xhr);
					childElements(xmlRoot("UnitBookData")).forEach(function (node) {
						addOption(invoice, attr(node, "InvDetails"), attr(node, "TransNo") + ":" + attr(node, "TotalCrDrValue"));
					});
					if (invoice) {
						invoice.multiple = false;
						invoice.size = 1;
					}
				} else {
					addOption(voucherType, "Sales Commision", "SC");
					addOption(voucherType, "Others", "OT");
					addOption(voucherType, "Sales Invoices", "OSC");
					addOption(voucherType, "Purchase Invoices", "OPC");
					if (invoice) {
						invoice.multiple = true;
						invoice.size = 4;
						invoice.options.length = 0;
					}
					xhr = syncGet("XMLCommisionDetails.asp?OrgId=" + encodeURIComponent(orgId) + "&AgentCode=" + encodeURIComponent(valueOf("hPartyCode")));
					loadResponseXml("UnitBookData", xhr);
					childElements(xmlRoot("UnitBookData")).forEach(function (node) {
						addOption(invoice, attr(node, 2), attr(node, 0));
					});
				}
				return outValue;
			});
			return false;
		};

		window.setInvoiceNo = function () {
			var invoice = field("selInvoiceNo");
			var value = valueOf("selVoucherType");
			if (value === "SR") {
				setSelectLength(invoice, 1);
				if (invoice) {
					invoice.size = 1;
				}
				window.PopOthInv("S");
			}
			disableIfPresent("selInvoiceNo", value === "OT");
			if (value === "SC") {
				window.popCommVouchers();
			}
			if (value === "OIS") {
				setSelectLength(invoice, 1);
				if (invoice) {
					invoice.size = 1;
				}
				window.PopOthInv("S");
			}
			if (value === "OIP") {
				window.PopOthInv("P");
			}
			if (value === "OSC") {
				window.PopOthComm("S");
			}
			if (value === "OPC") {
				setSelectLength(invoice, 1);
				if (invoice) {
					invoice.size = 1;
				}
				window.PopOthComm("P");
			}
		};

		window.PopOthInv = function (type) {
			var invoice = field("selInvoiceNo");
			var vouType = valueOf("selVoucherType") === "SR" ? "SR" : "O";
			var xhr = syncGet("XMLSalInvDetails.asp?BookCode=45&OrgId=" + encodeURIComponent(valueOf("selUnitId")) + "&Type=" + encodeURIComponent(vouType) + "&PartyCode=" + encodeURIComponent(valueOf("hPartyCode")) + "&sCallTy=" + encodeURIComponent(vouType));
			loadResponseXml("UnitBookData", xhr);
			childElements(xmlRoot("UnitBookData")).forEach(function (node) {
				if (String(attr(node, 5)) === String(type)) {
					addOption(invoice, attr(node, 2), attr(node, 0) + ":" + attr(node, "FromValue"));
				}
			});
		};

		window.PopOthComm = function (type) {
			var invoice = field("selInvoiceNo");
			var bookCode = String(type) === "S" ? "05" : "04";
			var xhr = syncGet("XMLInvDetails.asp?BookCode=" + bookCode + "&OrgId=" + encodeURIComponent(valueOf("selUnitId")) + "&PartyCode=" + encodeURIComponent(valueOf("hPartyCode")) + "&sCallTy=S");
			var root;
			loadResponseXml("CommData", xhr);
			root = xmlRoot("CommData");
			if (invoice) {
				invoice.options.length = 0;
				invoice.multiple = true;
				invoice.size = childElements(root).length > 8 ? 8 : Math.max(childElements(root).length, 4);
			}
			childElements(root).forEach(function (node) {
				addOption(invoice, attr(node, 2), attr(node, 0));
			});
		};

		window.popCommVouchers = function () {
			var invoice = field("selInvoiceNo");
			var xhr = syncGet("XMLCommisionDetails.asp?OrgId=" + encodeURIComponent(valueOf("selUnitId")) + "&AgentCode=" + encodeURIComponent(valueOf("hPartyCode")));
			var root;
			if (invoice) {
				invoice.options.length = 0;
				invoice.multiple = true;
			}
			loadResponseXml("UnitBookData", xhr);
			root = xmlRoot("UnitBookData");
			if (invoice) {
				invoice.size = childElements(root).length > 8 ? 8 : childElements(root).length;
			}
			childElements(root).forEach(function (node) {
				addOption(invoice, attr(node, 2), attr(node, 0));
			});
		};

		window.popVoucherNo = function (callType) {
			var unit = field("selUnitId");
			var book = field("selBook");
			var voucherType = field("selVoucherType");
			var page;
			var url;
			if (unit && unit.selectedIndex === 0) {
				alert("Select Unit ");
				unit.focus();
				return;
			}
			if (book && book.selectedIndex === 0) {
				alert("Select Book ");
				book.focus();
				return;
			}
			if (voucherType && voucherType.selectedIndex === 0) {
				alert("Select Voucher Type ");
				voucherType.focus();
				return;
			}
			page = valueOf("selVoucherType") === "SR" ? "VouchSelForCNDN.asp" : "VouchSelForSalPur.asp";
			url = page + "?flag=" + encodeURIComponent(callType) +
				"&orgId=" + encodeURIComponent(valueOf("selUnitId")) +
				"&BookCode=07&BookNo=" + encodeURIComponent(valueOf("selBook")) +
				"&TransType=CNR&sPurTy=0&sParTy=" + encodeURIComponent(valueOf("selPartyType")) +
				"&iParCode=" + encodeURIComponent(valueOf("hPartyCode")) +
				"&VouchTy=" + encodeURIComponent(valueOf("selVoucherType"));
			openDialog(url, "dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var parts;
				if (String(value || "0") === "0") {
					return;
				}
				parts = String(value).split("~");
				setValue("TransNo", parts[0] || "");
				setValue("hTransNo", parts[0] || "");
				setValue("txtVouchNo", valueOf("selVoucherType") === "SR" ? parts[1] || "" : parts[2] || "");
			});
		};

		window.VouView = function () {
			var voucherType;
			if (!window.validate()) {
				return false;
			}
			voucherType = valueOf("selVoucherType");
			if (voucherType === "SR") {
				submitTo("VouCNSalReturnDisplay.asp");
			} else if (voucherType === "SC") {
				window.DisButt();
				submitTo("VouCNCommisionView.asp");
			} else if (voucherType === "OIS") {
				window.DisButt();
				submitTo("VouCNOtherInvDisplay.asp");
			} else if (voucherType === "OIP") {
				window.DisButt();
				submitTo("VouDNPurReturnDisplay.asp");
			} else {
				window.DisButt();
				submitTo("VouCNOtherView.asp");
			}
			return false;
		};

		window.VouDel = function () {
			if (window.validate()) {
				window.DisButt();
				submitTo("VouCNCommDelView.asp");
			}
			return false;
		};

		window.VouAmend = function () {
			var voucherType;
			if (field("selUnitId") && field("selUnitId").selectedIndex < 1) {
				alert("Select Unit");
				return false;
			}
			if (field("selBook") && field("selBook").selectedIndex < 1) {
				alert("Select Credit Note");
				return false;
			}
			if (trim(valueOf("txtVouchNo")) === "") {
				alert("Select Voucher No ");
				return false;
			}
			setCommonHeaderValues();
			voucherType = valueOf("selVoucherType");
			if (voucherType === "SC") {
				form().action = "VouCNCommissionAmd.asp";
			} else if (voucherType === "SR") {
				form().action = "VouCNSalRetEntryAmd.asp";
			} else if (voucherType === "OIP") {
				form().action = "VouCNPurInvAmd.asp";
			} else if (voucherType === "OIS") {
				form().action = "VouCNSalInvAmd.asp";
			} else {
				form().action = "VouCNCommAmend.asp";
			}
			window.DisButt();
			form().submit();
			return false;
		};
	}

	function installDebit() {
		window.popPartyType = function () {
			var select = field("selPartyType");
			var xhr;
			setSelectLength(select, 1);
			xhr = syncGet("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(valueOf("selUnitId")));
			loadResponseXml("OutData", xhr);
			childElements(xmlRoot("OutData")).forEach(function (node) {
				addOption(select, trim(node.textContent || node.text || ""), attr(node, 0));
			});
			return true;
		};

		window.DisplayBook = function (objUnit) {
			var select = field("selBook");
			var unit = objUnit || field("selUnitId");
			var orgId;
			var xhr;
			setSelectLength(select, 1);
			if (!unit || unit.selectedIndex === 0 || String(unit.selectedIndex) === "0") {
				return false;
			}
			orgId = unit.options[unit.selectedIndex].value;
			xhr = syncGet("XMLGetOrgBook.asp?BkCode=06&orgID=" + encodeURIComponent(orgId));
			loadResponseXml("UnitBookData", xhr);
			populateSelectFromXml(select, xmlRoot("UnitBookData"), 1, 0);
			window.popPartyType();
			return true;
		};

		window.NewPage = function () {
			submitTo("DebitNoteNewPage.asp");
		};

		window.validate = function () {
			var unit = field("selUnitId");
			var book = field("selBook");
			var voucherType = field("selVoucherType");
			var invoice = field("selInvoiceNo");
			var ref = field("selRefNo");
			if (unit && unit.selectedIndex < 1) {
				alert("Select Unit");
				unit.focus();
				return false;
			}
			if (book && book.selectedIndex < 1) {
				alert("Select Debit Note");
				book.focus();
				return false;
			}
			if (voucherType && voucherType.selectedIndex < 1) {
				alert("Select Voucher type");
				voucherType.focus();
				return false;
			}
			if (valueOf("selVoucherType") !== "OT" && invoice && ref && invoice.selectedIndex < 1 && ref.selectedIndex < 1) {
				alert("Select Invoice No");
				return false;
			}
			setCommonHeaderValues();
			return true;
		};

		window.VouCreate = function () {
			if (!window.validate()) {
				return false;
			}
			if (valueOf("selVoucherType") === "OT") {
				form().action = "VouDNOthersEntry.asp";
			} else if (valueOf("selVoucherType") === "OP") {
				form().action = "VouDNOtherInvEntry.asp";
			} else if (valueOf("selVoucherType") === "SI") {
				form().action = "VouDNSalInvEntry.asp";
			}
			form().submit();
			return false;
		};

		window.selParty = function (select) {
			var orgId = valueOf("selUnitId");
			var partyType = select.value + "?" + selectedText(select);
			setSelectLength(field("selInvoiceNo"), 1);
			setSelectLength(field("selRefNo"), 1);
			setValue("txtPartyName", "");
			setValue("hPartyCode", "");
			if (!select || select.selectedIndex <= 0) {
				return false;
			}
			runStringDialog("PartySelection.asp", "orgId=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyType), "dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (outValue, parts) {
				setPartySelection(partyType, parts, select);
				return outValue;
			});
			return false;
		};

		window.PopComm = function () {
			var invoice = field("selInvoiceNo");
			var xhr = syncGet("XMLCommisionDetails.asp?OrgId=" + encodeURIComponent(valueOf("selUnitId")) + "&AgentCode=" + encodeURIComponent(valueOf("hPartyCode")));
			loadResponseXml("UnitBookData", xhr);
			childElements(xmlRoot("UnitBookData")).forEach(function (node) {
				addOption(invoice, attr(node, 2), attr(node, 0));
			});
		};

		window.PopOthInv = function (callType) {
			var invoice = field("selInvoiceNo");
			var ref = field("selRefNo");
			var xhr = syncGet("XMLSalInvDetails.asp?BookCode=04&OrgId=" + encodeURIComponent(valueOf("selUnitId")) + "&PartyCode=" + encodeURIComponent(valueOf("hPartyCode")) + "&sCallTy=O");
			loadResponseXml("UnitBookData", xhr);
			childElements(xmlRoot("UnitBookData")).forEach(function (node) {
				var nodeName = String(node.nodeName || "").toLowerCase();
				var value;
				if (nodeName === "salinv" && String(callType) === String(attr(node, 5))) {
					addOption(invoice, attr(node, 2), attr(node, 0));
				}
				if (nodeName === "invoice") {
					value = trim(attr(node, "ActNo")) + "-" + trim(attr(node, "InvNo")) + "-" + trim(attr(node, "InvDate"));
					addOption(ref, value, trim(attr(node, "ActNo")));
				}
			});
		};

		window.setInvoiceNo = function () {
			var value = valueOf("selVoucherType");
			var invoice = field("selInvoiceNo");
			var ref = field("selRefNo");
			if (value === "OT") {
				disableIfPresent("selInvoiceNo", true);
				disableIfPresent("selRefNo", true);
				disableIfPresent("btnAmend", false);
				return;
			}
			disableIfPresent("selInvoiceNo", false);
			disableIfPresent("selRefNo", false);
			setSelectLength(invoice, 1);
			setSelectLength(ref, 1);
			if (value === "OP") {
				window.PopOthInv("P");
			} else if (value === "SI") {
				window.PopOthInv("S");
			} else if (value === "SC") {
				window.PopComm();
			} else {
				window.PopOthInv("P");
			}
		};

		window.popVoucherNo = function (callType) {
			var unit = field("selUnitId");
			var book = field("selBook");
			var voucherType = field("selVoucherType");
			var page;
			var url;
			if (unit && unit.selectedIndex === 0) {
				alert("Select Unit ");
				unit.focus();
				return;
			}
			if (book && book.selectedIndex === 0) {
				alert("Select Book ");
				book.focus();
				return;
			}
			if (voucherType && voucherType.selectedIndex === 0) {
				alert("Select Voucher Type ");
				voucherType.focus();
				return;
			}
			page = valueOf("selVoucherType") === "PR" ? "VouchSelForCNDN.asp" : "VouchSelForSalPur.asp";
			url = page + "?flag=" + encodeURIComponent(callType) +
				"&orgId=" + encodeURIComponent(valueOf("selUnitId")) +
				"&BookCode=06&BookNo=" + encodeURIComponent(valueOf("selBook")) +
				"&TransType=DNR&sPurTy=0&sParTy=" + encodeURIComponent(valueOf("selPartyType")) +
				"&iParCode=" + encodeURIComponent(valueOf("hPartyCode")) +
				"&VouchTy=" + encodeURIComponent(valueOf("selVoucherType"));
			openDialog(url, "dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function (value) {
				var parts;
				if (String(value || "0") === "0") {
					return;
				}
				parts = String(value).split("~");
				setValue("hTransNo", parts[0] || "");
				setValue("txtVouchNo", parts[2] || "");
				disableIfPresent("btnAmend", String(callType) === "A");
				disableIfPresent("btnDelete", String(callType) === "A");
			});
		};

		window.PopOthComm = function (type) {
			var invoice = field("selInvoiceNo");
			var bookCode = String(type) === "S" ? "05" : "04";
			var xhr = syncGet("XMLInvDetails.asp?BookCode=" + bookCode + "&OrgId=" + encodeURIComponent(valueOf("selUnitId")) + "&PartyCode=" + encodeURIComponent(valueOf("hPartyCode")) + "&sCallTy=S");
			loadResponseXml("CommData", xhr);
			if (invoice) {
				invoice.options.length = 0;
				invoice.size = 8;
				invoice.multiple = true;
			}
			childElements(xmlRoot("CommData")).forEach(function (node) {
				addOption(invoice, attr(node, 2), attr(node, 0));
			});
		};

		window.SelInvChoice = function () {
			disableIfPresent("selRefNo", trim(valueOf("selInvoiceNo")) !== "S");
		};

		window.SelRefChoice = function () {
			disableIfPresent("selInvoiceNo", trim(valueOf("selRefNo")) !== "S");
		};
	}

	function install(options) {
		ensureCompat();
		options = options || {};
		if (options.mode === "debit") {
			installDebit();
		} else {
			installCredit();
		}
		window.DisButt = function () {
			disableIfPresent("btnCreate", true);
		};
	}

	window.ITMSCreditDebitNoteBookSelectionCompat = {
		install: install
	};
})(window, document);
