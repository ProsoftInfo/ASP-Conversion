(function (window, document) {
	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname;
	}

	function byId(id) {
		var direct = document.getElementById(id) || window[id];
		if (direct) {
			return direct;
		}
		var wanted = String(id).toLowerCase();
		var elements = document.querySelectorAll("[id]");
		for (var i = 0; i < elements.length; i += 1) {
			if (String(elements[i].id).toLowerCase() === wanted) {
				return elements[i];
			}
		}
		return null;
	}

	function textOf(id) {
		var element = byId(id);
		return element ? (element.innerText || element.textContent || element.innerHTML || "") : "";
	}

	function setText(id, value) {
		var element = byId(id);
		if (!element) {
			return;
		}
		if ("innerText" in element) {
			element.innerText = value;
		} else {
			element.textContent = value;
		}
	}

	function selectedValue(select) {
		if (!select || !select.options) {
			return "";
		}
		var option = select.options[select.selectedIndex];
		return option ? option.value : "";
	}

	function getDateValue(control) {
		if (!control) {
			return "";
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		return control.value || "";
	}

	function setDateValue(control, value) {
		if (!control) {
			return;
		}
		if (typeof control.setDate === "function") {
			control.setDate(value);
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
			return;
		}
		control.value = value || "";
	}

	function xmlIsland(name) {
		return window[name] || document[name] || document.getElementById(name);
	}

	function xmlString(node) {
		if (!node) {
			return "";
		}
		if (typeof node.xml === "string") {
			return node.xml;
		}
		return new XMLSerializer().serializeToString(node);
	}

	function createHttp() {
		return window.CreateObject ? window.CreateObject("Microsoft.XMLHTTP") : new XMLHttpRequest();
	}

	function postXml(url, node) {
		var xhr = createHttp();
		xhr.open("POST", url, false);
		xhr.send(xmlString(node));
		return xhr;
	}

	window.ChangePaymentMode = function () {
		var pay = form().radPayThru;
		var cashChecked = pay && pay[0] && pay[0].checked;
		var chequeNo = byId("tdChequeNo");
		var chequeDate = byId("tdChequeDate");
		if (chequeNo) {
			chequeNo.style.display = cashChecked ? "none" : "block";
		}
		if (chequeDate) {
			chequeDate.style.display = cashChecked ? "none" : "block";
		}
	};

	window.PaymentForChange = function () {
		var sValue = selectedValue(form().selPayFor);
		if (sValue === "O") {
			form().txtPayFor.value = "Payment for ";
			form().txtPayFor.disabled = false;
		} else if (sValue === "F") {
			form().txtPayFor.value = "Freight Payment for | " + textOf("RefNoDate");
			form().txtPayFor.disabled = true;
		} else {
			alert("Select Payment For ");
			form().selPayFor.focus();
		}
	};

	window.LoadPartySubType = function () {
		var xhr = createHttp();
		var partyData = xmlIsland("PartyData");
		var select = form().selPartySubType;
		xhr.open("GET", "../PartySubType.asp?ParCode=" + encodeURIComponent(form().hSupplierCode.value) + "&OrgCode=" + encodeURIComponent(form().hOrgID.value), false);
		xhr.send(null);
		if (xhr.responseXML && xhr.responseXML.documentElement && partyData && typeof partyData.loadXML === "function") {
			partyData.loadXML(xmlString(xhr.responseXML));
		} else if (trim(xhr.responseText) !== "" && partyData && typeof partyData.loadXML === "function") {
			partyData.loadXML(xhr.responseText);
		} else if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
		}

		var root = partyData && partyData.documentElement;
		if (!root || !root.hasChildNodes || !root.hasChildNodes()) {
			return;
		}
		select.length = 0;
		Array.prototype.forEach.call(root.childNodes, function (child) {
			if (child.nodeType !== 1) {
				return;
			}
			var option = new Option(child.textContent || child.text || "", child.getAttribute("SubType") || "");
			select.add(option);
		});
	};

	function applyReferenceSelection(refType, orgId) {
		var outData = xmlIsland("OutData");
		var root = outData && outData.documentElement;
		if (trim(refType) === "N" || !root || !root.hasChildNodes || !root.hasChildNodes()) {
			return;
		}
		Array.prototype.forEach.call(root.childNodes, function (refNode) {
			if (!refNode || refNode.nodeName !== "Reference") {
				return;
			}
			var receiptNo = refNode.getAttribute("ReferenceNo") || "";
			var refNoDate = (refNode.getAttribute("ReferenceCode") || "") + " - " + (refNode.getAttribute("ReferenceDate") || "");
			var partyCode = form().hSupplierCode.value || "";
			var remarks = trim(refNode.getAttribute("Remarks"));
			setText("RefNoDate", refNoDate);
			form().hRefTypeCode.value = refType;
			form().hRefno.value = receiptNo;
			form().hOrgID.value = orgId;
			form().hRefDate.value = refNode.getAttribute("ReferenceDate") || "";
			if (remarks !== "") {
				var parts = remarks.split("-");
				partyCode = parts[0] || "";
				setText("idSupplier", parts.slice(1).join("-"));
			}
			form().hSupplierCode.value = partyCode;
			form().hSupplierName.value = textOf("idSupplier");

			var voucherRoot = xmlIsland("VoucherData").documentElement;
			voucherRoot.setAttribute("ReferenceNo", textOf("RefNoDate"));
			voucherRoot.setAttribute("hRefNo", receiptNo);
			voucherRoot.setAttribute("AppRefNo", receiptNo);
			voucherRoot.setAttribute("AppRefDate", form().hRefDate.value);
			voucherRoot.setAttribute("AppRefType", refType);
			voucherRoot.setAttribute("PartyCode", partyCode);
			window.LoadPartySubType();
		});
	}

	window.RefType_Click = function () {
		var refType = selectedValue(form().selRefName || form().SelRefName);
		var orgId = form().hOrgID.value;
		var partyCode = form().hSupplierCode.value;
		if (trim(refType) === "N" && trim(partyCode) === "") {
			alert("Select the Party");
			return;
		}
		window.RefTypeSelection(refType, orgId, partyCode, "N", 1, "Y", 0, "PUR", function () {
			applyReferenceSelection(refType, orgId);
		});
	};

	window.AddNewParty = function () {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog("MisParCreate.asp?", "", "dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (value) {
			setText("idPayTo", value || "");
		});
	};

	window.SelMisParty = function () {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		function handle(value) {
			if (String(value) === "AN") {
				window.AddNewParty();
				return;
			}
			var parts = String(value || "").split(":");
			if (parts.length <= 1) {
				window.ITMSModernCompat.openModalDialog("../../Common/MisPartySelection.asp?" + encodeURIComponent(value || ""), "", "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", handle);
				return;
			}
			setText("idPayTo", parts[0] || "");
			form().hMisPartyCode.value = parts[1] || "";
			if (parts[0]) {
				form().txtPayTo.readOnly = true;
			}
		}
		window.ITMSModernCompat.openModalDialog("../../Common/MisPartySelection.asp", "", "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", handle);
	};

	window.Supplierselect = function () {
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		function handle(outValue) {
			if (!outValue || !outValue.hasChildNodes || !outValue.hasChildNodes()) {
				return;
			}
			var action = trim(outValue.getAttribute("Action")).toUpperCase();
			if (action && action !== "DONE" && action !== "CLOSE") {
				window.ITMSModernCompat.openModalDialog("../SupplierSelect.asp?" + trim(outValue.getAttribute("PassQuery")), xmlIsland("OutData"), "status:no", handle);
				return;
			}
			var codes = [];
			var names = [];
			Array.prototype.forEach.call(outValue.childNodes, function (node) {
				if (String(node.nodeName).toUpperCase() === "SUPPLIER") {
					codes.push(trim(node.getAttribute("SuppCode")));
					names.push(trim(node.getAttribute("SuppName")));
				}
			});
			setText("idSupplier", names.join(","));
			form().hSupplierCode.value = codes.join(",");
			form().hSupplierName.value = names.join(",");
			window.LoadPartySubType();
		}
		window.ITMSModernCompat.openModalDialog(
			"../SupplierSelect.asp?Unit=" + encodeURIComponent(form().hOrgID.value) + "&hSelectMode=S&Flag=1&ParType=" + encodeURIComponent(form().hParType.value),
			xmlIsland("OutData"),
			"status:no",
			handle
		);
	};

	window.CheckSubmit = function () {
		if (form().hSupplierCode.value === "") {
			alert("Select Party");
			return;
		}
		if (textOf("idPayTo") === "" && form().txtPayTo.value === "") {
			alert("Select Pay to Received from Party or Enter the party name in Textbox");
			return;
		}
		if (form().selPayFor.selectedIndex === 0) {
			alert("Select Payment For");
			form().selPayFor.focus();
			return;
		}

		var payFor = selectedValue(form().selPayFor);
		var payMode = form().radPayThru && form().radPayThru[0] && form().radPayThru[0].checked ? form().radPayThru[0].value : form().radPayThru[1].value;
		var adjustAgainstInvoice = form().AdjAgainBill && form().AdjAgainBill.checked ? "Y" : "N";
		var partyParts = selectedValue(form().selPartySubType).split("|");
		var checkPayToSelection = textOf("idPayTo") !== "" ? "YES" : "NO";
		var chequeNo = "";
		var chequeDate = "";
		if (form().radPayThru && form().radPayThru[1] && form().radPayThru[1].checked) {
			chequeNo = form().txtChequeNo.value;
			chequeDate = getDateValue(form().ctlChequeDate);
		}

		var voucherRoot = xmlIsland("VoucherData").documentElement;
		voucherRoot.setAttribute("VouDate", getDateValue(form().ctlInvoiceDate));
		voucherRoot.setAttribute("PartyType", partyParts[0] || "");
		voucherRoot.setAttribute("PartySubType", partyParts[1] || "");
		voucherRoot.setAttribute("PayFor", payFor);
		voucherRoot.setAttribute("hPayFor", payFor);
		voucherRoot.setAttribute("PaymentThru", payMode);
		voucherRoot.setAttribute("Code", form().hSupplierCode.value);
		voucherRoot.setAttribute("CheNo", chequeNo);
		voucherRoot.setAttribute("CheDate", chequeDate);

		if (form().hMiscInvNo) {
			voucherRoot.setAttribute("ReferenceNo", textOf("RefNoDate"));
			voucherRoot.setAttribute("hRefNo", form().hAppRefNo.value);
			voucherRoot.setAttribute("AppRefNo", form().hAppRefNo.value);
			voucherRoot.setAttribute("AppRefDate", trim(textOf("RefNoDate")) !== "N/A" ? (textOf("RefNoDate").split("-")[1] || "") : "");
			voucherRoot.setAttribute("AppRefType", selectedValue(form().SelRefName || form().selRefName));
			voucherRoot.setAttribute("PartyCode", form().hSupplierCode.value);
		}

		var entryRoot = xmlIsland("EntryData").documentElement;
		entryRoot.setAttribute("No", "1");
		entryRoot.setAttribute("CRDR", "C");
		entryRoot.setAttribute("Payto", form().hSupplierName.value);
		entryRoot.setAttribute("Amount", form().txtAmount.value);
		entryRoot.setAttribute("PayToSelCheck", checkPayToSelection);
		entryRoot.setAttribute("MiscPartyName", textOf("idPayTo") !== "" ? textOf("idPayTo") : form().txtPayTo.value);
		entryRoot.setAttribute("MiscPartyCode", form().hMisPartyCode.value);
		entryRoot.setAttribute("CheckVal", adjustAgainstInvoice);

		var narration = xmlIsland("EntryData").createElement("Narration");
		narration.textContent = form().txtPayFor.value;
		entryRoot.appendChild(narration);
		voucherRoot.appendChild(entryRoot);

		if (form().hMiscInvNo) {
			postXml("XMLSave.asp?Name=MISCPaymentEdit&SessionFlag=true", voucherRoot);
			form().action = "MiscInvoiceUpdate.asp?InvNo=" + encodeURIComponent(form().hMiscInvNo.value);
		} else {
			postXml("XMLSave.asp?Name=MISCPayment&SessionFlag=true", voucherRoot);
			form().action = "MiscInvoiceInsert.asp";
		}
		form().submit();
	};

	window.setdate = function () {
		var fromDate = form().hFromDate && form().hFromDate.value;
		var toDate = form().hToDate && form().hToDate.value;
		if (form().ctlInvoiceDate) {
			form().ctlInvoiceDate.setMinDate = fromDate;
			form().ctlInvoiceDate.setMaxDate = toDate;
			setDateValue(form().ctlInvoiceDate, toDate);
		}
	};

	window.Init = function () {
		var refSelect = form().SelRefName || form().selRefName;
		var subtypeSelect = form().selPartySubType;
		var payForSelect = form().selPayFor;
		var i;
		for (i = 0; refSelect && i < refSelect.length; i += 1) {
			if (refSelect.options[i].value === form().hAppRefType.value) {
				refSelect.selectedIndex = i;
				break;
			}
		}
		if (trim(form().hRefNoDate.value) !== "") {
			setText("RefNoDate", form().hRefNoDate.value.split(",")[1] || "");
		}
		setText("idSupplier", form().hSupplierName.value);
		window.LoadPartySubType();
		for (i = 0; subtypeSelect && i < subtypeSelect.length; i += 1) {
			if (subtypeSelect.options[i].value === form().hSubType.value) {
				subtypeSelect.selectedIndex = i;
				break;
			}
		}
		for (i = 0; payForSelect && i < payForSelect.length; i += 1) {
			if (payForSelect.options[i].value === form().hPayFor.value) {
				payForSelect.selectedIndex = i;
				break;
			}
		}
		if (form().ctlChequeDate && form().hChequeDate) {
			setDateValue(form().ctlChequeDate, form().hChequeDate.value);
		}
		window.PaymentForChange();
		window.ChangePaymentMode();
	};
})(window, document);
