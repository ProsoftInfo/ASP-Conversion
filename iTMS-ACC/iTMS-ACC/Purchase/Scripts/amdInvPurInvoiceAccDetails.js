(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function numberValue(value) {
		var normalized = String(value == null ? "" : value).replace(/,/g, "");
		var parsed = parseFloat(normalized);
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value, decimals) {
		return numberValue(value).toFixed(decimals == null ? 2 : decimals);
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements[name];
	}

	function invoiceDocument() {
		if (window.InvoiceDet && window.InvoiceDet.XMLDocument) {
			return window.InvoiceDet.XMLDocument;
		}
		if (window.InvoiceDet && window.InvoiceDet._doc) {
			return window.InvoiceDet._doc;
		}
		if (window.InvoiceDet && window.InvoiceDet.nodeType === 9) {
			return window.InvoiceDet;
		}
		return window.InvoiceDet && window.InvoiceDet.ownerDocument || document.implementation.createDocument("", "Root", null);
	}

	function invoiceRoot() {
		return invoiceDocument().documentElement || window.InvoiceDet.documentElement;
	}

	function selectNodes(context, expression) {
		if (!context) {
			return [];
		}
		if (typeof context.selectNodes === "function") {
			return context.selectNodes(expression);
		}
		return [];
	}

	function childElements(node) {
		return Array.prototype.filter.call(node && node.childNodes || [], function (child) {
			return child.nodeType === 1;
		});
	}

	function createElement(name) {
		if (window.InvoiceDet && typeof window.InvoiceDet.createElement === "function") {
			return window.InvoiceDet.createElement(name);
		}
		return invoiceDocument().createElement(name);
	}

	function firstElement(node) {
		return childElements(node)[0] || null;
	}

	function setSelectToValue(select, value) {
		if (!select || !select.options) {
			return;
		}
		for (var index = 0; index < select.options.length; index += 1) {
			if (select.options[index].value === value) {
				select.selectedIndex = index;
				return;
			}
		}
	}

	function accountHeadValue(index) {
		var select = field("mCmbItemAccHeadZ" + index);
		return trim(select && select.value);
	}

	function accountHeadSelect(index) {
		return field("mCmbItemAccHeadZ" + index);
	}

	function accountHeadNo(value) {
		var pipeAt = trim(value).indexOf("|");
		return pipeAt >= 0 ? trim(value).substring(0, pipeAt) : trim(value);
	}

	function findRootChild(root, name) {
		var wanted = String(name).toLowerCase();
		return childElements(root).filter(function (node) {
			return String(node.nodeName).toLowerCase() === wanted;
		})[0] || null;
	}

	function findVoucherDetails(root) {
		var voucher = findRootChild(root, "Voucher");
		if (!voucher) {
			return null;
		}
		return childElements(voucher).filter(function (node) {
			return String(node.nodeName).toLowerCase() === "details";
		})[0] || null;
	}

	function xmlText(value) {
		var target = value && (value.XMLDocument || value._doc || value);
		if (!target) {
			return "";
		}
		if (typeof target.xml === "string") {
			return target.xml;
		}
		return new XMLSerializer().serializeToString(target);
	}

	window.DisplayTotal = function (nPassCtr) {
		var root = invoiceRoot();
		var itemNodes = selectNodes(root, "//ItemDetails/Item");
		var firstAccountHead = accountHeadValue(1);
		var accountHeads = [];
		var totalValue = numberValue(field("hTotalTaxValue") && field("hTotalTaxValue").value);
		var index;
		var itemAccountHead;
		var select;

		for (index = 1; index < itemNodes.length; index += 1) {
			setSelectToValue(accountHeadSelect(index + 1), firstAccountHead);
		}

		for (index = 0; index < itemNodes.length; index += 1) {
			if (index + 1 === Number(nPassCtr)) {
				itemAccountHead = accountHeadValue(index + 1);
				if (itemAccountHead === "0") {
					alert("Select Account Head");
					select = accountHeadSelect(index + 1);
					if (select && select.focus) {
						select.focus();
					}
					return;
				}
			}
		}

		for (index = 0; index < itemNodes.length; index += 1) {
			itemAccountHead = accountHeadValue(index + 1);
			if (itemAccountHead !== "0") {
				itemAccountHead = accountHeadNo(itemAccountHead);
				if (accountHeads.indexOf(itemAccountHead) === -1) {
					accountHeads.push(itemAccountHead);
				}
			}
		}

		for (index = 0; index < itemNodes.length; index += 1) {
			if (field("ItemValueZ" + (index + 1))) {
				field("ItemValueZ" + (index + 1)).value = "";
			}
		}

		accountHeads.forEach(function (head, headIndex) {
			var total = 0;
			for (var itemIndex = 0; itemIndex < itemNodes.length; itemIndex += 1) {
				itemAccountHead = accountHeadValue(itemIndex + 1);
				if (itemAccountHead !== "0" && accountHeadNo(itemAccountHead) === head) {
					total += numberValue(itemNodes.item(itemIndex).getAttribute("ItemValue"));
				}
			}
			if (field("ItemValueZ" + (headIndex + 1))) {
				field("ItemValueZ" + (headIndex + 1)).value = formatNumber(total, 2);
			}
			if (total > 0) {
				totalValue += total;
			}
		});

		if (field("mTotalValue")) {
			field("mTotalValue").value = formatNumber(totalValue, 2);
		}
	};

	window.Goto_AccItemValuePage = function () {
		var doc = invoiceDocument();
		var root = invoiceRoot();
		var itemNodes = selectNodes(root, "//ItemDetails/Item");
		var previousAccountHead = "";
		var detailsNode;
		var accountHeadElement;
		var itemAccountHead;
		var index;

		for (index = 0; index < itemNodes.length; index += 1) {
			var selected = accountHeadValue(index + 1);
			var select = accountHeadSelect(index + 1);
			if (selected === "0") {
				alert("Select Account Head");
				if (select && select.focus) {
					select.focus();
				}
				return;
			}
		}

		for (index = 0; index < itemNodes.length; index += 1) {
			var currentAccountHead = accountHeadValue(index + 1);
			if (!previousAccountHead) {
				previousAccountHead = currentAccountHead;
			}
			if (previousAccountHead !== currentAccountHead) {
				alert("All Items Account Head should be same");
				return;
			}
		}

		detailsNode = findVoucherDetails(root);
		accountHeadElement = findRootChild(root, "AccountHead");

		if (detailsNode) {
			childElements(detailsNode).forEach(function (detailNode) {
				var itemCount = trim(detailNode.getAttribute("No"));
				var selectedText = accountHeadValue(itemCount);
				var selectedParts = selectedText.split("|");
				var selectedAccountHead = selectedParts[0] || "";
				var currentHead = firstElement(detailNode);

				if (currentHead && selectedAccountHead !== trim(currentHead.getAttribute("No"))) {
					detailNode.removeChild(currentHead);
					currentHead = createElement("AccHead");
					currentHead.setAttribute("No", selectedAccountHead);
					currentHead.setAttribute("CostCenter", selectedParts[1] || "");
					currentHead.setAttribute("Analytical", selectedParts[2] || "");
					currentHead.setAttribute("Name", detailNode.getAttribute("ItemDesc") || "");
					currentHead.setAttribute("Type", "G");
					currentHead.setAttribute("Group", "");
					detailNode.appendChild(currentHead);
				}
			});
		}

		if (!accountHeadElement) {
			accountHeadElement = createElement("AccountHead");
			root.appendChild(accountHeadElement);
		}

		itemAccountHead = createElement("ItemAccountHead");
		accountHeadElement.appendChild(itemAccountHead);

		for (index = 0; index < itemNodes.length; index += 1) {
			var itemNode = itemNodes.item(index);
			var selectedHead = accountHeadNo(accountHeadValue(index + 1));
			var accDetail = null;
			var previousTotal = 0;

			itemNode.setAttribute("ItemAccHead", selectedHead);

			childElements(itemAccountHead).some(function (node) {
				if (trim(node.getAttribute("No")) === selectedHead) {
					accDetail = node;
					previousTotal = numberValue(node.getAttribute("TotalAmt"));
					return true;
				}
				return false;
			});

			if (!accDetail) {
				accDetail = createElement("Acc");
				accDetail.setAttribute("No", selectedHead);
				accDetail.setAttribute("TotalAmt", itemNode.getAttribute("ItemValue"));
				itemAccountHead.appendChild(accDetail);
			} else {
				accDetail.setAttribute("TotalAmt", previousTotal + numberValue(itemNode.getAttribute("ItemValue")));
			}

			var itemDetail = createElement("Item");
			itemDetail.setAttribute("Desc", itemNode.getAttribute("ItmDescription"));
			itemDetail.setAttribute("Amount", itemNode.getAttribute("ItemValue"));
			itemDetail.setAttribute("ItemCode", itemNode.getAttribute("ItemCode"));
			itemDetail.setAttribute("ClassCode", itemNode.getAttribute("ClassificationCode"));
			accDetail.appendChild(itemDetail);
		}

		var xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSavePur.asp?Mod=PUR&Name=AmdNewInvItemValue", false);
		xhr.setRequestHeader("Content-Type", "text/xml; charset=UTF-8");
		xhr.send(xmlText(doc));

		form().action = "AmdInvPurInvoiceInsert_New.asp";
		form().submit();
	};
})(window, document);
