(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function isNumeric(value) {
		return trim(value) !== "" && !isNaN(Number(String(value).replace(/,/g, "")));
	}

	function formatNumber(value) {
		return toNumber(value).toFixed(2);
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && (frm.elements[name] || frm[name]) || null;
	}

	function xmlObject(name) {
		var element;
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
		element = document.getElementById(name);
		return window[name] || document[name] || element && element._itmsXmlIsland || element || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		return object && (object.XMLDocument || object._doc || object) || null;
	}

	function xmlRoot(name) {
		var doc = xmlDocument(name);
		return doc && doc.documentElement || null;
	}

	function childElements(node, nodeName) {
		var result = [];
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var nodes = [];
		if (!context) {
			return nodes;
		}
		if (typeof context.selectNodes === "function") {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return nodes;
		}
		found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (var i = 0; i < found.snapshotLength; i += 1) {
			nodes.push(found.snapshotItem(i));
		}
		return nodes;
	}

	function attr(node, name) {
		var attribute = node && node.attributes && node.attributes.getNamedItem(name);
		return attribute ? attribute.value : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function attrByIndex(node, index) {
		var attribute = node && node.attributes && node.attributes.item(index);
		return attribute ? attribute.nodeValue : "";
	}

	function setAttrByIndex(node, index, value) {
		var attribute = node && node.attributes && node.attributes.item(index);
		if (attribute) {
			attribute.nodeValue = value == null ? "" : String(value);
		}
	}

	function firstChild(node, nodeName) {
		var matches = childElements(node, nodeName);
		return matches.length ? matches[0] : null;
	}

	function baseAdvanceSuffix(advance, options) {
		var docNo = attrByIndex(advance, 0);
		var advNo = attr(advance, options.advanceNoAttr || "AdvNo");
		return docNo + "Z" + advNo;
	}

	function advanceSuffix(advance, index, options) {
		var suffix = baseAdvanceSuffix(advance, options);
		if (options.rowSuffix) {
			return suffix + "Z" + (index + 1);
		}
		return suffix;
	}

	function advanceField(prefix, advance, index, options) {
		var item = field(prefix + advanceSuffix(advance, index, options));
		if (!item && options.rowSuffix) {
			item = field(prefix + baseAdvanceSuffix(advance, options));
		}
		return item;
	}

	function postXml(url, doc) {
		var request = new XMLHttpRequest();
		request.open("POST", url, false);
		request.send(doc);
		return request.responseText || "";
	}

	function validateAdvanceAmount(advance, amount, options) {
		var toAccount = attrByIndex(advance, 9);
		var available;
		if (options.zeroBlankToAccount && trim(toAccount) === "") {
			toAccount = "0";
			setAttrByIndex(advance, 9, "0");
		}
		available = toNumber(attrByIndex(advance, 3)) - toNumber(attrByIndex(advance, 4));
		if (options.subtractToAccount) {
			available -= toNumber(toAccount);
		}
		if (available < toNumber(amount)) {
			alert(options.availableMessage || "To be Adjusted Amount is Greater Than available Amount");
			return false;
		}
		return true;
	}

	function collectSelectedAdvances(advRoot, options) {
		var advances = childElements(advRoot);
		var total = 0;
		for (var i = 0; i < advances.length; i += 1) {
			var check = advanceField("chkDocument", advances[i], i, options);
			var amountField;
			var amount;
			if (!check || !check.checked) {
				continue;
			}
			amountField = advanceField("txtAmount", advances[i], i, options);
			amount = amountField ? amountField.value : "";
			if (trim(amount) === "") {
				continue;
			}
			if (!isNumeric(amount)) {
				alert("Enter Numeric Value");
				if (amountField) {
					amountField.focus();
				}
				return { ok: false, total: total };
			}
			if (!validateAdvanceAmount(advances[i], amount, options)) {
				return { ok: false, total: total };
			}
			setAttrByIndex(advances[i], 5, amount);
			total += toNumber(amount);
		}
		return { ok: true, total: total, hadRows: advances.length > 0 };
	}

	function collectMiscAdvances(miscRoot) {
		var total = 0;
		childElements(miscRoot).forEach(function (node) {
			var miscNo = attr(node, "MiscNo");
			var check = field("chkDocumentZ" + miscNo);
			var amount;
			if (check && check.checked) {
				amount = attr(node, "Amount");
				setAttr(node, "TobeAdjAmount", amount);
				total += toNumber(amount);
			}
		});
		return total;
	}

	function collectCommissions(commRoot, options) {
		var total = 0;
		var agents = childElements(commRoot);
		for (var i = 0; i < agents.length; i += 1) {
			var agentCode = attrByIndex(agents[i], 0);
			var amountField = field("txtCommAmount" + agentCode);
			var amount = amountField ? amountField.value : "";
			if (trim(amount) === "") {
				continue;
			}
			if (!isNumeric(amount)) {
				alert("Enter Numeric Value");
				if (amountField) {
					amountField.focus();
				}
				return { ok: false, total: total };
			}
			setAttrByIndex(agents[i], 4, amount);
			total += toNumber(amount);
		}
		if (toNumber(options.nettAmount) < total) {
			alert("Total Commission Amount Cannot be Greater than Invoice Total");
			return { ok: false, total: total };
		}
		return { ok: true, total: total };
	}

	function confirmNoAdjustment(options) {
		if (!options.confirmNoAdjustment) {
			return true;
		}
		return window.confirm(options.confirmNoAdjustmentMessage || "Continue! Without Adjusting Advances?");
	}

	function actionDone(options) {
		var frm = form();
		var root = xmlRoot("AdvanceData");
		var advRoot = firstChild(root, "AdvanceDetails");
		var miscRoot = firstChild(root, "MiscAdvanceDetails");
		var commRoot = firstChild(root, "AgentDetails");
		var result = { ok: true, total: 0, hadRows: false };
		var commissionResult;
		var responseText;
		options = options || {};

		if (advRoot) {
			result = collectSelectedAdvances(advRoot, options);
			if (!result.ok) {
				return false;
			}
		}

		if (result.hadRows && result.total === 0 && !confirmNoAdjustment(options)) {
			return false;
		}

		if (options.includeMisc && miscRoot) {
			result.total += collectMiscAdvances(miscRoot);
		}

		if (toNumber(options.invoiceAmount) < result.total) {
			alert(options.invoiceMessage || "To be Adjusted Amount is Greater Than Invoice Amount");
			return false;
		}

		if (options.includeCommission && commRoot) {
			commissionResult = collectCommissions(commRoot, options);
			if (!commissionResult.ok) {
				return false;
			}
		}

		responseText = postXml(options.saveUrl, xmlDocument("AdvanceData"));
		if (trim(responseText) !== "") {
			alert(responseText);
			return false;
		}

		if (options.disableButton && field(options.disableButton)) {
			field(options.disableButton).disabled = true;
		}
		if (frm && options.emptyFieldAction && field(options.emptyFieldAction.fieldName)) {
			frm.action = trim(field(options.emptyFieldAction.fieldName).value) === "" ? options.emptyFieldAction.emptyAction : options.emptyFieldAction.valueAction;
		}
		if (frm) {
			frm.submit();
		}
		return true;
	}

	function setAmount(options) {
		var root = xmlRoot("AdvanceData");
		var total = toNumber(field("hInvVal") && field("hInvVal").value);
		options = options || {};
		selectNodes(root, "//AdvanceDetails/Advance").forEach(function (advance, index) {
			var amountField = advanceField("txtAmount", advance, index, options);
			var check = advanceField("chkDocument", advance, index, options);
			var amountAdjust = toNumber(attr(advance, "AmountRec")) - toNumber(attr(advance, "AmountAdj")) - toNumber(attr(advance, "ToAccount"));
			var value = amountAdjust > total ? total : amountAdjust;
			if (amountField) {
				amountField.value = formatNumber(value);
			}
			if (value !== 0 && check) {
				check.checked = true;
			}
			total = amountAdjust > total ? 0 : total - amountAdjust;
		});
	}

	function setCommission(objComRate, agentId, invoiceValue) {
		var root = xmlRoot("AdvanceData");
		var commRoot = firstChild(root, "AgentDetails");
		var amount = objComRate ? objComRate.value : "";
		var amountField = field("txtCommAmount" + agentId);
		var agents;
		var agent = null;
		var commissionType;
		var commissionValue;
		if (trim(amount) !== "" && !isNumeric(amount)) {
			alert("Enter Numeric Value");
			if (objComRate && objComRate.focus) {
				objComRate.focus();
			}
			return false;
		}
		if (trim(amount) === "") {
			if (amountField) {
				amountField.value = "";
			}
			return false;
		}
		agents = childElements(commRoot, "Agent");
		for (var i = 0; i < agents.length; i += 1) {
			if (attr(agents[i], "Agentcode") === String(agentId)) {
				agent = agents[i];
				break;
			}
		}
		if (!agent) {
			return false;
		}
		setAttr(agent, "Commision", amount);
		commissionType = attr(agent, "Commisiontype");
		if (commissionType === "Q") {
			commissionValue = toNumber(invoiceValue) * toNumber(amount);
		} else {
			commissionValue = (toNumber(invoiceValue) * toNumber(amount)) / 100;
		}
		setAttr(agent, "CommValue", commissionValue);
		if (amountField) {
			amountField.value = commissionValue;
		}
		return true;
	}

	function install(options) {
		options = options || {};
		window.actionDone = function () {
			return actionDone(options);
		};
		window.SetAmount = function () {
			return setAmount(options);
		};
		window.setCommision = setCommission;
		window.finalcancel = function () {};
	}

	window.ITMSAdvanceAdjustmentCompat = {
		install: install
	};
}(window, document));
