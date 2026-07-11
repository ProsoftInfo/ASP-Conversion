(function (window, document) {
	"use strict";

	var saved = false;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || frm.elements[String(name).toLowerCase()] || null : null;
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

	function checkbox(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || frm.elements[name.charAt(0).toLowerCase() + name.slice(1)] || null : null;
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
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

	function serializeXml(nodeOrDoc) {
		var doc = xmlDocument(nodeOrDoc);
		var root = xmlRoot(nodeOrDoc);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function childElements(node, nodeName) {
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function descendantsByName(node, nodeName) {
		var wanted = String(nodeName || "").toLowerCase();
		var result = [];
		childElements(node).forEach(function (child) {
			if (String(child.nodeName).toLowerCase() === wanted) {
				result.push(child);
			}
			result = result.concat(descendantsByName(child, nodeName));
		});
		result.Item = function (index) {
			return this[index];
		};
		result.item = result.Item;
		return result;
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
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

	function roundNumber(value) {
		return Math.round(toNumber(value));
	}

	function tdsNodes(root) {
		return descendantsByName(root, "TDS");
	}

	function tempRoot() {
		return xmlRoot("TempData");
	}

	function dialogRoot() {
		return xmlRoot(window.dialogArguments) || window.dialogArguments && window.dialogArguments.nodeType === 1 && window.dialogArguments || null;
	}

	function postTempData() {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", "XMLSaveForTDS.asp?Name=TDS_Cash", false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(serializeXml("TempData"));
		return xhr.responseText || "";
	}

	function updateRoundField(counter) {
		var check = checkbox("ChkRnd" + counter);
		var target = field("value" + counter);
		if (check && check.checked) {
			setValue("hRndOffVal" + counter, "Y");
			if (target) {
				target.value = formatNumber(roundNumber(target.value));
			}
		} else {
			setValue("hRndOffVal" + counter, "N");
			if (target) {
				target.value = formatNumber(target.value);
			}
		}
	}

	function formulaPercentage(formulaPart) {
		var pieces = String(formulaPart || "").split("-");
		return pieces.length > 1 ? toNumber(pieces[1]) : 0;
	}

	function formulaGroup(formulaPart) {
		return String(formulaPart || "").split("#")[0];
	}

	function formulaCounter(formulaPart) {
		var parts = String(formulaPart || "").split("#");
		return parts.length > 1 ? String(parts[1]).charAt(0) : "";
	}

	function CalculatePer(root) {
		var nodes = childElements(root);
		var counter = 1;
		var baseTdsAmount = 0;

		nodes.forEach(function (node) {
			var currentCounter;
			var formula;
			var parts;
			var totalAmount;
			var addFlag;

			if (trim(attr(node, "Ctr")) === trim(counter)) {
				currentCounter = attr(node, "Ctr");
				formula = trim(attr(node, "Formula"));
				if (Number(currentCounter) === 1) {
					baseTdsAmount = toNumber(attr(node, "TDSAmount"));
				}
				totalAmount = 0;
				if (formula !== "" && trim(attr(node, "CompMode")) === "P") {
					parts = formula.split(",");
					addFlag = parts.length > 1;
					if (trim(formulaGroup(parts.length > 1 ? parts[0] : formula)) === "0") {
						totalAmount = baseTdsAmount * (formulaPercentage(parts[0]) / 100);
						setValue("value" + currentCounter, formatNumber(totalAmount));
						setAttr(node, "PayRecAmount", totalAmount);
					}
					parts.forEach(function (part) {
						var groupCounter = formulaCounter(part);
						if (!groupCounter) {
							return;
						}
						childElements(root).some(function (sourceNode) {
							var sourceValue;
							if (trim(groupCounter) !== trim(attr(sourceNode, "Ctr"))) {
								return false;
							}
							sourceValue = toNumber(attr(sourceNode, "PayRecAmount"));
							if (addFlag) {
								totalAmount += sourceValue * (formulaPercentage(part) / 100);
							} else {
								totalAmount = sourceValue * (formulaPercentage(part) / 100);
							}
							setValue("value" + currentCounter, formatNumber(totalAmount));
							setAttr(node, "PayRecAmount", totalAmount);
							return Number(groupCounter) === 2;
						});
					});
				}
			}
			counter += 1;
			updateRoundField(attr(node, "Ctr"));
		});
		postTempData();
		return true;
	}

	function rewriteFormula(formula, percentage, compMode) {
		if (trim(compMode) === "F") {
			return String(parseInt(toNumber(percentage), 10) || 0);
		}
		return String(formula || "").split(",").map(function (part) {
			return String(part).split("-")[0] + "-" + percentage;
		}).join(",");
	}

	function TDSCalc(counter) {
		var changed = field("txtPer" + counter);
		var root = tempRoot();
		var nodes = tdsNodes(root);
		var count = Number(valueOf("htxtPer", "0"));
		var index;
		var node;
		var percentage;

		if (!changed || !isNumeric(changed.value)) {
			alert("Enter Numbers Only");
			return false;
		}
		setValue("hReCalc", "Y");
		setValue("hPer", changed.value);
		setValue("hCtr", counter);

		for (index = 1; index <= count; index += 1) {
			node = nodes[index - 1];
			if (!node || Number(attr(node, "Ctr")) !== index) {
				continue;
			}
			percentage = valueOf("txtPer" + index, "0");
			setAttr(node, "TdsPercentage", percentage);
			setAttr(node, "Formula", rewriteFormula(attr(node, "Formula"), percentage, attr(node, "CompMode")));
			setAttr(node, "CompMode", attr(node, "CompMode"));
			setAttr(node, "TdsRndOff", checkbox("ChkRnd" + index) && checkbox("ChkRnd" + index).checked ? "Y" : "N");
		}
		return CalculatePer(root);
	}

	function DisplayTot() {
		var count = Number(valueOf("htxtPer", "0"));
		var total = 0;
		var index;
		for (index = 1; index <= count; index += 1) {
			total += toNumber(valueOf("value" + index, "0"));
		}
		setValue("txtTdsTotal", formatNumber(total));
		return true;
	}

	function SetRoundOff(counter) {
		var check = checkbox("ChkRnd" + counter);
		if (check && check.checked) {
			setValue("hRndOffVal" + counter, "Y");
			setValue("value" + counter, formatNumber(roundNumber(valueOf("value" + counter, "0"))));
		}
		return TDSCalc(counter);
	}

	function SaveXml() {
		var root = tempRoot();
		var nodes = tdsNodes(root);
		var counter = 1;

		nodes.forEach(function (node) {
			if (trim(attr(node, "Ctr")) === trim(counter)) {
				setAttr(node, "TdsPercentage", formatNumber(valueOf("txtPer" + counter, "0")));
				setAttr(node, "PayRecAmount", formatNumber(valueOf("value" + counter, "0")));
				setAttr(node, "TdsRndOff", valueOf("hRndOffVal" + counter, "N"));
				counter += 1;
			}
		});
		saved = true;
		postTempData();
		return true;
	}

	function dialogId() {
		var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		return match ? decodeURIComponent(match[1]) : "";
	}

	function notifyDialogValue(id, value) {
		if (!id || !window.opener) {
			return;
		}
		try {
			if (window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
				window.opener.ITMSModernCompat._receiveDialogValue(id, value);
				return;
			}
		} catch (ignoreDirectReturn) {}
		try {
			window.opener.postMessage({ type: "itms-dialog-return", id: id, value: value }, window.location.origin || "*");
		} catch (ignoreMessageReturn) {}
	}

	function returnValue(value) {
		var id;
		window.returnValue = value;
		window.returnvalue = value;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(value);
			return;
		}
		id = dialogId();
		notifyDialogValue(id, value);
	}

	function Window_onunload() {
		var root = tempRoot();
		SaveXml();
		returnValue(root);
		window.close();
		return false;
	}

	function hasEntryForGroup(root, groupId) {
		return descendantsByName(root, "Entry").some(function (node) {
			return trim(attr(node, "GroupId")) === trim(groupId);
		});
	}

	function sourceTdsNodes() {
		var outRoot = xmlRoot("OutData");
		var argsRoot = dialogRoot();
		var currentRoot = outRoot || argsRoot || tempRoot();
		var newValue = trim(valueOf("hNewVal"));
		var groupId = valueOf("hNewGrpId");

		if (newValue !== "Y") {
			return hasEntryForGroup(currentRoot, groupId) ? tdsNodes(currentRoot) : tdsNodes(argsRoot);
		}
		return tdsNodes(argsRoot);
	}

	function Init() {
		var nodes = sourceTdsNodes();
		nodes.forEach(function (node, index) {
			var counter = index + 1;
			var roundOff = trim(attr(node, "TdsRndOff")).toUpperCase();
			setValue("txtPer" + counter, formatNumber(attr(node, "TdsPercentage")));
			if (roundOff === "Y") {
				if (checkbox("ChkRnd" + counter)) {
					checkbox("ChkRnd" + counter).checked = true;
				}
				setValue("hRndOffVal" + counter, "Y");
				setValue("value" + counter, formatNumber(roundNumber(attr(node, "PayRecAmount"))));
			} else if (roundOff === "N") {
				if (checkbox("ChkRnd" + counter)) {
					checkbox("ChkRnd" + counter).checked = false;
				}
				setValue("hRndOffVal" + counter, "N");
				setValue("value" + counter, formatNumber(attr(node, "PayRecAmount")));
			}
		});
		DisplayTot();
		return true;
	}

	window.TDSCalc = TDSCalc;
	window.TDScalc = TDSCalc;
	window.CalculatePer = CalculatePer;
	window.DisplayTot = DisplayTot;
	window.SetRoundOff = SetRoundOff;
	window.SaveXml = SaveXml;
	window.Window_onunload = Window_onunload;
	window.window_onunload = Window_onunload;
	window.Init = Init;

	window.addEventListener("beforeunload", function () {
		if (!saved && window.returnValue === undefined && window.returnvalue === undefined) {
			SaveXml();
			returnValue(tempRoot());
		}
	});
}(window, document));
