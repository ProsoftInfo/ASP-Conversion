(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function byId(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function textOf(id) {
		var item = byId(id);
		return trim(item ? item.innerText || item.textContent || "" : "");
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function xmlDocument(value) {
		ensureCompat();
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
		return value.ownerDocument || null;
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name);
	}

	function elementChildren(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (!wanted || String(children[i].nodeName).toLowerCase() === wanted)) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function attrAt(node, index) {
		var item = node && node.attributes && node.attributes.item(index);
		return trim(item ? item.nodeValue || item.value || "" : "");
	}

	function attr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, trim(value));
		}
	}

	function removeChildrenByName(root, nodeName) {
		var nodes = elementChildren(root, nodeName);
		for (var i = 0; i < nodes.length; i += 1) {
			root.removeChild(nodes[i]);
		}
	}

	function install(config) {
		var objTemp = null;
		var root = null;
		var itemCode = "";
		var classCode = "";
		var orgId = "";
		var tareValue = "";

		function setReturnValue() {
			if (!root) {
				root = xmlRoot(objTemp);
			}
			window.returnValue = root;
			window.returnvalue = root;
			if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
				window.ITMSModernCompat.returnModalValue(root);
			}
		}

		function closeWithReturn() {
			setReturnValue();
			window.close();
		}

		function getItemNode() {
			var nodes = elementChildren(root, "ITEM");
			for (var i = 0; i < nodes.length; i += 1) {
				if (attrAt(nodes[i], 0) === trim(itemCode) && attrAt(nodes[i], 1) === trim(classCode)) {
					return nodes[i];
				}
			}
			return null;
		}

		function syncInputsFromItem(formatQuantity) {
			var itemNode = getItemNode();
			var children = elementChildren(itemNode);
			var q;
			var value;
			for (var i = 0; i < children.length; i += 1) {
				q = field("txtQty" + (i + 1));
				if (!q) {
					continue;
				}
				value = attrAt(children[i], 2);
				q.value = formatQuantity ? toNumber(value).toFixed(3) : value;
			}
		}

		function loadLotDetails(recNo, flag) {
			var rec = field("hRec");
			var xhr = new XMLHttpRequest();
			var url = config.lotDetailsUrl + "?iRecNo=" + encodeURIComponent(recNo) + "&iItem=" + encodeURIComponent(itemCode) + "&iClass=" + encodeURIComponent(classCode) + "&sOrgID=" + encodeURIComponent(orgId) + "&sFlag=" + encodeURIComponent(flag) + "&sRecType=" + encodeURIComponent(rec ? rec.value : "");
			var island;
			var dataRoot;
			var rootDoc;

			xhr.open("GET", url, true);
			xhr.onreadystatechange = function () {
				if (xhr.readyState !== 4 || !(xhr.status === 0 || xhr.status >= 200 && xhr.status < 300) || !trim(xhr.responseText)) {
					return;
				}
				island = xmlIsland("XmlData");
				if (island && typeof island.loadXML === "function") {
					island.loadXML(xhr.responseText);
				}
				dataRoot = xmlRoot(island) || xmlRoot(xhr.responseXML);
				rootDoc = xmlDocument(root);
				if (dataRoot && root && rootDoc) {
					root.appendChild(rootDoc.importNode ? rootDoc.importNode(dataRoot, true) : dataRoot.cloneNode(true));
				}
			};
			xhr.send(null);
		}

		function totalEnteredQuantity() {
			var hiCtr = toNumber(field("hiCtr") && field("hiCtr").value);
			var total = 0;
			var q;
			for (var i = 1; i <= hiCtr; i += 1) {
				q = field("txtQty" + i);
				total += toNumber(q && q.value);
			}
			return total;
		}

		function resetInvalidQuantities() {
			var itemNode = getItemNode();
			var stores = elementChildren(itemNode);
			var lots;
			var i;
			var j;
			for (i = 0; i < stores.length; i += 1) {
				setAttr(stores[i], "QTY", 0);
				lots = elementChildren(stores[i]);
				for (j = 0; j < lots.length; j += 1) {
					stores[i].removeChild(lots[j]);
				}
			}
		}

		window.fnInit = function (value) {
			var arrTemp = String(value || "").split(":");
			var itemNode;
			var children;
			var recNo;
			var checkQtyFlag = false;
			var hiCtr = toNumber(field("hiCtr") && field("hiCtr").value);

			itemCode = arrTemp[0] || "";
			classCode = arrTemp[1] || "";
			orgId = arrTemp[2] || "";
			tareValue = arrTemp[4] || "";
			objTemp = window.dialogArguments;
			root = xmlRoot(objTemp);

			if (!hiCtr || !root) {
				return;
			}

			itemNode = getItemNode();
			children = elementChildren(itemNode);
			for (var i = 0; i < children.length; i += 1) {
				if (field("txtQty" + (i + 1))) {
					field("txtQty" + (i + 1)).value = attrAt(children[i], 2);
				}
				checkQtyFlag = toNumber(attrAt(children[i], 2)) > 0;
			}

			recNo = attr(root, "RECNO");
			loadLotDetails(recNo, checkQtyFlag ? "N" : "Y");
		};

		window.GetLot = function (obj, recType, counter, storeName) {
			var arrTemp = String(obj && obj.name || "").split(":");
			var localClass = arrTemp[1] || "";
			var localItem = arrTemp[2] || "";
			var localOrg = arrTemp[3] || "";
			var store = arrTemp[4] || "";
			var bin = arrTemp[5] || "";
			var args = recType + ":" + localItem + ":" + localClass + ":" + localOrg + ":" + store + ":" + bin + ":" + storeName + ":" + tareValue + ":True";

			function refresh() {
				var itemNode = getItemNode();
				var nodes = elementChildren(itemNode);
				var q;
				for (var i = 0; i < nodes.length; i += 1) {
					if (attrAt(nodes[i], 0) === trim(store) && attrAt(nodes[i], 1) === trim(bin)) {
						q = field("txtQty" + counter);
						if (q) {
							q.value = config.formatPopupQuantity ? toNumber(attrAt(nodes[i], 2)).toFixed(3) : attrAt(nodes[i], 2);
						}
					}
				}
			}

			if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
				window.ITMSModernCompat.openModalDialog(config.lotPopup + "?sTemp=" + encodeURIComponent(args), objTemp, "dialogHeight:370px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", refresh);
			} else {
				window.open(config.lotPopup + "?sTemp=" + encodeURIComponent(args), "_blank");
			}
		};

		window.CheckSubmit = function () {
			var hiCtr = toNumber(field("hiCtr") && field("hiCtr").value);
			var qtyTotal = 0;
			var itemNode;
			var children;
			var q;
			var rec = field("hRec");
			var lotNumber = textOf("idLotNumber");
			var i;

			if (!hiCtr) {
				closeWithReturn();
				return;
			}

			for (i = 1; i <= hiCtr; i += 1) {
				q = field("txtQty" + i);
				if (!q || trim(q.value) === "") {
					alert("Enter Quantity");
					if (q) {
						q.select();
					}
					return;
				}
				if (!checkNumbers(q.value)) {
					alert("Enter Numerals Only");
					q.select();
					return;
				}
				qtyTotal += toNumber(q.value);
			}

			if (qtyTotal !== toNumber(textOf("idQty"))) {
				alert("Total Quantity should be equal to Quantity to Account (" + textOf("idQty") + ")");
				return;
			}

			itemNode = getItemNode();
			children = elementChildren(itemNode);
			for (i = 0; i < children.length; i += 1) {
				q = field("txtQty" + (i + 1));
				setAttr(children[i], "QTY", q ? toNumber(q.value) : 0);
				setAttr(children[i], "RECTYPE", rec ? rec.value : "");
				setAttr(children[i], "LOTNUMBER", lotNumber);
			}

			closeWithReturn();
		};

		window.window_onunload = function () {
			if (!root) {
				root = xmlRoot(objTemp);
			}
			if (!root) {
				return;
			}
			if (!toNumber(field("hiCtr") && field("hiCtr").value)) {
				setReturnValue();
				return;
			}
			removeChildrenByName(root, "STOREDLOTDETAILS");
			if (totalEnteredQuantity() !== toNumber(textOf("idQty"))) {
				resetInvalidQuantities();
			}
			setReturnValue();
		};

		window.addEventListener("beforeunload", window.window_onunload);
	}

	function closeWithDialogArguments(message) {
		var root;
		if (message) {
			alert(message);
		}
		root = xmlRoot(window.dialogArguments);
		window.returnValue = root;
		window.returnvalue = root;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		}
		window.close();
	}

	window.ITMSReceiptStore = {
		install: install,
		closeWithDialogArguments: closeWithDialogArguments
	};
}(window, document));
