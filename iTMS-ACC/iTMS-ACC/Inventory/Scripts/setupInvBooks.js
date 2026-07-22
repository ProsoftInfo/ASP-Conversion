(function (window, document) {
	"use strict";

	var bookDoc = null;
	var bookRoot = null;

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
		return frm && frm.elements ? frm.elements[name] || document.getElementById(name) : document.getElementById(name);
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function xmlObject(id) {
		return document.getElementById(id) || window[id] || null;
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

	function parseXml(text) {
		return new DOMParser().parseFromString(text || "<Root/>", "text/xml");
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

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function resolveBookXml() {
		bookDoc = xmlDocument(xmlObject("BookXML")) || document.implementation.createDocument("", "Root", null);
		bookRoot = xmlRoot(xmlObject("BookXML")) || bookDoc.documentElement || bookDoc.appendChild(bookDoc.createElement("Root"));
	}

	function selectedBookNode() {
		var books = elementChildren(bookRoot, "Book");
		return books.length ? books[0] : null;
	}

	function selectBook(no) {
		var select = field("selBooks");
		var wanted = trim(no);
		if (!select || !select.options || !wanted) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value) === wanted) {
				select.selectedIndex = i;
				return;
			}
		}
	}

	function populateBooks(xml) {
		var root = xmlRoot(xml);
		var select = field("selBooks");
		if (!select || !root) {
			return;
		}
		elementChildren(root).forEach(function (node) {
			var option = document.createElement("option");
			option.text = getAttr(node, "BookName");
			option.value = getAttr(node, "BookNumber");
			select.add(option);
		});
	}

	function loadBooks() {
		var org = field("hOrgID") ? field("hOrgID").value : "";
		var xhr = new XMLHttpRequest();
		xhr.open("GET", "../../Accounts/Transaction/XMLGetOrgBook.asp?BkCode=08&orgID=" + encodeURIComponent(org), true);
		xhr.onreadystatechange = function () {
			var saved;
			if (xhr.readyState !== 4) {
				return;
			}
			if (xhr.status >= 200 && xhr.status < 300) {
				populateBooks(xhr.responseXML && xhr.responseXML.documentElement ? xhr.responseXML : parseXml(xhr.responseText));
			}
			resolveBookXml();
			saved = selectedBookNode();
			if (saved) {
				selectBook(getAttr(saved, "No"));
			}
		};
		xhr.send(null);
	}

	function serializeXml(doc) {
		if (window.XMLSerializer) {
			return new XMLSerializer().serializeToString(doc);
		}
		return doc.xml || "";
	}

	function saveBookXml() {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", "SetupInvBookInsert.asp", true);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.onreadystatechange = function () {
			if (xhr.readyState !== 4) {
				return;
			}
			alert(trim(xhr.responseText) || "Book Setup is done");
		};
		xhr.send(serializeXml(bookDoc));
	}

	window.Init = function () {
		ensureCompat();
		loadBooks();
		return false;
	};

	window.CheckSubmit = function () {
		var select = field("selBooks");
		var node;
		if (!select || select.selectedIndex < 0) {
			alert("Select the Book");
			if (select) {
				select.focus();
			}
			return false;
		}
		resolveBookXml();
		node = selectedBookNode();
		if (!node) {
			node = bookDoc.createElement("Book");
			bookRoot.appendChild(node);
		}
		node.setAttribute("No", select.options[select.selectedIndex].value);
		node.setAttribute("Name", select.options[select.selectedIndex].text);
		saveBookXml();
		return false;
	};
}(window, document));
