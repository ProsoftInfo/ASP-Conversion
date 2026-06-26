(function () {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function findElement(name) {
		return document.getElementById(name) ||
			(document.forms[0] && document.forms[0].elements[name]) ||
			document.getElementsByName(name)[0] ||
			null;
	}

	function applyLegacyOptions(element, options) {
		var html;
		var parsed;
		var attrs;
		if (!trim(options)) {
			return;
		}
		html = "<input " + options + ">";
		parsed = new DOMParser().parseFromString(html, "text/html").body.firstElementChild;
		if (!parsed) {
			return;
		}
		attrs = parsed.attributes;
		for (var i = 0; i < attrs.length; i += 1) {
			element.setAttribute(attrs[i].name, attrs[i].value);
		}
	}

	function createInput(type, name, value, size, maxLength, className, options) {
		var input = document.createElement("input");
		input.type = type;
		input.name = name || "";
		input.value = value == null ? "" : String(value);
		if (size) {
			input.size = Number(size);
		}
		if (maxLength) {
			input.maxLength = Number(maxLength);
		}
		if (className) {
			input.className = className;
		}
		applyLegacyOptions(input, options);
		return input;
	}

	window.InsertCell = function (oRow, iType, sName, sValue, sClass, sAlign, sValign, iSize, iMaxlen, iColspan, iRowspan, sOptions) {
		var objCell = oRow.insertCell();
		var cellType = Number(iType);

		if (cellType === 1) {
			objCell.innerHTML = sValue == null ? "" : String(sValue);
		} else if (cellType === 2) {
			objCell.appendChild(createInput("text", sName, sValue, iSize, iMaxlen, "Formelem", sOptions));
		} else if (cellType === 3) {
			objCell.appendChild(createInput("checkbox", sName, sValue, 0, 0, "", sOptions));
		} else if (cellType === 4) {
			objCell.appendChild(createInput("text", sName, sValue, iSize, iMaxlen, "FormelemRead", sOptions));
		}

		objCell.className = sClass || "";
		if (trim(sAlign)) {
			objCell.align = sAlign;
		}
		if (trim(sValign)) {
			objCell.vAlign = sValign;
			objCell.setAttribute("valign", sValign);
		}
		if (Number(iColspan) !== 0) {
			objCell.colSpan = Number(iColspan);
		}
		if (Number(iRowspan) !== 0) {
			objCell.rowSpan = Number(iRowspan);
		}
		return objCell;
	};

	window.ClearTable = window.ClearTable || function (objTable, startlen, Count) {
		var table = typeof objTable === "string" ? findElement(objTable) : objTable;
		var start = Number(startlen) || 0;
		var keep = Number(Count) || 0;
		if (!table || !table.rows) {
			return;
		}
		while (table.rows.length > start + keep) {
			table.deleteRow(start);
		}
	};
}());
