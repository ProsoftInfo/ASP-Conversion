(function () {
	"use strict";

	function parseNumber(value) {
		var parsed = Number(value);
		return isNaN(parsed) ? 0 : parsed;
	}

	function readConversion(xmlDoc) {
		var nodes;
		var first;
		if (!xmlDoc) {
			return { rate: 0, operator: 0 };
		}
		nodes = xmlDoc.getElementsByTagName("*");
		first = nodes.length ? nodes[0] : null;
		return {
			rate: parseNumber(first && first.getAttribute("OptionToBaseRate")),
			operator: parseInt(first && first.getAttribute("OptionToBaseOperator"), 10) || 0
		};
	}

	window.getRatePerQtyUoM = function (ORGID, CLASSCODE, ITEMCODE, QTYUOM, RATEUOM, RATE) {
		var xhr = new XMLHttpRequest();
		var url = "XMLgetUoMConvFactor.asp?ORGID=" + encodeURIComponent(ORGID) +
			"&CLASSCODE=" + encodeURIComponent(CLASSCODE) +
			"&ITEMCODE=" + encodeURIComponent(ITEMCODE) +
			"&QTYUOM=" + encodeURIComponent(QTYUOM) +
			"&RATEUOM=" + encodeURIComponent(RATEUOM) +
			"&FLAG=P";
		var conversion;

		xhr.open("GET", url, false);
		xhr.send(null);
		conversion = readConversion(xhr.responseXML);

		if (conversion.operator === 1) {
			return conversion.rate ? parseNumber(RATE) / conversion.rate : 0;
		}
		return parseNumber(RATE) * conversion.rate;
	};
}());
