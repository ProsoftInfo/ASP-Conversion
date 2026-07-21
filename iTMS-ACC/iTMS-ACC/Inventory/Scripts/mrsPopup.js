(function (window, document) {
	"use strict";

	function field(name) {
		var form = document.forms.formname || document.forms[0];
		return form && form.elements ? form.elements[name] : null;
	}

	function byId(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function textOf(item) {
		return item ? item.innerText || item.textContent || "" : "";
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		return window.open(url, "_blank");
	}

	window.DisplayDet = function (sText) {
		var arrTemp = String(sText || "").split("|");
		alert("To Purchase Requisition : " + (arrTemp[0] || "") + "\nFor Stock Transfer         : " + (arrTemp[1] || ""));
	};

	window.CheckQty = function (obj) {
		var arrTemp = String(obj && obj.name || "").split(":");
		var sClass = arrTemp[1] || "";
		var sItem = arrTemp[2] || "";
		var iEntNo = arrTemp[3] || "";
		var sOptName = arrTemp[4] || "";
		var hMRSNo = field("hMRSNo");
		var sTempValues = sItem + ":" + sClass + ":" + (hMRSNo ? hMRSNo.value : "") + ":" + iEntNo + ":" + sOptName;
		openDialog("mrsIssueQtyParaPoP.asp?sTemp=" + encodeURIComponent(sTempValues), window.OutData, "dialogHeight:385px;dialogWidth:350px;center:Yes;help:No;resizable:No;status:No");
	};

	window.DisplayItem = function (obj) {
		openDialog("itmDetailsPop.asp?sTemp=" + encodeURIComponent(obj || ""), textOf(byId("idOrgName")), "dialogHeight:360px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No");
	};

	window.CheckSch = function (obj, i) {
		var arrTemp = String(obj && obj.name || "").split(":");
		var sClass = arrTemp[1] || "";
		var sItem = arrTemp[2] || "";
		var iCounter = arrTemp[3] || "";
		var sOptName = arrTemp[4] || "";
		var qty = field("txtQtyIssue" + i);
		var hMRSNo = field("hMRSNo");
		var hOrgID = field("hOrgID");
		var sTempValues = (qty ? qty.value : "") + ":" + sItem + ":" + sClass + ":" + (hMRSNo ? hMRSNo.value : "") + ":" + (hOrgID ? hOrgID.value : "") + ":" + iCounter + ":" + sOptName;
		openDialog("mrsIssueSchedulePoP.asp?sTemp=" + encodeURIComponent(sTempValues), window.Data, "dialogHeight:480px;dialogWidth:390px;center:Yes;help:No;resizable:No;status:No");
	};
}(window, document));
