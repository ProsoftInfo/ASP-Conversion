(function (window, document) {
	"use strict";

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value;
		}
	}

	function submitTo(action) {
		var frm = form();
		if (frm) {
			frm.action = action;
			frm.submit();
		}
	}

	function postText(url, callback) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, true);
		xhr.onreadystatechange = function () {
			if (xhr.readyState === 4) {
				callback(xhr.responseText || "");
			}
		};
		xhr.send(null);
	}

	function byId(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function replaceOptions(select, entries) {
		if (!select) {
			return;
		}
		select.options.length = 1;
		for (var i = 0; i < entries.length; i += 1) {
			select.options[select.options.length] = new Option(entries[i].text, entries[i].value);
		}
	}

	function updateActionSelect(prefix, index, actionValue) {
		var select = field(prefix + "Z" + index);
		var linkId = prefix === "selAmend" ? "idAMHref" + index : "idAPHref" + index;

		if (actionValue === "C") {
			if (select) {
				select.disabled = true;
			}
			if (byId(linkId)) {
				byId(linkId).href = "#";
			}
		} else if (actionValue === "O") {
			replaceOptions(select, [
				{ text: "Un Hold", value: "U" },
				{ text: "Cancel", value: "C" }
			]);
		} else if (actionValue === "U") {
			replaceOptions(select, [
				{ text: "On Hold", value: "O" },
				{ text: "Cancel", value: "C" }
			]);
		}
	}

	function statusAction(obj, which, targetPrefix, selectPrefix, hrefPrefix) {
		var arrTemp;
		var target;
		var url;

		if (!obj || obj.selectedIndex === 0) {
			return;
		}
		if (!confirm("Do you want the MRS to be " + selectedText(obj))) {
			obj.selectedIndex = 0;
			return;
		}

		setValue("hSelected", obj.name);
		setValue("hWhichMRS", which);
		arrTemp = String(obj.name || "").split("Z");
		url = "mrsMgmtInsert.asp?hSelected=" + encodeURIComponent(obj.name || "") + "&hWhichMRS=" + encodeURIComponent(which) + "&sAction=" + encodeURIComponent(obj.value || "");

		postText(url, function (responseText) {
			if (responseText.substring(0, 3) === "MRS") {
				target = byId(targetPrefix + arrTemp[1]);
				if (target) {
					target.innerHTML = responseText;
				}
				updateActionSelect(selectPrefix, arrTemp[1], obj.value);
				return;
			}
			alert(responseText);
		});
	}

	window.RequisitionAction = function (obj) {
		var arrTemp;
		if (obj && obj.selectedIndex !== 0) {
			arrTemp = String(obj.value || "").split("?");
			setValue("mrs", arrTemp[0] || "");
			setValue("sAct", arrTemp[1] || "");
			setValue("hAction", selectedText(obj));
			submitTo("MRApprovalEntry.asp");
		}
	};

	window.IssueAction = function (obj) {
		var arrTemp;
		if (obj && obj.selectedIndex !== 0) {
			arrTemp = String(obj.value || "").split("?");
			setValue("mrs", arrTemp[0] || "");
			setValue("sAct", arrTemp[1] || "");
			setValue("hAction", selectedText(obj));
			submitTo("mrsIssueItemEntry.asp");
		}
	};

	window.AmendAction = function (obj) {
		statusAction(obj, "AM", "idAmend", "selAmend", "idAMHref");
	};

	window.ApproveAction = function (obj) {
		statusAction(obj, "AP", "idApprove", "selApprove", "idAPHref");
	};
}(window, document));
