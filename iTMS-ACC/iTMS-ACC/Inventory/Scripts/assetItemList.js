(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function valueOf(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function setCategoryText(value) {
		var item = document.getElementById("spCatName") || window.spCatName;
		if (item) {
			item.textContent = value || "";
		}
	}

	function openDialog(url, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", features, callback);
		} else {
			window.open(url, "_blank", features);
		}
	}

	function setDate() {
		return false;
	}

	function Validate() {
		setValue("hOrgId", valueOf("selUnitId"));
		form().action = "AssetItmListEntry.asp";
		form().submit();
		return false;
	}

	function ResetData() {
		return false;
	}

	function AssignPage(pageNo) {
		setValue("hPage", pageNo);
		form().action = "AssetItmListEntry.asp";
		form().submit();
		return false;
	}

	function popAsset() {
		var type = valueOf("selAssetType");
		var orgId = valueOf("hOrgId");
		setCategoryText("");
		if (type === "C") {
			openDialog("../../FixedAssets/Transaction/PopCategoryList.asp", "dialogHeight:420px;dialogWidth:285px;center:Yes;help:No;resizable:No;status:No", function (groupCode) {
				var parts;
				if (trim(groupCode) && trim(groupCode) !== "0") {
					parts = String(groupCode).split(":");
					setValue("hCategory", parts[0] || "");
					setCategoryText(parts[1] || "");
				}
			});
		} else if (type === "A") {
			openDialog("../../FixedAssets/Transaction/PopCategoryList.asp", "dialogHeight:420px;dialogWidth:285px;center:Yes;help:No;resizable:No;status:No", function (groupCode) {
				var parts;
				if (!trim(groupCode) || trim(groupCode) === "0") {
					return;
				}
				parts = String(groupCode).split(":");
				setCategoryText(parts[1] || "");
				openDialog("../../FixedAssets/Transaction/AssetSelMethodNEW.asp?CatCode=" + encodeURIComponent(parts[0] || "") + "&orgid=" + encodeURIComponent(orgId), "dialogHeight:420px;dialogWidth:470px;center:Yes;help:No;resizable:No;status:No", function (assetCodes) {
					if (trim(assetCodes).length >= 1) {
						setValue("hAssetCode", assetCodes);
					}
				});
			});
		}
		return false;
	}

	function CheckSubmit() {
		var count = Number(valueOf("hCnt")) || 0;
		var selectedCount = 0;
		var selectedAsset = "";
		var box;
		for (var i = 1; i <= count; i += 1) {
			box = field("Chkbox" + i);
			if (box && box.checked) {
				if (trim(box.value) !== "") {
					selectedAsset = box.value;
				}
				selectedCount += 1;
			}
		}
		if (selectedCount === 0 || selectedCount > 1) {
			alert("Select any one Asset");
			return false;
		}
		form().action = "AssetItmCreationEntry.asp?sTemp=" + encodeURIComponent(selectedAsset + ":INV");
		form().submit();
		return false;
	}

	window.setDate = setDate;
	window.Validate = Validate;
	window.ResetData = ResetData;
	window.AssignPage = AssignPage;
	window.popAsset = popAsset;
	window.CheckSubmit = CheckSubmit;
}(window, document));
