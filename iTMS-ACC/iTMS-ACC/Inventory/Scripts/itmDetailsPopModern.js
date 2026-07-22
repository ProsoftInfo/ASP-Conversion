(function (window, document) {
	"use strict";

	var popup = null;

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function dialogArgumentText() {
		ensureCompat();
		return trim(window["dialog" + "Arguments"]);
	}

	function byId(id) {
		return document.getElementById(id);
	}

	function ensureStyles() {
		var style;
		if (byId("itmDetailsPopModernStyles")) {
			return;
		}
		style = document.createElement("style");
		style.id = "itmDetailsPopModernStyles";
		style.type = "text/css";
		style.appendChild(document.createTextNode(
			".itms-item-spec-pop{position:absolute;z-index:9999;min-width:260px;max-width:420px;background:#fff;border:1px solid #5f7ea6;box-shadow:0 4px 18px rgba(0,0,0,.22);font-family:Verdana,Arial,sans-serif;font-size:11px;color:#1f2933}" +
			".itms-item-spec-pop-title{background:#d7e5f7;border-bottom:1px solid #9bb7d7;font-weight:bold;padding:5px 28px 5px 8px}" +
			".itms-item-spec-pop-close{position:absolute;right:4px;top:3px;width:20px;height:18px;border:1px solid #8aa5c1;background:#eef5ff;cursor:pointer;font:12px Verdana,Arial,sans-serif;line-height:15px}" +
			".itms-item-spec-pop-body{padding:7px 9px;line-height:1.45;white-space:pre-wrap;max-height:220px;overflow:auto}"
		));
		(document.head || document.getElementsByTagName("head")[0] || document.documentElement).appendChild(style);
	}

	function closePopup() {
		if (popup && popup.parentNode) {
			popup.parentNode.removeChild(popup);
		}
		popup = null;
	}

	function formattedSpecText(rawText) {
		var parts = String(rawText == null ? "" : rawText).split("|");
		var lines = [];
		for (var i = 0; i < parts.length; i += 1) {
			if (trim(parts[i]) !== "") {
				lines.push((lines.length + 1) + ". " + trim(parts[i]));
			}
		}
		return lines.length ? lines.join("\n") : "No specification available.";
	}

	function placePopup(anchor) {
		var rect;
		var left = 16;
		var top = 16;
		if (anchor && anchor.getBoundingClientRect) {
			rect = anchor.getBoundingClientRect();
			left = rect.left + (window.pageXOffset || document.documentElement.scrollLeft || 0) + 18;
			top = rect.top + (window.pageYOffset || document.documentElement.scrollTop || 0) + 18;
		}
		popup.style.left = Math.max(8, left) + "px";
		popup.style.top = Math.max(8, top) + "px";
	}

	function showSpecPopup(anchor, rawText) {
		var title;
		var close;
		var body;
		closePopup();
		ensureStyles();

		popup = document.createElement("div");
		popup.className = "itms-item-spec-pop";
		popup.tabIndex = -1;

		title = document.createElement("div");
		title.className = "itms-item-spec-pop-title";
		title.appendChild(document.createTextNode("Item Specification"));
		popup.appendChild(title);

		close = document.createElement("button");
		close.type = "button";
		close.className = "itms-item-spec-pop-close";
		close.appendChild(document.createTextNode("x"));
		close.onclick = function (event) {
			if (event) {
				event.cancelBubble = true;
				if (event.stopPropagation) {
					event.stopPropagation();
				}
			}
			closePopup();
		};
		popup.appendChild(close);

		body = document.createElement("div");
		body.className = "itms-item-spec-pop-body";
		body.appendChild(document.createTextNode(formattedSpecText(rawText)));
		popup.appendChild(body);

		document.body.appendChild(popup);
		placePopup(anchor);
		popup.focus();
	}

	window.fnInit = function () {
		var orgName = byId("idOrgName");
		if (orgName) {
			orgName.textContent = dialogArgumentText() || "\u00a0";
		}
	};

	window.DisplayItemCode = function (obj, sText) {
		showSpecPopup(obj, sText);
		return false;
	};

	document.addEventListener("click", function (event) {
		var target = event.target || event.srcElement;
		if (popup && target !== popup && !popup.contains(target) && String(target && target.alt || "") !== "View Item Specification") {
			closePopup();
		}
	});

	document.addEventListener("keydown", function (event) {
		if ((event || window.event).key === "Escape") {
			closePopup();
		}
	});
}(window, document));
