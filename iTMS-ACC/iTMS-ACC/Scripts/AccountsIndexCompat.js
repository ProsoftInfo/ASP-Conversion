(function (window, document) {
	"use strict";

	var menuOpen = true;

	function byId(id) {
		return document.getElementById(id);
	}

	function tableCell(table, rowIndex, cellIndex) {
		return table && table.rows[rowIndex] ? table.rows[rowIndex].cells[cellIndex] : null;
	}

	function setBodyWidths(menuWidth, bodyWidth) {
		var body = byId("tblBody");
		var menuCell = tableCell(body, 0, 0);
		var contentCell = tableCell(body, 0, 1);
		if (menuCell) {
			menuCell.width = menuWidth;
		}
		if (contentCell) {
			contentCell.width = bodyWidth;
		}
	}

	function setMenuHeader(isOpen) {
		var head = byId("tblMenuHead");
		var row;
		var titleCell;
		var imageCell;
		var label = isOpen ? "Collapse" : "Expand";
		var image = isOpen ? "../assets/images/CollapseButton.gif" : "../assets/images/ExpandButton.gif";

		if (!head) {
			return;
		}
		if (head.rows.length) {
			head.deleteRow(0);
		}
		row = head.insertRow(0);
		if (isOpen) {
			titleCell = row.insertCell();
			titleCell.innerHTML = "&nbsp;Menu";
			titleCell.className = "NavTitleText";
			titleCell.width = "50%";

			imageCell = row.insertCell();
			imageCell.className = "NavTitleImg";
			imageCell.width = "50%";
			imageCell.align = "right";
			imageCell.innerHTML = "<span style=\"cursor: pointer\"><img id=\"imgEC\" onclick=\"AccHome()\" title=\"" + label + "\" src=\"" + image + "\" border=\"2\" width=\"17\" height=\"14\" style=\"border-style: solid; border-color: #999999; background-color: #ffffff;\"></span>";
		} else {
			titleCell = row.insertCell();
			titleCell.height = 10;
			titleCell.width = 20;
			titleCell.colSpan = 2;
			titleCell.innerHTML = "<span style=\"cursor: pointer\"><img id=\"imgEC\" onclick=\"AccHome()\" title=\"" + label + "\" src=\"" + image + "\" border=\"2\" width=\"17\" height=\"14\" style=\"border-style: solid; border-color: #999999; background-color: #ffffff;\"></span>";
		}
	}

	window.hideMenu = function () {
		var head = byId("tblMenuHead");
		var menuCell;
		var menu = byId("Menu");

		setMenuHeader(false);
		menuCell = tableCell(head, 1, 0);
		if (menuCell) {
			menuCell.width = 20;
			menuCell.bgColor = "#cccccc";
		}
		setBodyWidths("10", "100%");
		if (menu) {
			menu.style.visibility = "hidden";
		}
		menuOpen = false;
	};

	window.showMenu = function () {
		var head = byId("tblMenuHead");
		var menuCell;
		var menu = byId("Menu");

		setMenuHeader(true);
		menuCell = tableCell(head, 1, 0);
		if (menuCell) {
			menuCell.bgColor = "#ffffff";
		}
		setBodyWidths("20%", "80%");
		if (menu) {
			menu.style.visibility = "visible";
		}
		menuOpen = true;
	};

	window.AccHome = function () {
		if (menuOpen) {
			window.hideMenu();
		} else {
			window.showMenu();
		}
	};

	window.DispErr = function () {};

	window.Help = function () {
		window.open("../Accounts/HelpFiles/AccHelp.htm", "", "toolbar=no,titlebar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=800px,height=500px,left=10,top=10");
		return false;
	};
}(window, document));
