var frm
frm = window.frames;

function Refresh() {
	frm(0).ctlClassificationTree.populateTree()
}

function Attributes() {
	gName = frm(0).ctlClassificationTree.GetText
	gPath = frm(0).ctlClassificationTree.GetFullPath
	gKey = frm(0).ctlClassificationTree.GetKey
	GetAttributes(gName,gPath,gKey)
}

function NewGroup() {
	gKey = frm(0).ctlClassificationTree.GetKey
	NewGroupValidate(gKey)
}
