function show(id){						
	window.returnValue = id;
	window.close()
}
function toggle(id,closed,opened,pre){				
	var myChild = document.getElementById(id);									var myPIcon = document.getElementById("picon" + id);						var myIcon = document.getElementById("icon" + id);																			
	if(myChild.style.display=="none"){
		myChild.style.display="block";
		myIcon.src=opened;
		myPIcon.src="../../assets/images/"+ pre + "minus.png";
	}else{
		myChild.style.display="none";
		myIcon.src=closed;
		myPIcon.src="../../assets/images/"+ pre +"plus.png";
	}					
}
