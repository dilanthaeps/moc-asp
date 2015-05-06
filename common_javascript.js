function val_length(form1) {
	var selimg = false;
	n=document.form1.v_deleteditems.length;
	if(document.form1.v_deleteditems.length > 1 ){
		for (i=0; i< n; i++) {
			if (document.form1.v_deleteditems[i].checked != 0) {
     			selimg = true;
				break;
			}
		}
		if(!selimg){
			alert ("Choose atleast one record to delete");
			return false;
		}
		
	} else {
		if(document.form1.v_deleteditems.checked == true){
			return true;}
		else{
			alert ("Choose atleast one record to delete");
			return false;
		}
	}
}
