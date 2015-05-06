

// Auto-select listbox

// mike pope, Visual Basic UE

// This script and the listbox on this page illustrates one 
// way to create an "auto-complete" listbox, where the

var toFind = "";              // Variable that acts as keyboard buffer
var timeoutID = "";           // Process id for timer (used when stopping 
                              // the timeout)
timeoutInterval = 500;        // Milliseconds. Shorten to cause keyboard 
                              // buffer to be cleared faster
var timeoutCtr = 0;           // Initialization of timer count down
var timeoutCtrLimit = 3 ;     // Number of times to allow timer to count 
                              // down
var oControl = "";            // Maintains a global reference to the 
                              // control that the user is working with.

function control_onkeypress(){
   // This function is called when the user presses a key while focus is in 
   // the listbox. It maintains the keyboard buffer.
   // Each time the user presses a key, the timer is restarted. 
   // First, stop the previous timer; this function will restart it.
   window.clearInterval(timeoutID)

   // Which control raised the event? We'll need to know which control to 
   // set the selection in.
   oControl = window.event.srcElement;

   var keycode = window.event.keyCode;
   if(keycode >= 32 ){
       // What character did the user type?
       var c = String.fromCharCode(keycode);
       c = c.toUpperCase(); 
       // Convert it to uppercase so that comparisons don't fail
       toFind += c ; // Add to the keyboard buffer
       find();    // Search the listbox
       timeoutID = window.setInterval("idle()", timeoutInterval);  
       // Restart the timer
    }
}

function control_onblur(){
   // This function is called when the user leaves the listbox.

   window.clearInterval(timeoutID);
   resetToFind();
}

function idle(){
   // This function is called if the timeout expires. If this is the 
   // third (by default) time that the idle function has been called, 
   // it stops the timer and clears the keyboard buffer

   timeoutCtr += 1
   if(timeoutCtr > timeoutCtrLimit){
      resetToFind();
      timeoutCtr = 0;
      window.clearInterval(timeoutID);
   }
}

function resetToFind(){
   toFind = ""
}


function find(){
	// Walk through the select list looking for a match

	//var allOptions = document.all.item(oControl.id);
	var allOptions = oControl.options;
	
	for (i=0; i < allOptions.length; i++){
		// Gets the next item from the listbox
		nextOptionText = allOptions(i).text.toUpperCase();

		// By default, the values in the listbox and as entered by the  
		// user are strings. This causes a string comparison to be made, 
		// which is not correct for numbers (1 < 11 < 2).
		// The following lines coerce numbers into an (internal) number 
		// format so that the subsequent comparison is done as a 
		// number (1 < 2 < 11).

		if(!isNaN(nextOptionText) && !isNaN(toFind) ){
			nextOptionText *= 1;        // coerce into number
			toFind *= 1;
		}

        // Does the next item match exactly what the user typed?
        if(toFind == nextOptionText){
			// OK, we can stop at this option. Set focus here
            oControl.selectedIndex = i;
            window.event.returnValue = false;
            break;
        }

        // If the string does not match exactly, find which two entries 
        // it should be between.
        if(i < allOptions.length-1){

			// If we are not yet at the last listbox item, see if the 
			// search string comes between the current entry and the next 
			// one. If so, place the selection there.

			lookAheadOptionText = allOptions(i+1).text.toUpperCase() ;
			if( toFind < nextOptionText){
				oControl.selectedIndex = i;
				window.event.cancelBubble = true;
				window.event.returnValue = false;
				break;
			}
			else{
				if( (toFind > nextOptionText) &&
					(toFind < lookAheadOptionText) ){
					oControl.selectedIndex = i+1;
					window.event.cancelBubble = true;
					window.event.returnValue = false;
					break;
				} // if
			}
		} // if

        else{

			// If we are at the end of the entries and the search string 
			// is still higher than the entries, select the last entry

			if(toFind > nextOptionText){
				oControl.selectedIndex = allOptions.length-1 // stick it 
                                                            // at the end
				window.event.cancelBubble = true;
				window.event.returnValue = false;
				break;
			} // if
		} // else
	}  // for
} // function
