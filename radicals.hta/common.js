
 //Added by KL Lai
 function openURL(dest_url) {
   if (dest_url != "") {
     window.open(dest_url, "_blank")
   }
 }

 function validateAllRequiredFormFields(thisform) {
   var count, thisfield;
   for (count=0; count < thisform.elements.length; count++) {
     thisfield = thisform.elements[count];
     if (thisfield.disabled) continue;

     thisfield.value = thisfield.value.replace(/^\s*/, '');
     thisfield.value = thisfield.value.replace(/\s*$/, '');

     if (thisfield.className=="required" && thisfield.value=="" ) 
     {
       try { thisfield.focus(); 
        } catch(e) { }
       alert("Please complete this required field. [" + thisfield.name +"]");
       return(false);
     }

     if ( thisfield.maxlength )
     {
       if ( thisfield.value.length > thisfield.maxlength ) {
         try { thisfield.focus(); } catch(e) { }
         alert("Your input contains " + thisfield.value.length + " characters. "
               + "It has exceeded the maximum number \nof characters (" 
               + thisfield.maxlength + ") that the system can save.");
         return(false);
       }
     }

     if ( thisfield.datatype && thisfield.value!="" ) {
       if ( thisfield.datatype == "numeric" ) {
         if ( isNaN(thisfield.value) ) {
           try { thisfield.focus(); } catch(e) { }
           alert("Please enter a numeric value. [" + thisfield.name +"]");
           return(false);
         }
       } else if ( thisfield.datatype == "date" ) {
         thisDate = thisfield.value;
         var thisDateElements = thisDate.split("-");
         //thisDateElements = thisDate.split("/");
         var monVal = {
           "JAN":0,"FEB":1,"MAR":2,"APR":3,"MAY":4,"JUN":5,
           "JUL":6,"AUG":7,"SEP":8,"OCT":9,"NOV":10,"DEC":11
         }
         if ( thisDateElements.length!=3 || isNaN(thisDateElements[0]) 
               || isNaN(monVal[thisDateElements[1].toUpperCase()]) || isNaN(thisDateElements[2]) 
               || thisDateElements[0] > 31 || thisDateElements[2] > 99 ) {
           try { thisfield.focus(); } catch(e) { }
           alert("Please enter a valid date of dd-mon-yy. [" + thisfield.name +"]");
           return(false);
         }
       }
     }
   }
   return(true);
 }

 //remember to update common_upload_inc.pl accordingly
 function showFileUpload(thisDIV, field, uploadedFile, url) {
    if (! confirm("Are you sure?") ) {
      return;
    }
    thisDIV.innerHTML = "<INPUT TYPE='file' NAME='" + field + "' SIZE='30' CLASS='deliverable'> [" 
       + "<A href='javascript:showUploadedFile(" + thisDIV.id + ",\"" + field + "\"," 
       + "\"" + uploadedFile + "\", \"" + url + "\");'>X</A>] ";
   alert("You may either leave this field blank to delete the previous " 
         + "uploaded file, or click Browse to select another file replacing the previous "
         + "file. To cancel this, please click [X].");
 }
 function showUploadedFile(thisDIV, field, uploadedFile, url) {
    thisDIV.innerHTML = uploadedFile + " [<A target=_blank href='" + url + "'>View</A>] ["
       + "<A href='javascript:showFileUpload(" + thisDIV.id + ",\"" + field + "\","
       + "\"" + uploadedFile + "\", \"" + url + "\");'>Change</A>] ";
 }

 //original common routines

 function setChkIndx(theName) { 
  for (var i=0; i < theName.length; i++ ) {
   if (theName[i].checked == true) return i; 
  }
  return "*"; // One item is always supposed to be checked but... 
 }

 function findChkIndx_MS(theName) { 
  var indxlist = "";
  for (var i=0; i < theName.options.length; i++ ) {
   if (theName.options[i].selected == true) indxlist = indxlist + "," + i; 
  }
  return (indxlist == "")?indxlist = "*":indxlist;
 }

 function setChkIndx_MS(formObj,theList) {
  for (var i=0; i < formObj.options.length; i++) { 
   formObj.options[i].selected = false;
  }
  var ilen = 0;
  while ( ilen < theList.length-1 ) { 
   var indxstart = theList.indexOf(',',ilen);
   if (indxstart == -1) return;
   ilen = theList.indexOf(',',indxstart+1);
   if (ilen == -1) ilen = theList.length;
   var indx = parseInt(theList.substring(indxstart+1,ilen) ,10);
   formObj.options[indx].selected = true;
  }
 }

 // Add Option to pulldown menu function
 function Add_Option(item) {
   entries = item.length;
   if ( item.options[item.selectedIndex].text.match(/Add /) ) { 
     var val = prompt('Enter the new option name', '');
     if (val) {
       var pattern = /(\w)(\w*)/; // a letter, and then one, none or more letters 
       var a = val.split(/\s+/g); // split the sentence into an array of words
       for (i = 0 ; i < a.length ; i ++ ) {
        var parts = a[i].match(pattern); // just a temp variable to store the fragments in.
        var firstLetter = parts[1].toUpperCase();
        var restOfWord = parts[2].toLowerCase();
        a[i] = firstLetter + restOfWord; // re-assign it back to the array and move on
       }
       var project_name = a.join(' '); // join it back together
       optionName = new Option(project_name, project_name, false, true)
       item.options[entries] = optionName;
     } else {
       item.selectedIndex = 0; 
     }
   }
 }

 function GetCookie (CookieName) {
  var cname = CookieName + "=";
  var i = 0;
  while (i < document.cookie.length) {
   var j = i + cname.length;
   if (document.cookie.substring(i, j) == cname){
    var leng = document.cookie.indexOf (";", j);
    if (leng == -1) leng = document.cookie.length;
    return unescape(document.cookie.substring(j, leng));
   }
   i = document.cookie.indexOf(" ", i) + 1;
   if (i == 0) break;
  }
  return "*";
 }

 function MakeCookieArray(cookieValue) {
  var i = 0,indx = 0, citemlen = 0;
  ckArray = new Array();
  if ( cookieValue == null ) {ckArray[0]= "*";return}//Data has expired or never entered.
  if ( cookieValue == "*") {ckArray[0]= "*";return}//Data has expired or never entered.
  while (citemlen < cookieValue.length) {
   citemlen=(cookieValue.indexOf("`", indx)>0)
   ?cookieValue.indexOf("`", indx):cookieValue.length;
   ckArray[i]= cookieValue.substring(indx, citemlen); i++;
   indx = citemlen + 1;
  }
 }

 // Popup Help Functions
 var helpwindow;
 function Help(header, message, scroll) {
  if (scroll == "yes") {
   helpwindow=window.open('../blank.htm','Help','resizable=yes,menubar=no,scrollbars=yes,directories=no,status=no,location=no,WIDTH=320,HEIGHT=320');
  }
  else {
   helpwindow=window.open('../blank.htm','Help','resizable=yes,menubar=no,scrollbars=no,directories=no,status=no,location=no,WIDTH=300,HEIGHT=300');
  }
  if (!helpwindow.opener == null) helpwindow.opener = self;
  var H = "<FONT FACE='arial' SIZE='2'><U>" + header + "</U></FONT>";
  var M = "<FONT FACE='arial' SIZE='2'>" + message + "</FONT>";
  htmlpage = "";
  htmlpage = "<BODY BGCOLOR='#ffffee'>" + "<P>" + H + "</P>" + M + "<P><CENTER><FORM><INPUT TYPE='button' NAME='Close' VALUE='Close' ONCLICK='window.close();'></FORM></CENTER></P>";
  helpwindow.document.open();
  helpwindow.document.write(htmlpage);
  helpwindow.focus();
 }

 function RemoveHelp() {
  if (window.helpwindow) {
   if (!helpwindow.closed) {
    helpwindow.close();
   }
  }
 }

 // Wait Window Functions
 var waitwindow;
 function ShowWaitWindow(width,height) {
  x = (640 - width)/2, y = (480 - height)/2;
  if (screen) {
   y = (screen.availHeight - height)/2;
   x = (screen.availWidth - width)/2;
  }
  waitwindow=window.open('../uploading_please_wait.htm','UploadWait','scrollbars=no,resizable=no,WIDTH='+width+',height='+height+',screenX='+x+',screenY='+y+',top='+y+',left='+x);
  if (waitwindow.opener == null) waitwindow.opener = self;
 }

 function CloseWaitWindow() {
  if (window.waitwindow) {
   if (!waitwindow.closed) {
    waitwindow.close();
   }
  }
 }
