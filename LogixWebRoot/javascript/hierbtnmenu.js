    var elem = null;
// version:7.3.1.138972.Official Build (SUSDAY10202)
    
    elem = document.getElementById("btnAddToGroup"); 
    if (elem != null) {
        elem.onmouseover = addOver;
        elem.onmouseout = addOut;
        addOut();
    }

    elem = document.getElementById("btnRemove"); 
    if (elem != null) {
        elem.onmouseover = removeOver;
        elem.onmouseout = removeOut;
        removeOut();
    }

    elem = document.getElementById("btnRemoveAll")
    if (elem != null) {     
        elem.onmouseover = removeAllOver;
        elem.onmouseout = removeAllOut;
        removeAllOut();
    }   
    
    elem = document.getElementById("btnAdd")
    if (elem != null) {     
        elem.onmouseover = addHierOver;
        elem.onmouseout = addHierOut;
        addHierOut();
    }   
    
    elem = document.getElementById("btnDelete")
    if (elem != null) {     
        elem.onmouseover = deleteHierOver;
        elem.onmouseout = deleteHierOut;
        deleteHierOut();
    }   

    function handleMouseEvent(elem, highlighted) {
        if (elem != null) {
            if (highlighted) {
                elem.style.color = "white";
                elem.style.backgroundColor = "#000080";
            } else {
                elem.style.color = "black";
                elem.style.backgroundColor = "#f4f4f0";
            }
        }
    }
    
    function addOver()    { handleMouseEvent(document.getElementById("btnAddToGroup"), true); }
    function addOut()     { handleMouseEvent(document.getElementById("btnAddToGroup"), false); }
    function removeOver() { handleMouseEvent(document.getElementById("btnRemove"), true); }
    function removeOut()  { handleMouseEvent(document.getElementById("btnRemove"), false); }
    function removeAllOver() { handleMouseEvent(document.getElementById("btnRemoveAll"), true); }
    function removeAllOut()  { handleMouseEvent(document.getElementById("btnRemoveAll"), false); }
    function addHierOver()    { handleMouseEvent(document.getElementById("btnAdd"), true); }
    function addHierOut()     { handleMouseEvent(document.getElementById("btnAdd"), false); }
    function deleteHierOver()    { handleMouseEvent(document.getElementById("btnDelete"), true); }
    function deleteHierOut()     { handleMouseEvent(document.getElementById("btnDelete"), false); }
    
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    }
    else {
        document.onclick=handlePageClick;
    }

    function handlePageClick(e) {
        var el=(typeof event!=='undefined')? event.srcElement : e.target        
        if (el != null && el.id != 'actions') {
            if (document.getElementById("actionmenu") != null) {
                var  bOpen = (document.getElementById("actionmenu").style.display == 'inline');
                if (bOpen) {
                    toggleDropdown();
                }
            }
        }
    }

