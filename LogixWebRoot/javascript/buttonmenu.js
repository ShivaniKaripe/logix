    var elem = null;
// version:7.3.1.138972.Official Build (SUSDAY10202)

    elem = document.getElementById("delete")
    if (elem != null) { 
        elem.onmouseover = deleteOver;
        elem.onmouseout = deleteOut;
        deleteOut();
    }
    
    elem = document.getElementById("deploy"); 
    if (elem != null) {
        elem.onmouseover = deployOver;
        elem.onmouseout = deployOut;
        deployOut();
    }
    
    elem = document.getElementById("deploycrm"); 
    if (elem != null) {
        elem.onmouseover = crmOver;
        elem.onmouseout = crmOut;
        crmOut();
    }
    
    elem = document.getElementById("deferdeploy"); 
    if (elem != null) {
        elem.onmouseover = deferOver;
        elem.onmouseout = deferOut;
        deferOut();
    }
    
    elem = document.getElementById("download");
    if (elem != null) {
        elem.onmouseover = downloadOver;
        elem.onmouseout = downloadOut;
        downloadOut();
    }
    
    elem = document.getElementById("export")
    if (elem != null) {     
        elem.onmouseover = exportOver;
        elem.onmouseout = exportOut;
        exportOut();
    }
    
    elem = document.getElementById("exportCME")
    if (elem != null) {     
        elem.onmouseover = exportCMEOver;
        elem.onmouseout = exportCMEOut;
        exportCMEOut();
    }
    
    elem = document.getElementById("new")
    if (elem != null) {     
        elem.onmouseover = newOver;
        elem.onmouseout = newOut;
        newOut();
    }
    
    elem = document.getElementById("OfferFromTemp");
    if (elem != null) {
        elem.onmouseover = offerFromOver;
        elem.onmouseout = offerFromOut;
        offerFromOut();
    }
    
    elem = document.getElementById("save");
    if (elem != null) {
        elem.onmouseover = saveOver;
        elem.onmouseout = saveOut;
        saveOut();
    }
    
    elem = document.getElementById("saveastemp");
    if (elem != null) {
        elem.onmouseover = saveAsOver;
        elem.onmouseout = saveAsOut;
        saveAsOut();
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

    function deleteOver()    { handleMouseEvent(document.getElementById("delete"), true); }
    function deleteOut()     { handleMouseEvent(document.getElementById("delete"), false); }
    function deployOver()    { handleMouseEvent(document.getElementById("deploy"), true); }
    function deployOut()     { handleMouseEvent(document.getElementById("deploy"), false); }
    function downloadOver()  { handleMouseEvent(document.getElementById("download"), true); }
    function downloadOut()   { handleMouseEvent(document.getElementById("download"), false); }
    function crmOver()       { handleMouseEvent(document.getElementById("deploycrm"), true); }
    function crmOut()        { handleMouseEvent(document.getElementById("deploycrm"), false); }
    function deferOver()     { handleMouseEvent(document.getElementById("deferdeploy"), true); }
    function deferOut()      { handleMouseEvent(document.getElementById("deferdeploy"), false); }
    function exportOver()    { handleMouseEvent(document.getElementById("export"), true); }
    function exportOut()     { handleMouseEvent(document.getElementById("export"), false); }
    function exportCMEOver() { handleMouseEvent(document.getElementById("exportCME"), true); }
    function exportCMEOut()  { handleMouseEvent(document.getElementById("exportCME"), false); }
    function newOver()       { handleMouseEvent(document.getElementById("new"), true); }
    function newOut()        { handleMouseEvent(document.getElementById("new"), false); }
    function saveOver()      { handleMouseEvent(document.getElementById("save"), true); }
    function saveOut()       { handleMouseEvent(document.getElementById("save"), false); }
    function saveAsOver()    { handleMouseEvent(document.getElementById("saveastemp"), true); }
    function saveAsOut()     { handleMouseEvent(document.getElementById("saveastemp"), false); }
    function offerFromOver() { handleMouseEvent(document.getElementById("OfferFromTemp"), true); }
    function offerFromOut()  { handleMouseEvent(document.getElementById("OfferFromTemp"), false); }

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
            if (document.getElementById("actionsmenu") != null) {
                var  bOpen = (document.getElementById("actionsmenu").style.visibility == 'visible');
                if (bOpen) {
                    toggleDropdown();
                }
            }
        }
    }
