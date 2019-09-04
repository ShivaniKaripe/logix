/// Logix 5 JavaScript functions for desktop ///
// version:7.3.1.138972.Official Build (SUSDAY10202)

function handleClose() {
  var window = parent.document.getElementById('window');
  var iframe = parent.document.getElementById('masterframe');
  
  if (window == null && iframe == null) {
    window = parent.parent.document.getElementById('window');
    iframe = parent.parent.document.getElementById('masterframe');
  }
  
  if (window != null && iframe != null) {
    window.style.display = 'none';
    iframe.src = 'wait.aspx';
  }
}

function enableSave(e) {
  var elem = document.getElementById("save");
  var key = 0;
  
  if (typeof e != 'undefined') {
    key = e.which ? e.which : e.keyCode;
  }
  
  if (key == 9 || key == 13) {
    // tab or enter, so do nothing
  } else {
    if (elem != null) {
      elem.src = "/images/desktop/window/save-on.png";
      elem.disabled = false;
      elem.style.cursor="hand";
    }
  }
}