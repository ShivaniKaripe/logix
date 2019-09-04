var gpuInterval;
// version:7.3.1.138972.Official Build (SUSDAY10202)
var gpuWidth = 200;
var gpuHeight = 100;
var gpuCount = 1;

function showGrowPopup (e,strHTML, width, height) {
  var elem = null;
  var arrPos = null;

  gpuWidth = width;
  gpuHeight = height;

  elem = document.getElementById('gpuDiv');
  if (elem == null) {  
    arrPos = gpuGetMouseXY(e);
    elem = document.createElement('div');
    elem.id = 'gpuDiv';
    elem.style.position = 'absolute';
    elem.style.textAlign = 'left';
    elem.style.top = (arrPos[1] - 35) + 'px';
    elem.style.left = (arrPos[0] - 150)  + 'px';
    elem.style.height = '21px';
    elem.style.width = '1px';
    elem.style.backgroundColor='#FFFF80';  
    elem.style.border = 'solid 1px';
    elem.style.zIndex = 999;
    elem.innerHTML = '<div style="float:right;cursor:pointer;font-size:11pt;border:solid 1px #404040;background-color:red;color:white;margin:1px;" onclick="gpuHideBox();">X</div><div id="gpuContent" style="visibility:visible;">' + strHTML + '</div>';
    document.body.appendChild(elem);  

    gpuInterval = setInterval('gpuGrowBox()', 30);
  } 
}

function gpuHideBox() {
  var elem = document.getElementById('gpuDiv');

  if (elem != null) {
    document.body.removeChild(elem);
  }  
}

function gpuGrowBox() {
  var elem = document.getElementById('gpuDiv');
  var contentElem = document.getElementById('gpuContent');
  gpuCount = gpuCount + 1;

  if (elem != null) {
      elem.style.top = (parseInt(elem.style.top) - 10) + 'px';
      elem.style.left = (parseInt(elem.style.left) - 20) + 'px';
      elem.style.height = (parseInt(elem.style.height) + 20) + 'px';
      elem.style.width = (parseInt(elem.style.width) + 40) + 'px';
   if (gpuCount > 100 || parseInt(elem.style.width) >= gpuWidth || parseInt(elem.style.height) >= gpuHeight) {
      clearInterval(gpuInterval);
      if (contentElem != null) {
        contentElem.style.visibility = 'visible';
      }
      gpuCount = 1;
    }
  }
}

function gpuGetMouseXY(e) {
  var tempX, tempY;
  var IE = document.all?true:false

  if (!e) e = window.event;

  if (IE) { // grab the x-y pos.s if browser is IE
    tempX = e.clientX + document.body.scrollLeft
    tempY = e.clientY + document.body.scrollTop
  } else {  // grab the x-y pos.s if browser is NS
    tempX = e.pageX
    tempY = e.pageY
  }  
  // catch possible negative values in NS4
  if (tempX < 0){tempX = 0}
  if (tempY < 0){tempY = 0}  

  return new Array(tempX, tempY);
}
