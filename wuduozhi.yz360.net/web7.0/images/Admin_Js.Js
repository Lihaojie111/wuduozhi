
//自动获取图片组
function doChange(objText, objDrop){

     if (!objDrop) return;

     var str = objText.value;

     var arr = str.split("|");

     objDrop.length=0;

     for (var i=0; i<arr.length; i++){

         objDrop.options[i] = new Option(arr[i], arr[i]);

     }

}
//-->