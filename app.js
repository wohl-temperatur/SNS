function myWrite(path) {
  if (!confirm('投稿してよろしいですか？\r\n\r\n' + document.FmName.tArea.value)) {
    return;
  }

  var PATH_VBS = "\\\\Pc-z560\\sns\\sendMail.vbs";

  var objNetWork = new ActiveXObject("WScript.Network");
  var jsUName = objNetWork.UserName;

  var pathFile = path;
  pathFile = pathFile.replace("file:", "");
  pathFile = pathFile.replace("%20", " ");

  var objFSO = new ActiveXObject('Scripting.FileSystemObject');
  var objText = objFSO.OpenTextFile(pathFile, 1); // 1: 読み取り専用, 2: 書き込み, 8: 追記

  var txtComment = document.FmName.tArea.value;
  txtComment = txtComment.replace("\r\n" , "<br>");
  txtComment = txtComment.replace("\n" , "<br>");
  txtComment = txtComment.replace("\r" , "<br>");
  txtComment = txtComment.replace("<br>" , "<br>\r\n");
  var now = new Date();
  var today = now.getFullYear()
    + "/" + ("0" + (now.getMonth() + 1)).slice(-2)
    + "/" + ("0" + now.getDate()).slice(-2)
    + " " + ("0" + now.getHours()).slice(-2)
    + ":" + ("0" + now.getMinutes()).slice(-2)
    + ":" + ("0" + now.getSeconds()).slice(-2);

  // <!DOCTYPE html>
  var txtLine = objText.ReadLine();
  var txtAll = txtLine + "\r\n";

  // コメント番号カウントアップ
  txtLine = objText.ReadLine();
  var myNum = 1 * txtLine.replace("<!-- ", "").replace(" -->", "");
  txtAll = txtAll + txtLine.replace(myNum, myNum + 1) + '\r\n';

  while (!objText.AtEndOfStream) {
    txtLine = objText.ReadLine();
    txtAll = txtAll + txtLine + "\r\n";
    if (txtLine == "<!-- コメントここから -->") {
      txtAll = txtAll + '<div class="comment"><span class="data">No.' + myNum + '：' + jsUName + '_' + today + '</span><br>\r\n' + txtComment + '\r\n</div>\r\n\r\n';
    }
  }
  objText.Close();
  objText = objFSO.OpenTextFile(pathFile, 2)
  objText.Write(txtAll);
  objText.Close()

  if (confirm("正常に更新できました。\r\n\r\nメールで通知しますか？")) {
    //メール送信
    var nameFile = pathFile;
    nameFile = nameFile.slice(nameFile.lastIndexOf('/') + 1, nameFile.lastIndexOf('.'));
    new ActiveXObject("WScript.Shell").Run('WScript.exe "' + PATH_VBS + '" "' + jsUName + '" "【更新】' + nameFile + '" "' + document.FmName.tArea.value + '\r\n\r\n<' + pathFile + '>"');
  }
}