function myWrite(path) {
  if (!confirm('���e���Ă�낵���ł����H\r\n\r\n' + document.FmName.tArea.value)) {
    return;
  }

  var PATH_VBS = "\\\\Pc-z560\\sns\\sendMail.vbs";

  var objNetWork = new ActiveXObject("WScript.Network");
  var jsUName = objNetWork.UserName;

  var pathFile = path;
  pathFile = pathFile.replace("file:", "");
  pathFile = pathFile.replace("%20", " ");

  var objFSO = new ActiveXObject('Scripting.FileSystemObject');
  var objText = objFSO.OpenTextFile(pathFile, 1); // 1: �ǂݎ���p, 2: ��������, 8: �ǋL

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

  // �R�����g�ԍ��J�E���g�A�b�v
  txtLine = objText.ReadLine();
  var myNum = 1 * txtLine.replace("<!-- ", "").replace(" -->", "");
  txtAll = txtAll + txtLine.replace(myNum, myNum + 1) + '\r\n';

  while (!objText.AtEndOfStream) {
    txtLine = objText.ReadLine();
    txtAll = txtAll + txtLine + "\r\n";
    if (txtLine == "<!-- �R�����g�������� -->") {
      txtAll = txtAll + '<div class="comment"><span class="data">No.' + myNum + '�F' + jsUName + '_' + today + '</span><br>\r\n' + txtComment + '\r\n</div>\r\n\r\n';
    }
  }
  objText.Close();
  objText = objFSO.OpenTextFile(pathFile, 2)
  objText.Write(txtAll);
  objText.Close()

  if (confirm("����ɍX�V�ł��܂����B\r\n\r\n���[���Œʒm���܂����H")) {
    //���[�����M
    var nameFile = pathFile;
    nameFile = nameFile.slice(nameFile.lastIndexOf('/') + 1, nameFile.lastIndexOf('.'));
    new ActiveXObject("WScript.Shell").Run('WScript.exe "' + PATH_VBS + '" "' + jsUName + '" "�y�X�V�z' + nameFile + '" "' + document.FmName.tArea.value + '\r\n\r\n<' + pathFile + '>"');
  }
}