Sub myWrite(path)
  If MsgBox(Document.FmName.tArea.Value, vbYesNo, "���M���Ă�낵���ł����H") = vbNo Then
    WScript.Quit
  End If

  Const PATH_VBS = "\\Pc-z560\sns\sendMail.vbs" 

  Dim pathFile
  pathFile = path
  pathFile = Replace(pathFile, "file:///", "")
  pathFile = Replace(pathFile, "%20", " ")

  Dim objFSO
  Dim objText
  msgbox 0
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  msgbox 1
  Set objText = objFSO.OpenTextFile(pathFile, 1)
  msgbox 2

  Dim txtComment
  txtComment = Document.FmName.tArea.Value
  txtComment = Replace(txtComment, vbCrLf , "<br>")
  txtComment = Replace(txtComment, vbLf , "<br>")
  txtComment = Replace(txtComment, vbCr , "<br>")
  txtComment = Replace(txtComment, "<br>", "<br>" & vbLf)
  Dim myToday
  myToday = Year(Now) & "/" & Month(Now) & "/" & Day(Now)

  '<!DOCTYPE html>
  Dim txtLine
  Dim txtAll
  txtLine = objText.ReadLine
  txtAll = txtLine & vbLf

  '�R�����g�ԍ��J�E���g�A�b�v
  Dim myNum
  txtLine = objText.ReadLine
  myNum = Clng(Replace(Replace(txtLine, "<!-- ", ""), " -->", ""))
  txtAll = txtAll & Replace(txtLine, myNum, myNum + 1) & vbLf

  Do While objText.AtEndOfStream <> True
    txtLine = objText.ReadLine
    txtAll = txtAll & txtLine & vbLf
    If txtLine = "<!-- �R�����g�������� -->" then
        txtAll = txtAll & "<div class=""comment""><span class=""data"">No." & myNum & "�F" & jsUName & "_" & myToday & "</span><br>" & vbLf & txtComment & vbLf & "</div>" & vbLf & vbLf
    End If
  Loop
  objText.Close
  Set objText = objFSO.OpenTextFile(pathFile, 2)
    objText.Write txtAll
  objText.Close
  'objText = Nothing    '��������G���[
  'objFSO = Nothing    '��������G���[
  MsgBox "����ɍX�V�ł��܂����B"

  '���[�����M
  Dim nameFile
  nameFile = pathFile
  nameFile = Left(nameFile, InStrRev(nameFile, ".") - 1)
  nameFile = Right(nameFile, Len(nameFile) - InStrRev(nameFile, "/"))
  CreateObject("WScript.Shell").Run "WScript.exe """ & PATH_VBS & """ """ & jsUName & """ ""�y�X�V�z" & nameFile & """ """ & Document.FmName.tArea.Value & vbLf & vbLf & "<" & pathFile & ">"""
End Sub