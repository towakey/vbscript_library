' ���C�u������ǂݍ��ނ��܂��Ȃ�
dim objFSO_lib,objWSH_lib
set objFSO_lib=CreateObject("Scripting.FileSystemObject")
set objWSH_lib=objFSO_lib.OpenTextFile("./lib.vbs")
ExecuteGlobal objWSH_lib.ReadAll()
objWSH_lib.Close
set objWSH_lib=Nothing
set objFSO_lib=Nothing
' ���C�u������ǂݍ��ނ��܂��Ȃ�

' �e�֐��̎g����
dim arr
' arr=get_line_array("./test.txt")

' arr=get_file_array("C:\work\tools\vbs\")

' arr=get_folder_array("C:\work\tools\vbs\lib��\")

' WScript.Echo get_current_directory()

' call get_json(arr,"./test.json")
' WScript.Echo arr.date.year
' WScript.Echo arr.contents.[0]
' WScript.Echo arr.contents.length
