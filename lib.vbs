Option Explicit

' ファイルを読み込み、1行を1個とした配列を返す
' 引数：String(ファイルパス)
' 返数：Array(1行を1個に入れた配列)
function get_line_array(path)
    ' エラー発生しても続けて処理
    On Error Resume Next

    dim objFSO
    set objFSO=CreateObject("Scripting.FileSystemObject")
    dim objOTF
    set objOTF=objFSO.OpenTextFile(path)
    if Err.Number<>0 then
        get_line_array="File Not Open."
        objOTF.Close
        set objOTF=Nothing
        set objFSO=Nothing
        On Error Goto 0
        exit function
    end if
    dim arr()
    dim line,cnt
    cnt=0
    redim arr(cnt)
    do until objOTF.AtEndOfStream
        line=objOTF.ReadLine
        arr(Ubound(arr))=line
        redim preserve arr(Ubound(arr)+1)
        cnt=cnt+1
    loop
    if cnt>0 then
        redim preserve arr(Ubound(arr)-1)
    end if
    objOTF.Close
    set objOTF=Nothing
    set objFSO=Nothing

    On Error Goto 0
    get_line_array=arr
end function

' フォルダ内にファイルリストを配列として返す
' 引数：String(フォルダパス)
' 返数：Array(ファイル名を入れた配列)
function get_file_array(path)
    ' エラー発生しても続けて処理
    On Error Resume Next

    dim objFSO
    set objFSO=CreateObject("Scripting.FileSystemObject")
    dim objGF
    set objGF=objFSO.GetFolder(path)
    if Err.Number<>0 then
        get_file_array="FileList Not Get"
        set objGF=Nothing
        set objFSO=Nothing
        On Error Goto 0
        exit function
    end if
    dim arr,cnt
    dim file

    if objGF.Files.Count=0 then
        redim arr(0)
    else
        cnt=0
        redim arr(objGF.Files.Count-1)
        for each file in objGF.Files
            arr(cnt)=file.Name
            cnt=cnt+1
        next
    end if
    set objGF=Nothing
    set objFSO=Nothing
    On Error Goto 0
    get_file_array=arr
end function

' フォルダ内にファイルリストを配列として返す
' 引数：String(フォルダパス)
' 返数：Array(ファイル名を入れた配列)
function get_folder_array(path)
    ' エラー発生しても続けて処理
    On Error Resume Next

    dim objFSO
    set objFSO=CreateObject("Scripting.FileSystemObject")
    dim objGF
    set objGF=objFSO.GetFolder(path)
    if Err.Number<>0 then
        get_folder_array="FolderList Not Get"
        set objGF=Nothing
        set objFSO=Nothing
        On Error Goto 0
        exit function
    end if

    dim arr,cnt
    dim folder

    if objGF.SubFolders.Count=0 then
        redim arr(0)
    else
        cnt=0
        redim arr(objGF.SubFolders.Count-1)
        for each folder in objGF.SubFolders
            arr(cnt)=folder.Name
            cnt=cnt+1
        next
    end if
    set objGF=Nothing
    set objFSO=Nothing

    On Error Goto 0
    get_folder_array=arr
end function

' カレントディレクトリを返す
' 引数：なし
' 返数：String(カレントディレクトリ)
function get_current_directory()
    ' エラー発生しても続けて処理
    On Error Resume Next
    dim objWS
    set objWS=CreateObject("WScript.Shell")
    dim curdir
    curdir=objWS.CurrentDirectory

    set objWS=Nothing

    On Error Goto 0
    get_current_directory=curdir
end function

function get_json(path)
    ' エラー発生しても続けて処理
    On Error Resume Next
    dim objADO
    set objADO=CreateObject("ADODB.Stream")
    objADO.Charset="UTF-8"
    objADO.Open
    objADO.LoadFromFile(path)
    objADO.Position=0

    dim json_text
    json_text=objADO.ReadText()
    objADO.Close

    dim objHF
    set objHF=CreateObject("HtmlFile")
    objHF.write "<meta http-equiv='X-UA-Compatible' content='IE=9' />"
    objHF.write "<script>document.JsonParse=function (s) {return eval('(' + s + ')');}</script>"
    objHF.write "<script>document.JsonStringify=JSON.stringify;</script>"

    dim json
    set json=objHF.JsonParse(json_text)

    ' WScript.Echo json.date.year
    ' WScript.Echo objHF.JsonStringify(json)

    set objHF=Nothing
    set objADO=Nothing
    
    On Error Goto 0

    get_json=json
    ' set json=Nothing
end function