Option Explicit

Dim objFSO, objShell, strSourceFolder, objLog, strLogPath

' 创建文件系统对象和Shell对象
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' 弹出文件选择对话框选择源文件夹
strSourceFolder = objShell.BrowseForFolder(0, "选择待整理的文件夹", 0, "\\.").Self.Path

' 创建日志文件
strLogPath = strSourceFolder & "\log.txt"
Set objLog = objFSO.CreateTextFile(strLogPath, True)

' 展示整理前的目录结构
ShowFolderStructure objFSO.GetFolder(strSourceFolder), ""

' 调用函数处理文件夹
ProcessFolder objFSO.GetFolder(strSourceFolder)

' 展示整理后的目录结构
ShowFolderStructure objFSO.GetFolder(strSourceFolder), ""

' 关闭日志文件
objLog.Close

' 处理文件夹的函数
Sub ProcessFolder(Folder)
    Dim SubFolder, arrFolderName, strNewPath, i
    For Each SubFolder in Folder.SubFolders
        ' 使用"-"分割文件夹名称
        arrFolderName = Split(SubFolder.Name, "-")
        ' 如果文件夹名称没有"-"，则跳过
        If UBound(arrFolderName) = 0 Then
            objLog.WriteLine "跳过文件夹: " & SubFolder.Path
            Continue For
        End If
        strNewPath = Folder.Path
        ' 创建新的文件夹结构
        For i = 0 To UBound(arrFolderName)
            strNewPath = strNewPath & "\" & arrFolderName(i)
            If Not objFSO.FolderExists(strNewPath) Then
                objFSO.CreateFolder(strNewPath)
                objLog.WriteLine "创建文件夹: " & strNewPath
            End If
        Next
        ' 移动文件到新的文件夹
        If objFSO.FolderExists(strNewPath) Then
            objFSO.MoveFile SubFolder.Path & "\*.*", strNewPath & "\"
            objLog.WriteLine "移动文件: " & SubFolder.Path & " 到 " & strNewPath
        End If
        ' 确保所有子文件夹和文件都被移动
        MoveSubFoldersAndFiles SubFolder, strNewPath
        ' 删除原文件夹
        objFSO.DeleteFolder SubFolder.Path
        objLog.WriteLine "删除文件夹: " & SubFolder.Path
    Next
End Sub

' 移动子文件夹和文件的函数
Sub MoveSubFoldersAndFiles(SourceFolder, DestinationFolder)
    Dim SubFolder, File
    For Each SubFolder in SourceFolder.SubFolders
        If Not objFSO.FolderExists(DestinationFolder & "\" & SubFolder.Name) Then
            objFSO.CreateFolder(DestinationFolder & "\" & SubFolder.Name)
        End If
        MoveSubFoldersAndFiles SubFolder, DestinationFolder & "\" & SubFolder.Name
    Next
    For Each File in SourceFolder.Files
        If Not objFSO.FileExists(DestinationFolder & "\" & File.Name) Then
            objFSO.MoveFile File.Path, DestinationFolder & "\"
        End If
    Next
End Sub

' 展示文件夹结构的函数
Sub ShowFolderStructure(Folder, strIndent)
    Dim SubFolder
    objLog.WriteLine strIndent & Folder.Name
    For Each SubFolder in Folder.SubFolders
        ShowFolderStructure SubFolder, strIndent & "    "
    Next
End Sub
