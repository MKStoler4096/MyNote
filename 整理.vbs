Option Explicit

Dim objFSO, objShell, strSourceFolder, objLog, strLogPath

' �����ļ�ϵͳ�����Shell����
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' �����ļ�ѡ��Ի���ѡ��Դ�ļ���
strSourceFolder = objShell.BrowseForFolder(0, "ѡ���������ļ���", 0, "\\.").Self.Path

' ������־�ļ�
strLogPath = strSourceFolder & "\log.txt"
Set objLog = objFSO.CreateTextFile(strLogPath, True)

' չʾ����ǰ��Ŀ¼�ṹ
ShowFolderStructure objFSO.GetFolder(strSourceFolder), ""

' ���ú��������ļ���
ProcessFolder objFSO.GetFolder(strSourceFolder)

' չʾ������Ŀ¼�ṹ
ShowFolderStructure objFSO.GetFolder(strSourceFolder), ""

' �ر���־�ļ�
objLog.Close

' �����ļ��еĺ���
Sub ProcessFolder(Folder)
    Dim SubFolder, arrFolderName, strNewPath, i
    For Each SubFolder in Folder.SubFolders
        ' ʹ��"-"�ָ��ļ�������
        arrFolderName = Split(SubFolder.Name, "-")
        ' ����ļ�������û��"-"��������
        If UBound(arrFolderName) = 0 Then
            objLog.WriteLine "�����ļ���: " & SubFolder.Path
            Continue For
        End If
        strNewPath = Folder.Path
        ' �����µ��ļ��нṹ
        For i = 0 To UBound(arrFolderName)
            strNewPath = strNewPath & "\" & arrFolderName(i)
            If Not objFSO.FolderExists(strNewPath) Then
                objFSO.CreateFolder(strNewPath)
                objLog.WriteLine "�����ļ���: " & strNewPath
            End If
        Next
        ' �ƶ��ļ����µ��ļ���
        If objFSO.FolderExists(strNewPath) Then
            objFSO.MoveFile SubFolder.Path & "\*.*", strNewPath & "\"
            objLog.WriteLine "�ƶ��ļ�: " & SubFolder.Path & " �� " & strNewPath
        End If
        ' ȷ���������ļ��к��ļ������ƶ�
        MoveSubFoldersAndFiles SubFolder, strNewPath
        ' ɾ��ԭ�ļ���
        objFSO.DeleteFolder SubFolder.Path
        objLog.WriteLine "ɾ���ļ���: " & SubFolder.Path
    Next
End Sub

' �ƶ����ļ��к��ļ��ĺ���
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

' չʾ�ļ��нṹ�ĺ���
Sub ShowFolderStructure(Folder, strIndent)
    Dim SubFolder
    objLog.WriteLine strIndent & Folder.Name
    For Each SubFolder in Folder.SubFolders
        ShowFolderStructure SubFolder, strIndent & "    "
    Next
End Sub
