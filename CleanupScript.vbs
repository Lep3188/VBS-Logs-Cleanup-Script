'Script for deleting log files based on config file |Created by Luis Peralta (9052809), luis.peralta@spiritaero.com
Dim Subcount,sfolderavailabel, filesammount, AppPath, ammountofdays, nodecountvar, SubDelete, objFSO, objFolder, colSubfolders,Fextension, i, logtext
Set objFSO = CreateObject("Scripting.FileSystemObject")
GetCurrentFolder= objFSO.GetParentFolderName(WScript.ScriptFullName)
Set msg = CreateObject("CDO.Message")
set xmlDoc = CreateObject("Microsoft.XMLDOM")
Dim SFolderList : SFolderList = Array() 'Dynamic aray to store the sub folders names. 
sfolderavailabel = False
'==========================================================================================================================
ConfFileConn()
Wscript.quit()
'==========================================================================================================================
Sub SFolderCheck()
Set objFolder = objFSO.GetFolder(AppPath)
Set colSubfolders = objFolder.Subfolders
ReDim Preserve SFolderList(UBound(SFolderList)+1)
SFolderList(UBound(SFolderList)) = AppPath
If SubDelete Then
    Subcount = 0
        For Each objSubfolder in colSubfolders
            REDIM PRESERVE SFolderList(Subcount)
            Subcount = Subcount + 1
        
            ReDim Preserve SFolderList(UBound(SFolderList)+1)
            SFolderList(UBound(SFolderList)) = AppPath & objSubfolder.name
        Next 
    If Subcount >= 1 Then
        sfolderavailabel = True
    End If
End If
    MainSub()
End Sub
Sub MainSub()
    Dim srcFile, deletedfiles
    If sfolderavailabel = True Then 'If Sub folders exist go to all folders
        For i = 0 to ubound(SFolderList)
            Set objFolder = objFSO.GetFolder(SFolderList(i))
            If objFolder.Files.Count = 0 Then
                logtext = logtext & " Currently there are no logs to delete in "  & SFolderList(i) & vbCrLf
            Else ' If sub folders dont exist just delete in the main folder. 
            DeletingFiles()
            End If
        Next
    Else
        If objFolder.Files.Count = 0 Then
                logtext = logtext & "Currently there are no logs to delete in " & AppPath & vbCrLf
        Else 
            DeletingFiles()
        End If
    End If 
End Sub
'==========================================================================================================================
Sub DeletingFiles()
    deletedfiles = 0
    Filesinfolder = 0
    Set srcFile = objFolder.Files
    For Each srcFile in objFolder.Files
        If DateDiff("d", Now, srcFile.DateLastModified) <= -ammountofdays  Then 'REMEMBER TO ADJUST THE NUMBER OF DAYSOLD
            If UCase(objFSO.GetExtensionName(srcFile.name)) = Fextension Then
                deletedfiles = deletedfiles + 1
                objFSO.DeleteFile srcFile, True
            End If
        End If
        Filesinfolder = Filesinfolder + 1
    Next
    logtext = (logtext & "Files: " & deletedfiles & " out of " & Filesinfolder &" Deleted in " & SFolderList(i) & vbCrLf)
    
End Sub
'==========================================================================================================================
Sub ConfFileConn()
LogCreation()
xmlDoc.async= "False"
xmlDoc.load(GetCurrentFolder + "\CleanupConfig.XML")
Set Nodes = xmlDoc.selectNodes("//confData/Configuration")
'msgBox ("Ammount of parent nodes: " & Nodes.length)
nodecountvar = 0
For each value in Nodes
    Set Node = Nodes.item(nodecountvar)
    Set ChildNodes = Node.childNodes
    for each x in ChildNodes
        Select Case x.nodename
            Case "IsSubFoldersRequiredDeletion"
                SubDelete = x.text
            Case "MaxFileAge"
                ammountofdays = x.text
            Case "FolderPath"
                AppPath = x.text
            Case "FileExtensionToDelete"
                Fextension = x.text                
        End Select
    next
    nodecountvar = nodecountvar + 1
    SFolderCheck()
    ReDim SFolderList(-1)
  Next 
  LogCreation()
End Sub
'==========================================================================================================================
Sub LogCreation()
Set FSO = CreateObject("Scripting.FileSystemObject")
dToday = Date()
sToday = Right("0" & Day(dToday), 2) & MonthName(Month(dToday), True) & Year(dToday)
If Not FSO.FolderExists(GetCurrentFolder & "\LogsFolder\") Then
    newfolder = FSO.CreateFolder (GetCurrentFolder & "\LogsFolder\")
End If
Set OutPutFile = FSO.OpenTextFile(GetCurrentFolder & "\LogsFolder\" & sToday & "_Log.txt" ,8 , True)
OutPutFile.WriteLine(logtext)
Set FSO= Nothing
End Sub
