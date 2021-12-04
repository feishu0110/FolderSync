Main()
Sub Main()
    on error resume next
    strSourceFolderPath="" '源文件夹
    strTargetFolderPath="" '目标文件夹，保持与源文件夹一致，如文件存在则被覆盖

    Set fso =CreateObject("Scripting.FileSystemObject")
    strCurPath= fso.GetFile(Wscript.ScriptFullName).ParentFolder.Path
    strConfigFileName="\config.xml"
    blnExist=IsFolderExist(fso, strSourceFolderPath, strTargetFolderPath,strCurPath,strConfigFileName) '检查是否已设置文件夹
    if blnExist=false Then
        blnSelected= SelectFolder(fso, strSourceFolderPath, strTargetFolderPath)
        if blnSelected=false then
            msgbox "Failed! Selected source folder and backup folder.",64,"提示"
            exit sub
        end if    
         msg=msgbox("Please confirm the follow information."&vbcrlf&"The Source folder is :"&strSourceFolderPath&vbcrlf&"The Backup folder is:"&strTargetFolderPath,64+4,"Info" )
        if msg<>6 then exit sub
        call CreateConfigFile(strSourceFolderPath,strTargetFolderPath,strCurPath,strConfigFileName)  'Create Config File
    end if
   

    'Sync Files and Folders
    dblFilesCount=0
    dblFolderCount=0
    strStartTime=now    
    call FolderSync(strSourceFolderPath,strTargetFolderPath,fso,dblFilesCount,dblFolderCount)
    'Write Log File
    strEndTime=now
    strLog=strStartTime&"-"&strStartTime&" Synchronized Successfully!"&vbcrlf&" Source Folder "&strSourceFolderPath&vbcrlf&" Backup Folder "&strTargetFolderPath&vbcrlf&" ToTal "&dblFilesCount&" Files and "&dblFolderCount&" Folders."&vbcrlf
  
    if err.number<>0 then
        strLog=strLog&vbcrlf&now&" Error "&err.number&err.description &vbcrlf
    end if
    pathLog=strCurPath&"\log.txt"
    if fso.FileExists(pathLog) then
        set fileLog=fso.OpenTextFile(pathLog,8)
        fileLog.Write strLog
    else
        set fileLog=fso.CreateTextFile(pathLog)
        fileLog.Write strLog
    end if
    fileLog.Close
    
    set fso=nothing
end sub
function IsFolderExist(fso,byref strSourceFolderPath,byref strTargetFolderPath,strCurPath,strConfigFileName)
    IsFolderExist=false
    strConfigFile=strCurPath&strConfigFileName
    fileConfig=fso.FileExists(strConfigFile)
    blnExistConfigFile=false
    if fileConfig Then
        Set xmlDoc = CreateObject("MSXML2.DOMDocument")
        xmlDoc.load(strConfigFile)
        
        FindString="//SourceFolderPath"
        Set nodes=Xmldoc.SelectNodes(FindString)
        if nodes.length>0 Then
            set node=nodes.item(0)
            strSourceFolderPath=node.getattribute("Path")
            if isNull(strSourceFolderPath) then strSourceFolderPath=""
            if fso.FolderExists(strSourceFolderPath)=false Then
                exit function
            end if 
        end if
        FindString="//TargetFolderPath"
        Set nodes=Xmldoc.SelectNodes(FindString)
        if nodes.length>0 Then
            set node=nodes.item(0)
            strTargetFolderPath=node.getattribute("Path")
             if isNull(strTargetFolderPath) then strTargetFolderPath=""
            if fso.FolderExists(strTargetFolderPath)=false Then
                exit function
            end if 
        end if
    else
        exit function
    end if
    IsFolderExist=true
end function
Function FolderSync(strSourceFolderPath,strTargetFolderPath,fso,byref dblFilesCount,byref dblFolderCount)  
    'on error resume next
   Set folderSource=fso.getFolder(strSourceFolderPath)
   Set subFolders=folderSource.subFolders
   Set files=folderSource.files
   strTargetFolderPath=strTargetFolderPath&"\"
   If Not fso.FolderExists(strTargetFolderPath) Then
   	fso.createFolder strTargetFolderPath
  End If
   For Each file In files
    fso.CopyFile folderSource&"\"&file.name,strTargetFolderPath&"\", true
    dblFilesCount=dblFilesCount+1
   Next  
   For Each subFolder In subFolders       
       Call FolderSync(subFolder.path,strTargetFolderPath&subfolder.name,fso,dblFilesCount,dblFolderCount) 
   Next
   dblFolderCount=dblFolderCount+1
End Function
function  SelectFolder(fso,byref strSourceFolderPath,byref strTargetFolderPath)
    SelectFolder=false
    Set spShell = CreateObject("Shell.Application")  
    Set spFolder1 = spShell.BrowseForFolder(0, "Please select Source Folder:", 0, ssfDRIVES) 
    If SPFolder1 Is Nothing Then Exit function
    strSourceFolderPath=SPFolder1.self.path
    Set spFolder2 = spShell.BrowseForFolder(0, "Please select Backup Folder:", 0, ssfDRIVES) 
    If SPFolder2 Is Nothing Then Exit function  
    strTargetFolderPath=SPFolder2.self.path
    if fso.FolderExists(strSourceFolderPath) and  fso.FolderExists(strTargetFolderPath) Then
        SelectFolder=true
    else
      SelectFolder=false 
    end if
end function
function CreateConfigFile(strSourceFolderPath,strTargetFolderPath,strCurPath,strConfigFileName)
 'Create Config File
    Set xmlDoc = CreateObject("MSXML2.DOMDocument") 
    Set NodeRoot=xmlDoc.SelectSingleNode("FolderConfiguration")
    If NodeRoot Is Nothing Then	
        Set NodeRoot=xmlDoc.createElement("FolderConfiguration")  
        xmlDoc.appendchild NodeRoot
    End If
    Set NodeSourceFolder=xmlDoc.SelectSingleNode("FolderConfiguration/SourceFolderPath")
    If NodeSourceFolder Is Nothing Then
        Set NodeSourceFolder=xmlDoc.createElement("SourceFolderPath")  
        call NodeSourceFolder.setattribute("Path",strSourceFolderPath)
        NodeRoot.appendchild NodeSourceFolder
    end if
    Set NodeTargetFolder=xmlDoc.SelectSingleNode("FolderConfiguration/TargetFolderPath")
    If NodeTargetFolder Is Nothing Then
        Set NodeTargetFolder=xmlDoc.createElement("TargetFolderPath")  
        call NodeTargetFolder.setattribute("Path",strTargetFolderPath)
        NodeRoot.appendchild NodeTargetFolder
    end if
    
    xmlDoc.save strCurPath&strConfigFileName
    set xmlDoc=nothing
end function
