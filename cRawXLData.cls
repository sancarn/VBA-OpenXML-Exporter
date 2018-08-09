VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRawXLData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Description:
' This class has been designed to allow the export of OpenXML data any excel workbook and
'  load it into VBA.
' Author: James Warren
' Website: www.github.co.uk/sancarn
' License: This project is published under the MIT license.
' Host: https://github.com/sancarn/cRawXLData
' Please report issues to:
'  https://github.com/sancarn/cRawXLData/issues

'   Sub USAGE_EXAMPLE()
'       Dim raw_data As New RawXLData
'       raw_data.Init ActiveWorkbook
'       set Sheet1Pictures = raw_data.Data("xl\drawings\drawing1.xml")
'       '...
'
'       'ALTERNATIVE:
'       Dim raw_data As New RawXLData
'       raw_data.Init ActiveWorkbook, "xl\drawings"
'       set Sheet1Pictures     = raw_data.Data("drawing1.xml")
'       set Sheet1PicsRels = raw_data.Data("_rels\drawing1.xml.rels")
'
'       'ALTERNATIVE
'       Dim raw_data As New RawXLData
'       raw_data.Init ActiveWorkbook, "xl\drawings"
'       Dim key as Variant
'       for each key in raw_data.Data.Keys()
'           if key like "_rels*" then
'             Dim xml as Object
'             xml = raw_data.Data(key)
'
'             '...
'             'DO STUFF WITH XML
'             '...
'           end if
'       next key
'   End Sub


Private Declare Function MoveFileA Lib "kernel32.dll" (ByVal src As String, ByVal dest As String) As Boolean
Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal file As String) As Long

'File data
Public Data As Object

'Private properties
Private psFilterPath As String
Private pTargetWorkbook As Workbook
Private RAW_DATA_DIR As String
Private pErrorLevel As Long
Private pErrorText As String


'Error information
Public Property Get ErrorLevel() As Long
    ErrorLevel = pErrorLevel
End Property
Public Property Get ErrorText() As Long
    ErrorText = pErrorText
End Property

'Property readers
Public Property Get sFilterPath() As String
    sFilterPath = psFilterPath
End Property
Public Property Get wbTarget() As String
    Set wbTarget = pTargetWorkbook
End Property
Public Property Get sDataDir() As String
    sDataDir = RAW_DATA_DIR
End Property



'On initialize, initialize dictionary
Private Sub Class_Initialize()
    Set Data = CreateObject("Scripting.Dictionary")
End Sub


Public Function Init(wb As Workbook, Optional ByVal sFilterPath As String = "") As cRawXLData
    'Set private properties
    psFilterPath = sFilterPath
    Set pTargetWorkbook = wb
    
    'Get raw data
    RAW_DATA_DIR = GetRawWorkbookData()
    
    'Get file list
    Dim cFileList As Variant
    Set cFileList = getFullFileList(RAW_DATA_DIR & "\")
    
    'Get FSO
    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")
    
    Dim filePath As Variant
    For Each filePath In cFileList
        Select Case FSO.GetExtensionName(filePath)
            'If it's an XML file, parse the XML file
            Case "xml", "rels"
                Dim xml As Object
                Set xml = CreateObject("Microsoft.XMLDOM")
                xml.async = False
                xml.Load filePath
                Data.Add Mid(filePath, Len(RAW_DATA_DIR) + 2), xml
            'Else assume binary/txt format
            Case Else
                Data.Add Mid(filePath, Len(RAW_DATA_DIR) + 2), ReadFile(filePath)
        End Select
    Next filePath
    
    Set Init = Me
End Function




Private Function GetRawWorkbookData() As String
    On Error GoTo ERR_OCCURRED
        'Alias
        Set wb = pTargetWorkbook
        
        'Save workbook as Open XML
        Dim FSO As Object
        Set FSO = CreateObject("scripting.filesystemobject")
        
        Dim wbPath As String
        wbPath = Environ("Temp") & "\" & "TMP_" & wb.Name
        If CopyFile(wb.FullName, wbPath, 0) = 0 Then
            Err.Raise Number:=-1, Description:="Unable to copy file"
        End If
        If wb.FileFormat <> xlOpenXMLWorkbook Then
            'Create temp path
            Dim sTmpPath As String
            sTmpPath = Environ("Temp") & "\" & FSO.GetTempName() & ".xlsx"
            
            'create new workbook as OpenXML format
            Dim newWB As Workbook
            Set newWB = workbooks.Open(wbPath)
            newWB.SaveAs sTmpPath, Excel.XlFileFormat.xlOpenXMLWorkbook
            
            'Kill old file
            Kill wbPath
            
            'Get wbPath and close
            wbPath = sTmpPath
            newWB.Close
        End If
        
        'Unzip open XML document
        Dim oShell, oData As Object
        Set oShell = CreateObject("Shell.Application")
        If FSO.fileExists(wbPath & ".zip") Then
            If DeleteFile(wbPath & ".zip") = 0 Then
                Err.Raise Number:=-1, Description:="Unable to delete existing zip file"
            End If
        End If
        If MoveFileA(wbPath, wbPath & ".zip") = 0 Then
            Err.Raise Number:=-1, Description:="Unable to rename xlsx to zip"
        End If
        
        'Export Raw data from workbook
        RAW_DATA_DIR = Environ("Temp") & "\" & FSO.GetTempName()
        FSO.CreateFolder RAW_DATA_DIR
        
        'Export xl to zip file
        Dim o As Object
        Set o = oShell.Namespace(RAW_DATA_DIR & "\")
        If psFilterPath <> "" Then
            o.CopyHere oShell.Namespace(wbPath & ".zip").ParseName(psFilterPath).GetFolder().Items(), 4
        Else
            o.CopyHere oShell.Namespace(wbPath & ".zip").Items, 4
        End If
        
        'Delete zip file
        DeleteFile wbPath & ".zip"
        
        GetRawWorkbookData = RAW_DATA_DIR
        Exit Function
ERR_OCCURRED:
    On Error Resume Next
    pErrorLevel = 1
    pErrorText = "ERROR: Cannot extract raw data from workbook."
    GetRawWorkbookData = ""
    
    If RAW_DATA_DIR <> "" Then
        FSO.DeleteFolder RAW_DATA_DIR
    End If
End Function


Private Function getFullFileList(directoryPath As String) As Collection
    'Helper object
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'Call recursive filelist
    Dim filelist As New Collection
    getFullFileList_callback FSO.GetFolder(directoryPath), filelist
    
    'return generated filelist
    Set getFullFileList = filelist
End Function

Private Sub getFullFileList_callback(Folder As Object, ByRef filelist As Collection)
    'Loop through sub folders and recurse
    Dim SubFolder As Object
    For Each SubFolder In Folder.SubFolders
        getFullFileList_callback SubFolder, filelist
    Next
    
    'Loop through files, add file to filelist
    Dim file As Object
    For Each file In Folder.files
        filelist.Add file.path
    Next
End Sub

Private Function ReadFile(ByVal strPath As String) As String
    'Open file
    Open strPath For Input As #1
        'read textlines into text
        Do Until EOF(1)
            Line Input #1, textline
            text = text & textline
        Loop
    Close #1
    
    ReadFile = text
End Function

