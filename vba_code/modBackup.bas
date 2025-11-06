Attribute VB_Name = "modBackup"
' ===============================================================================
' MODULE: modBackup
' AUTHOR: AI Assistant
' DATE: 2024-12
' PURPOSE: Backup management with automatic retention (keep last 10)
' ===============================================================================

Option Explicit

Private Const MAX_BACKUPS As Long = 10

'-------------------------------------------------------------------------------
' FUNCTION: CreateBackup
' PURPOSE: Creates a backup before import operation
' PARAMETERS: plantID - "ALL" or specific plant
' RETURNS: String - Path to backup file, or empty string on error
'-------------------------------------------------------------------------------
Public Function CreateBackup(plantID As String) As String
    On Error GoTo ErrorHandler

    Dim backupPath As String
    Dim timestamp As String
    Dim fso As Object

    ' Create Backups folder if it doesn't exist
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim backupFolder As String
    backupFolder = ThisWorkbook.Path & "\Backups"

    If Not fso.FolderExists(backupFolder) Then
        fso.CreateFolder backupFolder
    End If

    ' Generate backup filename
    timestamp = Format(Now, "yyyymmdd_hhnnss")
    backupPath = backupFolder & "\Import_Backup_" & timestamp & ".json"

    ' Export current state to JSON
    If modImportExport.ExportToJSON(plantID, backupPath) Then
        ' Backup created successfully
        CreateBackup = backupPath

        ' Clean up old backups
        CleanupOldBackups backupFolder
    Else
        CreateBackup = ""
    End If

    Set fso = Nothing
    Exit Function

ErrorHandler:
    CreateBackup = ""
    Set fso = Nothing
    Debug.Print "Error in CreateBackup: " & Err.Description
End Function

'-------------------------------------------------------------------------------
' SUB: CleanupOldBackups
' PURPOSE: Keeps only the last 10 backups, deletes older ones
' PARAMETERS: backupFolder - Path to backups folder
'-------------------------------------------------------------------------------
Private Sub CleanupOldBackups(backupFolder As String)
    On Error Resume Next

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim backupFiles As Collection
    Dim sortedFiles() As String
    Dim fileDates() As Date
    Dim i As Long
    Dim j As Long
    Dim count As Long
    Dim temp As String
    Dim tempDate As Date

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(backupFolder)
    Set backupFiles = New Collection

    ' Collect all backup files
    count = 0
    For Each file In folder.Files
        If file.Name Like "Import_Backup_*.json" Then
            count = count + 1
            ReDim Preserve sortedFiles(1 To count)
            ReDim Preserve fileDates(1 To count)
            sortedFiles(count) = file.Path
            fileDates(count) = file.DateLastModified
        End If
    Next file

    ' If we have more than MAX_BACKUPS, delete oldest
    If count > MAX_BACKUPS Then
        ' Simple bubble sort by date (oldest first)
        For i = 1 To count - 1
            For j = i + 1 To count
                If fileDates(i) > fileDates(j) Then
                    ' Swap
                    temp = sortedFiles(i)
                    sortedFiles(i) = sortedFiles(j)
                    sortedFiles(j) = temp

                    tempDate = fileDates(i)
                    fileDates(i) = fileDates(j)
                    fileDates(j) = tempDate
                End If
            Next j
        Next i

        ' Delete oldest files (keep only last MAX_BACKUPS)
        For i = 1 To count - MAX_BACKUPS
            fso.DeleteFile sortedFiles(i), True
            Debug.Print "Deleted old backup: " & sortedFiles(i)
        Next i
    End If

    Set fso = Nothing
    Set folder = Nothing
End Sub

'-------------------------------------------------------------------------------
' FUNCTION: GetBackupFolder
' PURPOSE: Returns path to backups folder
' RETURNS: String - Full path to backups folder
'-------------------------------------------------------------------------------
Public Function GetBackupFolder() As String
    GetBackupFolder = ThisWorkbook.Path & "\Backups"
End Function

'-------------------------------------------------------------------------------
' FUNCTION: ListBackups
' PURPOSE: Returns list of available backup files
' RETURNS: Collection of backup file paths (newest first)
'-------------------------------------------------------------------------------
Public Function ListBackups() As Collection
    On Error Resume Next

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim backups As Collection
    Dim backupFolder As String

    Set backups = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")

    backupFolder = GetBackupFolder()

    If Not fso.FolderExists(backupFolder) Then
        Set ListBackups = backups
        Exit Function
    End If

    Set folder = fso.GetFolder(backupFolder)

    ' Collect backup files
    For Each file In folder.Files
        If file.Name Like "Import_Backup_*.json" Then
            backups.Add file.Path
        End If
    Next file

    Set ListBackups = backups

    Set fso = Nothing
    Set folder = Nothing
End Function
