'************************************************************************************************************
'************************************************************************************************************
'* SiouxSoft             	
'* DLL        
'* ssFile
'* File Functions
'*
'* Copyright 2013 (c) SiouxSoft LLC          
'* http://www.siouxsoft.com

Imports System.Windows.Forms
Imports System.IO
Imports System.Runtime.InteropServices

Public Class ssFile

    'Message
    Dim ssMessage As String

    'Errors
    Dim noErrors As Boolean
    Dim listErrors As New List(Of String)

    '************************************************************************************************************
    '************************************************************************************************************
    '* Select File
    '*
    Public Function SelectFile(Optional ByVal startPath As String = "", Optional ByRef filter As String = "")

        Dim fileBrowser As OpenFileDialog = New OpenFileDialog()
        Dim selectedFile As String = ""

        noErrors = True

        Try
            With fileBrowser

                If startPath <> "" Then
                    .InitialDirectory = startPath
                End If
                .RestoreDirectory = True

                If filter = "images" Then
                    .Filter = "Image Files|*.bmp;*.jpg;*.png;*.gif;*.tiff"
                ElseIf filter = "videos" Then
                    .Filter = "Video Files|*.mp4;*.avi;*.wav;*.mov;*.wmv;*.m4v;*.mpeg"
                ElseIf filter <> "" Then
                    .Filter = filter
                Else
                    .Filter = "All files (*.*)|*.*"
                End If

                .Title = "Select File."

                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    selectedFile = .FileName
                    listErrors.Add("'" + selectedFile + "'" + " file selected.")
                Else
                    selectedFile = ""
                    listErrors.Add("No file selected!")
                End If

            End With

        Catch ex As Exception

            noErrors = False
            listErrors.Add(ex.Message)

        End Try

        Return New Tuple(Of Boolean, List(Of String), String)(noErrors, listErrors, selectedFile)

    End Function


    '************************************************************************************************************
    '************************************************************************************************************
    '* Copy File
    '*
    Public Function CopyFile(ByVal ssSource As System.IO.FileInfo, ByVal ssDestination As System.IO.FileInfo)

        noErrors = True

        Try
            ssMessage = "Copying File '" + ssSource.FullName + "'" + " to '" + ssDestination.FullName + "'"
            Dim ssDestinationPath As String = ssDestination.Directory.ToString + "\" + ssSource.Name
            My.Computer.FileSystem.CopyFile(ssSource.FullName, ssDestinationPath)
        Catch ex As Exception
            SSException(ex)
            listErrors.Add("Problem " + ssMessage + " !")
        End Try

        Return New Tuple(Of Boolean, List(Of String))(noErrors, listErrors)

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* Move File
    '*
    Public Function MoveFile(ByVal ssSource As System.IO.FileInfo, ByVal ssDestination As System.IO.FileInfo)

        noErrors = True

        Try
            ssMessage = "Moving File '" + ssSource.FullName + "'" + " to '" + ssDestination.FullName + "'"
            Dim ssDestinationPath As String = ssDestination.Directory.ToString + "\" + ssSource.Name
            My.Computer.FileSystem.CopyFile(ssSource.FullName, ssDestinationPath)
            My.Computer.FileSystem.DeleteFile(ssSource.FullName)
        Catch ex As Exception
            SSException(ex)
            listErrors.Add("Problem " + ssMessage + " !")
        End Try

        Return New Tuple(Of Boolean, List(Of String))(noErrors, listErrors)

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* Rename File
    '*
    Public Function RenameFile(ByVal ssSource As System.IO.FileInfo, ByVal ssNewName As String)

        noErrors = True

        Try
            ssMessage = "Renaming File '" + ssSource.FullName + "'" + " to '" + ssNewName + "'"
            My.Computer.FileSystem.RenameFile(ssSource.FullName, ssNewName)
        Catch ex As Exception
            noErrors = False
            SSException(ex)
            listErrors.Add("Problem " + ssMessage + " !")
        End Try

        Return New Tuple(Of Boolean, List(Of String))(noErrors, listErrors)

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* Delete File
    '*
    Public Function DeleteFile(ByVal parmFile As System.IO.FileInfo) As Tuple(Of Boolean, List(Of String))

        noErrors = True

        Try
            ssMessage = "Deleting File '" + parmFile.FullName + "'"
            My.Computer.FileSystem.DeleteFile(parmFile.FullName)
        Catch ex As Exception
            SSException(ex)
            listErrors.Add("Problem " + ssMessage + " !")
        End Try

        Return New Tuple(Of Boolean, List(Of String))(noErrors, listErrors)

    End Function


    Public Sub Delete(ByVal parmFile As String,
                    ByRef parmNoErrors As Boolean,
                    ByRef parmListErrors As List(Of String))


        Try
            ssMessage = "Deleting File '" + parmFile + "'"
            My.Computer.FileSystem.DeleteFile(parmFile)
            parmNoErrors = True
        Catch ex As Exception
            SSException(ex)
            parmListErrors.Add("Problem " + ssMessage + " !")
            parmNoErrors = False
        End Try

    End Sub



    '************************************************************************************************************
    '************************************************************************************************************
    '* Verify files exists
    '*
    Public Function Verify(ByVal listFiles As List(Of String)) As Tuple(Of Boolean, List(Of String))

        Dim fileVerified As Boolean = True
        Dim filesVerified As Boolean = True

        For Each file In listFiles
            fileVerified = Exists(file)
            If Not fileVerified Then
                filesVerified = False
                listErrors.Add("Problem locating file :")
                listErrors.Add("'" + file + "'")
            End If
        Next

        Return New Tuple(Of Boolean, List(Of String))(filesVerified, listErrors)

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* File exists
    '*
    Public Function Exists(ByVal parmFullName As String) As Boolean

        If IO.File.Exists(parmFullName) Then
            noErrors = True
        Else
            noErrors = False
        End If

        Return noErrors

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* File in use
    '*
    Public Function InUse(ByVal parmFileInfo As FileInfo) As Boolean

        Dim returnResult As Boolean = False

        Dim tempStream As FileStream = Nothing
        Try
            tempStream = parmFileInfo.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            tempStream.Close()
        Catch ex As Exception
            If TypeOf ex Is IOException AndAlso IsFileLocked(ex) Then
                returnResult = True
            End If
        End Try

        Return returnResult

    End Function

    Private Shared Function IsFileLocked(parmException As Exception) As Boolean

        Dim returnErrorCode As Integer

        returnErrorCode = Marshal.GetHRForException(parmException) And ((1 << 16) - 1)

        Return returnErrorCode = 32 OrElse returnErrorCode = 33

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* Duplicate
    '* Returns True if the filename ends in "(<number>)"
    '* [(filename) / (True/False)]
    '*
    Public Function Duplicate(ByVal parmString As String) As Boolean

        Dim returnDuplicate As Boolean = False

        Dim tempString As String
        Dim tempStringPosition As Integer
        Dim tempStringInteger As String

        tempString = Path.GetFileNameWithoutExtension(parmString)

        If Mid(tempString, tempString.Length - 10, 11) = " - Shortcut" Then
            tempString = Mid(tempString, 1, tempString.Length - 11)
            tempString = Path.GetFileNameWithoutExtension(tempString)
        End If


        If Mid(tempString, tempString.Length, 1) = ")" Then

            tempStringPosition = tempString.Length - 1
            tempStringInteger = ""
            Do Until (Mid(tempString, tempStringPosition, 1) = "(") Or (tempStringPosition = 1)
                tempStringInteger = (Mid(tempString, tempStringPosition, 1)) + tempStringInteger
                tempStringPosition -= 1
            Loop
            If tempStringPosition = 1 Then
                returnDuplicate = False
            Else
                If IsNumeric(tempStringInteger) Then
                    returnDuplicate = True
                Else
                    returnDuplicate = False
                End If
            End If

        End If


        If Mid(tempString, tempString.Length - 3, 4) = "Copy" Then

            returnDuplicate = True

        End If


        Return returnDuplicate

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* SSException
    '* 
    Private Sub SSException(ByVal ssException As Exception)

        noErrors = False
        listErrors.Add(ssException.Message)

    End Sub


End Class
