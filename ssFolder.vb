'************************************************************************************************************
'************************************************************************************************************
'* SiouxSoft             	
'* DLL        
'* ssFolder
'* Folder Functions
'*
'* Copyright 2015 (c) SiouxSoft LLC          
'* http://www.siouxsoft.com

Imports System.Windows.Forms
Imports System.IO


Public Class ssFolder

    'SiouxSoft Libraries
    Dim ssSystem As New Global.ssSystem.ssSystem

    'Messages
    Dim ssAction As String

    'Errors
    Dim noErrors As Boolean
    Dim listErrors As New List(Of String)

    '************************************************************************************************************
    '************************************************************************************************************
    '* Search for files in folder
    '*
    Public Function SearchForFiles(ByVal parmFullName As String,
                                   Optional ByVal parmFilters As String() = Nothing,
                                   Optional ByVal parmSearchSubFolders As Boolean = False) As Tuple(Of Boolean, 
                                                                                            List(Of String), 
                                                                                            List(Of IO.FileInfo))

        Dim filterVideos() As String = {"*.mp4", "*.mov", "*.avi", "*.wmv", "*.mpg", "*.mpeg", "*.m4v"}

        If parmFilters Is Nothing Then
            parmFilters = {"*.*"}
        ElseIf parmFilters(0) = "Videos" Then
            parmFilters = filterVideos
        End If

        Dim listFileInfo As New List(Of IO.FileInfo)

        Try

            If Not IO.Directory.Exists(parmFullName) Then
                Throw New Exception("Folder '" + parmFullName + "' not found.")
            End If

            Dim searchOptions As IO.SearchOption
            If parmSearchSubFolders Then
                searchOptions = IO.SearchOption.AllDirectories
            Else
                searchOptions = IO.SearchOption.TopDirectoryOnly
            End If

            listFileInfo = parmFilters.SelectMany(Function(filter) New IO.DirectoryInfo(parmFullName).GetFiles(filter, searchOptions)).Where(Function(x) x.Extension <> ".db").ToList

            noErrors = True

        Catch ex As Exception

            listFileInfo = Nothing
            listErrors.Add(ex.Message)

            noErrors = False

        End Try

        Return New Tuple(Of Boolean, List(Of String), List(Of IO.FileInfo))(noErrors, listErrors, listFileInfo)

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* File Search
    '*

    Public Function FileSearch(ByRef parmNoErrors As Boolean,
                               ByRef parmListErrors As List(Of String),
                               ByVal parmFolder As String,
                               Optional ByVal parmFilter As String = "*",
                               Optional ByVal parmSearchOption As IO.SearchOption = IO.SearchOption.TopDirectoryOnly)

        Dim returnListFileInfo As New List(Of IO.FileInfo)

        Dim tempArrayFiles As String()
        Dim tempFileInfo As IO.FileInfo

        Try
            ssAction = "locating file information for '" + parmFolder + "'"
            tempArrayFiles = Directory.GetFiles(parmFolder, parmFilter, parmSearchOption)
            For Each tempFile In tempArrayFiles
                tempFileInfo = New IO.FileInfo(tempFile)
                returnListFileInfo.Add(tempFileInfo)
            Next
            parmNoErrors = True
        Catch ex As Exception
            ssSystem.Exception(ssAction, ex, parmNoErrors, parmListErrors)
        End Try

        Return returnListFileInfo

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* Return list of folders
    '*
    Public Function List(ByVal parmFolder As String) As String()

        Dim returnFolderList() As String = {}

        'Create object
        Dim tempFolderInfo As New IO.DirectoryInfo(parmFolder)

        'Loop through subfolders
        For Each subfolder As IO.DirectoryInfo In tempFolderInfo.GetDirectories()
            'Add this folders name
            Array.Resize(returnFolderList, returnFolderList.Length + 1)
            returnFolderList(returnFolderList.Length - 1) = subfolder.FullName
            'Recall function with each subdirectory
            For Each tempSubFolder As String In List(subfolder.FullName)
                Array.Resize(returnFolderList, returnFolderList.Length + 1)
                returnFolderList(returnFolderList.Length - 1) = tempSubFolder
            Next
        Next

        Return returnFolderList

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* Return list of folder information
    '*
    Public Function InfoList(ByVal parmFolder As String,
                             ByRef parmNoErrors As Boolean,
                             ByRef parmListErrors As List(Of String)) _
                                As List(Of IO.DirectoryInfo)

        Dim returnListFolderInfo As New List(Of IO.DirectoryInfo)

        Dim tempArrayFolders As String()
        Dim tempFolderInfo As IO.DirectoryInfo

        Try
            ssAction = "locating folder information for '" + parmFolder + "'"
            tempArrayFolders = Directory.GetDirectories(parmFolder, "*", SearchOption.AllDirectories)
            For Each tempFolder In tempArrayFolders
                tempFolderInfo = New IO.DirectoryInfo(tempFolder)
                returnListFolderInfo.Add(tempFolderInfo)
            Next
            parmNoErrors = True
        Catch ex As Exception
            ssSystem.Exception(ssAction, ex, parmNoErrors, parmListErrors)
        End Try

        Return returnListFolderInfo

    End Function


    '************************************************************************************************************
    '************************************************************************************************************
    '* Verify folder(s) exists
    '*
    Public Function Verify(ByVal parmListFolders As List(Of String)) As Tuple(Of Boolean, List(Of String))

        Dim folderVerified As Boolean = True
        Dim foldersVerified As Boolean = True

        For Each folder In parmListFolders
            folderVerified = Exists(folder)
            If Not folderVerified Then
                foldersVerified = False
                listErrors.Add("Problem locating folder :")
                listErrors.Add("'" + folder + "'")
            End If
        Next

        Return New Tuple(Of Boolean, List(Of String))(foldersVerified, listErrors)

    End Function

    Public Function Exists(ByVal parmFolder As String) As Boolean

        If IO.Directory.Exists(parmFolder) Then
            noErrors = True
        Else
            noErrors = False
        End If

        Return noErrors

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* Select Folder 
    ''*
    'Public Function SelectFolder(Optional ByVal parmStartFolder As String = "",
    '                             Optional ByVal parmTitle As String = "Select Folder") As Tuple(Of Boolean, 
    '                                                                                            List(Of String),
    '                                                                                            String)
    Public Function SelectFolder(ByVal parmStartFolder As String,
                                 ByVal parmSelectMessage As String,
                                 ByRef returnErrors As Boolean, ByRef returnListErrors As List(Of String)) _
                                    As String

        Dim returnFolder As String = ""

        'noErrors = True
        'listErrors.Clear()

        Dim folderBrowser = New FolderBrowserDialog()


        Try

            ssAction = "retrieving folder location"

            With folderBrowser

                If parmStartFolder <> "" Then
                    .SelectedPath = parmStartFolder
                End If

                .Description = parmSelectMessage
                .ShowNewFolderButton = False

                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    returnFolder = .SelectedPath
                    listErrors.Add("'" + parmStartFolder + "'" + " folder selected.")
                Else
                    returnFolder = ""
                    listErrors.Add("No folder selected!")
                End If

            End With

            returnErrors = False

        Catch ex As Exception

            ssSystem.Exception(ssAction, ex, returnErrors, returnListErrors)
            returnErrors = True 'remove when new exception

            'noErrors = False
            'listErrors.Add(ex.Message)

        End Try

        'Return New Tuple(Of Boolean, List(Of String), String)(noErrors, listErrors, selectedFolder)

        Return returnFolder

    End Function

    '************************************************************************************************************
    '************************************************************************************************************
    '* Delete Folder
    '*
    Public Sub DeleteFolder(ByVal parmFolder As System.IO.DirectoryInfo,
                                 ByRef parmNoErrors As Boolean,
                                 ByRef parmListErrors As List(Of String))

        Try
            ssAction = "deleting folder '" + parmFolder.FullName + "'"
            Directory.Delete(parmFolder.FullName)
            parmNoErrors = True
        Catch ex As Exception
            ssSystem.Exception(ssAction, ex, parmNoErrors, parmListErrors)
        End Try

    End Sub

    '************************************************************************************************************
    '* List File Information

    'Private Function ListFileInformation(ByVal parmFolder As String,
    '                                     ByVal parmFilters As String())

    'Dim returnListFileInfo As List(Of IO.FileInfo) = Nothing

    'Dim filesSearched = SearchForFiles(parmFolder, parmFilters, True)
    'If filesSearched.Item1 = False Then
    '    ssSystem.Warning("Problem searching for file information!",
    '                     "Verify files are accessable.",
    '                     filesSearched.Item2)
    'Else
    '    returnListFileInfo = filesSearched.Item3
    'End If

    'Return returnListFileInfo



    'Dim returnListFileInfo As New List(Of IO.FileInfo)

    'Dim tempArrayFiles As String()
    'Dim tempFileInfo As IO.FileInfo

    'Try
    '    ssAction = "locating file information for '" + parmFolder + "'"
    '    tempArrayFiles = Directory.GetFiles(parmFolder, parmFilters, parmSearchOption)
    '    For Each tempFile In tempArrayFiles
    '        tempFileInfo = New IO.FileInfo(tempFile)
    '        returnListFileInfo.Add(tempFileInfo)
    '    Next
    '    parmNoErrors = True
    'Catch ex As Exception
    '    ssSystem.Exception(ssAction, ex, parmNoErrors, parmListErrors)
    'End Try

    'Return returnListFileInfo


    'End Function


    '************************************************************************************************************
    '************************************************************************************************************
    '* Return list of folder information
    '*
    Public Function ListFolderInfo(ByVal parmFolder As String,
                                   ByRef returnErrors As Boolean, ByRef returnListErrors As List(Of String),
                                   Optional ByVal parmFilter As String = "*") _
                                        As List(Of IO.DirectoryInfo)

        Dim returnListFolderInfo As New List(Of IO.DirectoryInfo)

        Dim tempArrayFolders As String()
        Dim tempFolderInfo As IO.DirectoryInfo

        Try
            ssAction = "locating folder information for '" + parmFolder + "'"
            tempArrayFolders = Directory.GetDirectories(parmFolder, parmFilter, SearchOption.AllDirectories)
            For Each tempFolder In tempArrayFolders
                tempFolderInfo = New IO.DirectoryInfo(tempFolder)
                returnListFolderInfo.Add(tempFolderInfo)
            Next
            returnErrors = False
        Catch ex As Exception
            ssSystem.Exception(ssAction, ex, returnErrors, returnListErrors)
            returnErrors = True 'remove when new exception
        End Try

        Return returnListFolderInfo

    End Function


    '************************************************************************************************************
    '************************************************************************************************************
    '* Return list of file information
    '*
    Public Function ListFileInfo(ByVal parmFolder As String,
                                 ByRef returnErrors As Boolean, ByRef returnListErrors As List(Of String),
                                 Optional ByVal parmSearchPattern As String() = Nothing,
                                 Optional ByVal parmSearchOption As SearchOption = SearchOption.TopDirectoryOnly) _
                                        As List(Of IO.FileInfo)

        Dim returnListFileInfo As New List(Of IO.FileInfo)

        If parmSearchPattern.Count = 0 Then
            LocateFileInfo(parmFolder, "*", parmSearchOption, returnListFileInfo, returnErrors, returnListErrors)
        Else
            For Each pattern In parmSearchPattern
                LocateFileInfo(parmFolder, pattern, parmSearchOption, returnListFileInfo, returnErrors, returnListErrors)
            Next
        End If

        Return returnListFileInfo


        'Dim tempArrayFiles As String()
        'Dim tempFileInfo As IO.FileInfo

        'Try
        '    ssAction = "locating file information for '" + parmFolder + "'"
        '    tempArrayFiles = Directory.GetFiles(parmFolder, parmSearchPattern, parmSearchOption)
        '    For Each tempFolder In tempArrayFiles
        '        tempFileInfo = New IO.FileInfo(tempFolder)
        '        returnListFileInfo.Add(tempFileInfo)
        '    Next
        '    returnErrors = False
        'Catch ex As Exception
        '    ssSystem.Exception(ssAction, ex, returnErrors, returnListErrors)
        '    returnErrors = True 'remove when new exception
        'End Try

        'Return returnListFileInfo




    End Function

    Private Sub LocateFileInfo(ByVal parmFolder As String,
                               ByVal parmSearchPattern As String,
                               ByVal parmSearchOption As SearchOption,
                               ByRef returnFileInfoList As List(Of FileInfo),
                               ByRef returnErrors As Boolean, ByRef returnListErrors As List(Of String))

        'Dim tempFileInfoList As List(Of FileInfo)
        Dim tempDirectoryInfo As DirectoryInfo = New DirectoryInfo(parmFolder)


        For Each tempFileInfo In tempDirectoryInfo.EnumerateFiles(parmSearchPattern, parmSearchOption)
            'returnFileInfoList.AddRange(tempFileInfo)
            'MsgBox(tempFileInfo)
            returnFileInfoList.Add(tempFileInfo)
        Next


        'tempFileInfoList = tempDirectoryInfo.EnumerateFiles(parmSearchPattern, parmSearchOption)

        'returnFileInfoList.AddRange(tempFileInfoList)

    End Sub

    'This is the end of the class

End Class
