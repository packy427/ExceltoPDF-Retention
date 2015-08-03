Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office
Imports System.IO

'************************************************************************************************************************
'   RetentionProgram.vb     (V1.3)
'   July 2015
'   Patrick Kennedy (IPC Intern)
'   Kennedy.Patrick27@gmail.com
'
'   This program's purpose is to copy Excel files from a source folder convert them .pdf format and then save them
'   into ISO with a mirrored folder structure. It does this by opening each file silently in the background and then
'   using the "Export to PDF" function of Excel to save it as a pdf. All object creations are late binding, so it
'   should be compatible with Office 2003 and above (including future versions of Office). Version log commented below.
'
'   Functions:
'       Main()                                          main function, called from button press
'       ProcessFolder(FileSystem, Folder, Integer)      processes each folder/subfolder, recursive function
'       ProcessFile(File Object, String)                opens and exports each file to PDF and updates status bar
'       MyMkDir(String)                                 if directory is not present, makes that directory
'       GetSourceFolder(String)                         opens file select dialog box for better UI, flexibility
'
'   Globals:
'       SOURCEROOTDIR   where the source files are located, this is default for the file dialog box
'       TARGETROOTDIR   where the source files are going, should match the folder level of source
'       xlTypePDF       Excel constant for file type. Defined here for max compatibility
'       xlQualStandard  Excel constant for exported file quality. Defined here for max compatibility
'************************************************************************************************************************
Public Class MainForm
    'Public Const SOURCEROOTDIR As String = "C:\Users\PKIII\Desktop\testSource\BF4 DATA\BF4 CF DATA"
    'Public Const TARGETROOTDIR As String = "C:\Users\PKIII\Desktop\testTarget\BF4 DATA\"
    'Public Const SOURCEROOTDIR As String = "C:\Users\IPC3682\Desktop\testSource\BF4 DATA\BF4 CF DATA"
    'Public Const TARGETROOTDIR As String = "C:\Users\IPC3682\Desktop\testTarget"

    Public Const xlTypePDF As Integer = 0
    Public Const xlQualStandard As Integer = 0
    Public Const msoFileDialogFolderPicker As Integer = 4
    Public folderDialog As Object
    Public fileCount As Integer

    Public xltmp As Object
    Public SOURCEROOTDIR As String
    Public TARGETROOTDIR As String

    Public Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtDate.Text = DateTime.Now.ToLongDateString()
        SOURCEROOTDIR = My.Settings.SourceRootDirectory
        TARGETROOTDIR = My.Settings.TargetRootDirectory
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sourceDir As String

        Try
            xltmp = New Excel.Application
            'set application to temporarily force new Excel windows to work silently
            xltmp.Application.ScreenUpdating = False   'block new window from being visible
            xltmp.Application.DisplayAlerts = False    'block "Close without saving?" dialog on close
            xltmp.Application.AskToUpdateLinks = False 'block "Update Data Links" dialog on open

            'update default paths
            SOURCEROOTDIR = My.Settings.SourceRootDirectory
            TARGETROOTDIR = My.Settings.TargetRootDirectory
            Debug.Print("LOOK HERE. THIS IS WHERE I START." & vbCrLf & vbCrLf & "\/  RIGHT HERE. \/")
            Debug.Print("Source: " & SOURCEROOTDIR)
            Debug.Print("Target: " & TARGETROOTDIR)

            sourceDir = GetSourceFolder(SOURCEROOTDIR)  'allow user to select source folder

            'check for canceled file dialog
            If sourceDir = "-1" Then
                GoTo earlyExit
            End If

            'format progress bar
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = countFiles(sourceDir)
            ProgressBar1.Visible = True
            ProgressBar1.Value = ProgressBar1.Minimum

            fileCount = 0             'reset file count
            ProcessFolder(sourceDir)  'copy all files over to ISO

            My.Settings.lastSynced = "Last synced: " & DateTime.Now.ToShortDateString() & " " & DateTime.Now.ToShortTimeString

            ProgressBar1.Value = ProgressBar1.Maximum

            MsgBox("Finished uploading files to ISO!")  'let user know process has finished
        Catch ex As Exception
            MsgBox("An error occured. Cancelling operations. See help tab for possible explanations")
        End Try
earlyExit:
        ' Quit Excel and release the ApplicationClass object.
        If Not xltmp Is Nothing Then
            xltmp.Quit()
            xltmp = Nothing
        End If

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub ProcessFolder(ByVal vPath As String)
        Debug.Print("Entered Process Folder")
        Dim sDirInfo As New System.IO.DirectoryInfo(vPath)  'source directory info
        Dim tDirInfo As System.IO.DirectoryInfo             'target directory info
        Dim sFileInfo As System.IO.FileInfo
        Dim targetPath As String

        If Not sDirInfo.Exists Then Exit Sub

        'get all files' sizes in current path
        On Error Resume Next
        For Each vFile As String In System.IO.Directory.GetFiles(sDirInfo.FullName)
            sFileInfo = My.Computer.FileSystem.GetFileInfo(vFile)

            If sFileInfo.Extension Like ".xls*" Then
                fileCount = fileCount + 1
                updateProgressBar(fileCount)

                targetPath = sFileInfo.DirectoryName & "\" & System.IO.Path.GetFileNameWithoutExtension(sFileInfo.FullName)
                targetPath = Replace(targetPath, SOURCEROOTDIR, TARGETROOTDIR) & ".pdf"

                tDirInfo = (New System.IO.FileInfo(targetPath)).Directory

                Debug.Print("Target Dir: " & tDirInfo.FullName)

                If Not tDirInfo.Exists Then  'check if folder doesn't already exist
                    Debug.Print("Folder no exists! " & tDirInfo.Name)
                    recursiveMkDir(tDirInfo)                 'if not, create folder and all parent folders
                End If

                If Not My.Computer.FileSystem.FileExists(targetPath) Then
                    Debug.Print("Target Path: " & targetPath)
                    ProcessFile(sFileInfo, targetPath)
                End If
            End If
        Next

        'do the same for all subfolders
        For Each vSubDir As String In System.IO.Directory.GetDirectories(sDirInfo.FullName)
            ProcessFolder(vSubDir)
        Next
    End Sub

    Sub ProcessFile(fileInfo As System.IO.FileInfo, targetPath As String)
        'initialize Excel aplpication and workbook
        Dim xlworkbook As Object
        'define parameters for PDF output
        Dim pExportFormat As Excel.XlFixedFormatType = Excel.XlFixedFormatType.xlTypePDF
        Dim pExportQuality As Excel.XlFixedFormatQuality = Excel.XlFixedFormatQuality.xlQualityStandard
        Dim pOpenAfterPublish As Boolean = False
        Dim pIncludeDocProps As Boolean = True
        Dim pIgnorePrintAreas As Boolean = True
        Dim pFromPage As Object = Type.Missing
        Dim pToPage As Object = Type.Missing

        Try
            ' Open the source workbook.
            xlworkbook = xlTmp.Workbooks.Open(fileInfo.FullName)


            ' Save it in the target format.
            If Not xlworkbook Is Nothing Then
                xlworkbook.ExportAsFixedFormat(pExportFormat, targetPath, pExportQuality, pIncludeDocProps, pIgnorePrintAreas, pFromPage, pToPage, pOpenAfterPublish)
            End If

        Catch ex As Exception
            MsgBox("Something went wrong while trying to export.  Here's the file we tried to export: " & targetPath)
            GoTo exitSub
        End Try
        ' Close the workbook object.
        If Not xlworkbook Is Nothing Then
            xlworkbook.Close(False)
            xlworkbook = Nothing
            'Runtime.InteropServices.Marshal.ReleaseComObject(xlworkbook)
        End If
exitSub:
        Exit Sub
    End Sub
    Public Sub recursiveMkDir(dirInfo As System.IO.DirectoryInfo)

        'check if parent directory exists
        If Not dirInfo.Parent.Exists Then
            recursiveMkDir(dirInfo.Parent)
        End If
        Debug.Print("Found root folder: " & dirInfo.Parent.FullName)

        'make directory once there is a parent that actually exists
        Try
            MkDir(dirInfo.FullName)
            'Debug.Print("Made dir: " & dirInfo.FullName)
        Catch ex As System.IO.IOException               '
            MsgBox("Directory already exists. You can safely ignore this message.")
        Catch ex As System.Security.SecurityException
            MsgBox("You do not have the required permissions to access this directory. Cancelling all operations.")
            'TODO: Actually cancel all operations
        End Try
    End Sub     'end myMkDir
    Function GetSourceFolder(strPath As String) As String
        Dim selection As String

        'create and format folder picker dialog
        folderDialog = New FolderBrowserDialog
        folderDialog.Description = "Select a Folder to Copy to ISO"
        folderDialog.RootFolder = Environment.SpecialFolder.Desktop
        folderDialog.selectedpath = SOURCEROOTDIR

        If (folderDialog.ShowDialog() = DialogResult.OK) Then
            selection = folderDialog.SelectedPath
        Else
            selection = "-1"
        End If

        'Debug.Print(selection)
        GetSourceFolder = selection       'return user selection
        folderDialog = Nothing            'destruct variable

    End Function    'end GetSourceFolder
    Function countFiles(path As String) As Integer
        countFiles = System.IO.Directory.GetFiles(path, "*.xls*", SearchOption.AllDirectories).Count()
        'Debug.Print("WE FOUND THIS MANY: " & fileCount)
    End Function

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Me.Close()
    End Sub
    Sub updateProgressbar(count As Integer)
        Try
            ProgressBar1.Value = fileCount
        Catch ex As ArgumentOutOfRangeException
            MsgBox("There was an error with the progress bar. You can safely ignore this error.")
        End Try
    End Sub
    Private Sub ChangeDefaultsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangeDefaultsToolStripMenuItem.Click
        Dim selection As String

        folderDialog = New FolderBrowserDialog
        folderDialog.Description = "Select a folder to set as the new default source. This folder will NOT be copied over to ISO. Only the folders within the selected folder will be included."
        folderDialog.RootFolder = Environment.SpecialFolder.Desktop
        folderDialog.selectedpath = SOURCEROOTDIR

        If (folderDialog.ShowDialog() = DialogResult.OK) Then
            selection = folderDialog.SelectedPath
            My.Settings.SourceRootDirectory = selection
        End If
        'Debug.Print(My.Settings.SourceRootDirectory)
    End Sub

    Private Sub ChangeTargetFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangeTargetFolderToolStripMenuItem.Click
        Dim selection As String

        folderDialog = New FolderBrowserDialog
        folderDialog.Description = "Select a folder to set as the target. Make sure this folder is the same level as the default source to avoid errors."
        folderDialog.RootFolder = Environment.SpecialFolder.Desktop
        folderDialog.selectedpath = TARGETROOTDIR

        If (folderDialog.ShowDialog() = DialogResult.OK) Then
            selection = folderDialog.SelectedPath
            My.Settings.TargetRootDirectory = selection
        End If
        'Debug.Print(My.Settings.TargetRootDirectory)
    End Sub
End Class

'versions
'2.0 intital export to .exe
'2.1 added progress bar and last synced info
'2.2 added menu bar and option to change default paths by using application settings
'2.3 coverted from filesystem objects to System.IO namespace with FileInfo and DirInfo
'2.4 better process handling, faster algorithms
'       .1  Progress bar seperate subprocedure
'       .2  Instead of one public Excel app, new app every file
'       .3  Skips files instead of overwrite
'       .4  Error handling for close/exit from folder dialog
'       .5  Back to one excel application
'
'ERROR NUMBERS
'