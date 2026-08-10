Imports System.IO
Imports System.IO.Compression
Imports System.Reflection
Imports System.Collections.Generic
Imports Excel = Microsoft.Office.Interop.Excel

Friend NotInheritable Class DemoPackageManager
    Private Sub New()
    End Sub

    Public Shared Function ExtractAndOpen(
        application As Excel.Application,
        zipFileName As String,
        demoFolderName As String) As Excel.Workbook

        If application Is Nothing Then
            Throw New ArgumentNullException("application")
        End If

        If String.IsNullOrWhiteSpace(zipFileName) Then
            Throw New ArgumentException(
                "Der ZIP-Dateiname darf nicht leer sein.",
                "zipFileName")
        End If

        If String.IsNullOrWhiteSpace(demoFolderName) Then
            Throw New ArgumentException(
                "Der Name des Demo-Ordners darf nicht leer sein.",
                "demoFolderName")
        End If

        Dim zipPath As String = Path.Combine(
            GetAddInDirectory(),
            "Demos",
            zipFileName)
        'Dim zipPath As String = FindDemoZipFile(zipFileName)
        If Not File.Exists(zipPath) Then
            Throw New FileNotFoundException(
                "Die Demo-ZIP-Datei wurde nicht gefunden.",
                zipPath)
        End If

        Dim demoRoot As String = Path.Combine(
            Path.GetTempPath(),
            "Excel-VSTO-Toolbox")

        Dim destinationDirectory As String = Path.Combine(
            demoRoot,
            demoFolderName)

        DeleteExistingDemoDirectory(destinationDirectory)

        Directory.CreateDirectory(demoRoot)
        ZipFile.ExtractToDirectory(zipPath, destinationDirectory)

        Dim workbookPath As String =
            FindWorkbookFile(destinationDirectory)

        If String.IsNullOrEmpty(workbookPath) Then
            Throw New FileNotFoundException(
                "Im entpackten Demo-Ordner wurde keine XLSX- " &
                "oder XLSB-Datei gefunden.",
                destinationDirectory)
        End If

        Return application.Workbooks.Open(workbookPath)
    End Function
    Private Shared Function GetAddInDirectory() As String

        Dim assembly As Assembly =
        Assembly.GetExecutingAssembly()

        If Not String.IsNullOrEmpty(assembly.CodeBase) Then
            Dim codeBaseUri As New Uri(assembly.CodeBase)

            If codeBaseUri.IsFile Then
                Dim assemblyPath As String =
                codeBaseUri.LocalPath

                Return Path.GetDirectoryName(assemblyPath)
            End If
        End If

        If Not String.IsNullOrEmpty(assembly.Location) Then
            Return Path.GetDirectoryName(assembly.Location)
        End If

        Throw New InvalidOperationException(
        "Der Installationsordner des Add-ins " &
        "konnte nicht ermittelt werden.")
    End Function
    Private Shared Sub DeleteExistingDemoDirectory(
        directoryPath As String)

        If Not Directory.Exists(directoryPath) Then
            Return
        End If

        ValidateDemoDirectory(directoryPath)

        Try
            Directory.Delete(directoryPath, True)
        Catch ex As IOException
            Throw New IOException(
                "Der vorhandene Demo-Ordner konnte nicht gelöscht werden. " &
                "Möglicherweise ist eine Datei daraus noch in Excel " &
                "oder einem anderen Programm geöffnet." &
                Environment.NewLine &
                directoryPath,
                ex)
        Catch ex As UnauthorizedAccessException
            Throw New UnauthorizedAccessException(
                "Keine Berechtigung zum Löschen des Demo-Ordners:" &
                Environment.NewLine &
                directoryPath,
                ex)
        End Try
    End Sub

    Private Shared Sub ValidateDemoDirectory(
        directoryPath As String)

        Dim expectedRoot As String = Path.GetFullPath(
            Path.Combine(
                Path.GetTempPath(),
                "Excel-VSTO-Toolbox")).
            TrimEnd(Path.DirectorySeparatorChar) &
            Path.DirectorySeparatorChar

        Dim actualPath As String =
            Path.GetFullPath(directoryPath)

        If Not actualPath.StartsWith(
            expectedRoot,
            StringComparison.OrdinalIgnoreCase) Then

            Throw New InvalidOperationException(
                "Aus Sicherheitsgründen wird nur innerhalb dieses " &
                "Demo-Ordners gelöscht:" &
                Environment.NewLine &
                expectedRoot)
        End If
    End Sub

    Private Shared Function FindWorkbookFile(
        directoryPath As String) As String

        Dim allowedExtensions As String() = {
            ".xlsx",
            ".xlsb"
        }

        Dim files As String() = Directory.GetFiles(
            directoryPath,
            "*.*",
            SearchOption.AllDirectories)

        For Each filePath As String In files
            Dim extension As String =
                Path.GetExtension(filePath)

            For Each allowedExtension As String In allowedExtensions
                If extension.Equals(
                    allowedExtension,
                    StringComparison.OrdinalIgnoreCase) Then

                    Return filePath
                End If
            Next
        Next

        Return Nothing
    End Function
End Class