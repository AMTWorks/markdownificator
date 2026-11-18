Attribute VB_Name = "Markdownificator_v1_2"
Option Explicit

' Markdownificator v1.2
' EN: Export the active Word document to GitHub Flavored Markdown (GFM) with Pandoc.
' ES: Exporta el documento activo de Word a GitHub Flavored Markdown (GFM) con Pandoc.
'
' EN: J.F. Franco (@AMTWorks) - project direction and testing in Word.
' ES: J.F. Franco (@AMTWorks) - dirige el proyecto y realiza las pruebas en Word.
' EN: Claude (Anthropic) - development of v1.1.
' ES: Claude (Anthropic) - desarrollo de v1.1.
' EN: ChatGPT (OpenAI) - development of v1.1.1 and v1.2.
' ES: ChatGPT (OpenAI) - desarrollo de v1.1.1 y v1.2.

Private Type InterfaceMessages
    GetActiveDocument As String
    SelectDestination As String
    LocatePandoc As String
    CreateSnapshot As String
    ConvertDocument As String
    SaveMarkdown As String
    FailurePrefix As String
    DialogTitle As String
    ProtectedView As String
    NoDocument As String
    FolderConflict As String
    ReplaceFile As String
    PandocMissing As String
    MarkdownMissing As String
    SaveError As String
    ImagesWarning As String
    PandocWarnings As String
    SavedTo As String
    DialogError As String
    PandocExit As String
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As LongPtr
    lpstrCustomFilter As LongPtr
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As LongPtr
    nMaxFile As Long
    lpstrFileTitle As LongPtr
    nMaxFileTitle As Long
    lpstrInitialDir As LongPtr
    lpstrTitle As LongPtr
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As LongPtr
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
    pvReserved As LongPtr
    dwReserved As Long
    FlagsEx As Long
End Type

Private Declare PtrSafe Function MoveFileExW Lib "kernel32" ( _
    ByVal existingPath As LongPtr, ByVal newPath As LongPtr, _
    ByVal flags As Long) As Long
Private Declare PtrSafe Function GetSaveFileNameW Lib "comdlg32.dll" ( _
    ByRef dialog As OPENFILENAME) As Long
Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
    ByVal owner As LongPtr, ByVal prompt As LongPtr, ByVal caption As LongPtr, _
    ByVal buttons As Long) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr

Private Declare PtrSafe Function CLSIDFromString Lib "ole32" ( _
    ByVal textPtr As LongPtr, ByRef classId As GUID) As Long
Private Declare PtrSafe Function SHGetKnownFolderPath Lib "shell32" ( _
    ByRef folderId As GUID, ByVal flags As Long, ByVal token As LongPtr, _
    ByRef pathPtr As LongPtr) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal textPtr As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal destination As LongPtr, ByVal source As LongPtr, ByVal length As LongPtr)
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal memory As LongPtr)
Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_EXPLORER As Long = &H80000
Private Const FOLDERID_DOWNLOADS As String = "{374DE290-123F-4565-9164-39C4925E467B}"

Private Const MOVEFILE_REPLACE_EXISTING As Long = 1
Private Const WINDOW_HIDDEN As Long = 0
Private Const APP_TITLE As String = "Markdownificator v1.2"

Public Sub ExportarAMarkdown()
    Dim ui As InterfaceMessages
    Dim source As Document
    Dim snapshot As Document
    Dim para As Paragraph
    Dim titleStyleName As String
    Dim fs As Object
    Dim shell As Object
    Dim destination As String
    Dim workFolder As String
    Dim stagingPath As String
    Dim xmlPath As String
    Dim docxPath As String
    Dim markdownPath As String
    Dim errorLogPath As String
    Dim pandocPath As String
    Dim openXml As String
    Dim stepName As String
    Dim failure As String
    Dim diagnostics As String
    Dim message As String
    Dim moveFlags As Long
    Dim windowsError As Long
    Dim hasImages As Boolean
    Dim tempCandidate As String
    Dim stagingFile As Object
    Dim previousSecurity As MsoAutomationSecurity
    Dim restoreSecurity As Boolean

    On Error GoTo Failed
    LoadMessages ui, UseSpanishUI()

    stepName = ui.GetActiveDocument
    If IsProtectedViewActive() Then
        ShowMessage ui.ProtectedView, vbExclamation
        Exit Sub
    End If
    If Documents.Count = 0 Then
        ShowMessage ui.NoDocument, vbInformation
        Exit Sub
    End If
    Set source = ActiveDocument
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    stepName = ui.SelectDestination
    destination = ChooseDestination(source, fs, ui)
    If Len(destination) = 0 Then Exit Sub
    If fs.FolderExists(destination) Then
        Err.Raise vbObjectError + 106, APP_TITLE, ui.FolderConflict
    End If
    If fs.FileExists(destination) Then
        If ShowMessage(ui.ReplaceFile & vbCrLf & destination, _
                  vbYesNo Or vbQuestion Or vbDefaultButton2) <> vbYes Then Exit Sub
        moveFlags = MOVEFILE_REPLACE_EXISTING
    End If

    stepName = ui.LocatePandoc
    pandocPath = FindPandoc(fs, shell)
    If Len(pandocPath) = 0 Then
        Err.Raise vbObjectError + 101, APP_TITLE, _
            ui.PandocMissing
    End If

    stepName = ui.CreateSnapshot
    tempCandidate = fs.BuildPath(fs.GetSpecialFolder(2).Path, "Word2MD-" & fs.GetBaseName(fs.GetTempName))
    workFolder = fs.CreateFolder(tempCandidate).Path
    xmlPath = fs.BuildPath(workFolder, "documento.xml")
    docxPath = fs.BuildPath(workFolder, "documento.docx")
    markdownPath = fs.BuildPath(workFolder, "documento.md")
    errorLogPath = fs.BuildPath(workFolder, "pandoc-errores.txt")

    openXml = source.WordOpenXML
    hasImages = (InStr(1, openXml, "<pic:pic", vbBinaryCompare) > 0) Or _
                (InStr(1, openXml, "<v:imagedata", vbBinaryCompare) > 0)
    WriteUtf8 xmlPath, openXml
    openXml = vbNullString

    previousSecurity = Application.AutomationSecurity
    restoreSecurity = True
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set snapshot = Documents.Open(FileName:=xmlPath, ConfirmConversions:=False, _
        AddToRecentFiles:=False, Visible:=False, NoEncodingDialog:=True)
    Application.AutomationSecurity = previousSecurity
    restoreSecurity = False
    ' EN: Map the built-in Title style to Heading 1 in the temporary copy.
    ' ES: Aplica el encabezado de nivel 1 a los textos con el estilo integrado Title en la copia temporal.
    titleStyleName = snapshot.Styles(wdStyleTitle).NameLocal
    For Each para In snapshot.Paragraphs
        If para.Style.NameLocal = titleStyleName Then
            para.Style = wdStyleHeading1
        End If
    Next para

    snapshot.SaveAs2 FileName:=docxPath, FileFormat:=wdFormatXMLDocument, _
        AddToRecentFiles:=False
    snapshot.Close SaveChanges:=wdDoNotSaveChanges
    Set snapshot = Nothing

    stepName = ui.ConvertDocument
    diagnostics = ConvertWithPandoc(shell, fs, pandocPath, docxPath, markdownPath, errorLogPath, ui)
    If Not fs.FileExists(markdownPath) Then
        Err.Raise vbObjectError + 102, APP_TITLE, ui.MarkdownMissing
    End If

    stepName = ui.SaveMarkdown
    tempCandidate = fs.BuildPath(fs.GetParentFolderName(destination), _
        "~Word2MD-" & fs.GetBaseName(fs.GetTempName) & ".tmp")
    Set stagingFile = fs.CreateTextFile(tempCandidate, False)
    stagingPath = tempCandidate
    stagingFile.Close
    Set stagingFile = Nothing
    fs.CopyFile markdownPath, stagingPath, True
    If MoveFileExW(StrPtr(stagingPath), StrPtr(destination), moveFlags) = 0 Then
        windowsError = Err.LastDllError
        Err.Raise vbObjectError + 103, APP_TITLE, _
            Replace(ui.SaveError, "{code}", CStr(windowsError))
    End If
    stagingPath = vbNullString
    GoTo CleanUp

Failed:
    failure = ui.FailurePrefix & stepName & "." & vbCrLf & Left$(Err.Description, 650)
    Resume CleanUp

CleanUp:
    On Error Resume Next
    If restoreSecurity Then Application.AutomationSecurity = previousSecurity
    If Not snapshot Is Nothing Then snapshot.Close SaveChanges:=wdDoNotSaveChanges
    If Not stagingFile Is Nothing Then stagingFile.Close
    If Len(stagingPath) > 0 Then
        If fs.FileExists(stagingPath) Then fs.DeleteFile stagingPath, True
    End If
    If Len(workFolder) > 0 Then
        If fs.FolderExists(workFolder) Then fs.DeleteFolder workFolder, True
    End If
    On Error GoTo 0

    If Len(failure) > 0 Then
        ShowMessage failure, vbExclamation
        Exit Sub
    End If

    ' EN: After export, show a message only for images or Pandoc diagnostics.
    ' ES: Tras exportar, avisa solo si hay referencias a archivos de imagen o avisos de Pandoc.
    If hasImages Then
        message = ui.ImagesWarning
    End If
    If Len(diagnostics) > 0 Then
        If Len(message) > 0 Then message = message & vbCrLf & vbCrLf
        message = message & ui.PandocWarnings & vbCrLf & Left$(diagnostics, 450)
    End If
    If Len(message) > 0 Then
        ShowMessage ui.SavedTo & vbCrLf & destination & vbCrLf & vbCrLf & message, _
            vbExclamation
    End If
End Sub

Private Function UseSpanishUI() As Boolean
    ' EN: The low ten bits identify the primary language across regional variants.
    ' ES: Los diez bits inferiores identifican el idioma principal de cada variante regional.
    Const PRIMARY_LANGUAGE_MASK As Long = &H3FF&
    Const LANG_SPANISH As Long = &HA&
    On Error GoTo LanguageUnavailable
    UseSpanishUI = ((Application.LanguageSettings.LanguageID(msoLanguageIDUI) And PRIMARY_LANGUAGE_MASK) = LANG_SPANISH)
    Exit Function
LanguageUnavailable:
    UseSpanishUI = False
End Function

Private Sub LoadMessages(ByRef ui As InterfaceMessages, ByVal spanish As Boolean)
    With ui
        If spanish Then
            .GetActiveDocument = "obtener el documento activo"
            .SelectDestination = "elegir el archivo de destino"
            .LocatePandoc = "localizar Pandoc"
            .CreateSnapshot = "crear una copia temporal del documento"
            .ConvertDocument = "convertir el documento con Pandoc"
            .SaveMarkdown = "guardar el Markdown en la carpeta elegida"
            .FailurePrefix = "No se pudo "
            .DialogTitle = "Exportar a Markdown"
            .ProtectedView = "El documento activo est" & ChrW(&HE1) & " en Vista protegida." & vbCrLf & _
                "Habilita la edici" & ChrW(&HF3) & "n y vuelve a intentarlo."
            .NoDocument = "Abre el documento que quieras exportar."
            .FolderConflict = "Ya existe una carpeta con ese nombre en el destino."
            .ReplaceFile = ChrW(&HBF) & "Quieres reemplazar este archivo?"
            .PandocMissing = "No se encontr" & ChrW(&HF3) & " pandoc.exe. Comprueba que Pandoc est" & ChrW(&HE1) & " instalado y en PATH."
            .MarkdownMissing = "Pandoc no gener" & ChrW(&HF3) & " el archivo Markdown."
            .SaveError = "Windows no pudo guardar el archivo (c" & ChrW(&HF3) & "digo {code})."
            .ImagesWarning = "El documento contiene im" & ChrW(&HE1) & "genes. " & _
                "El Markdown incluye referencias a archivos de imagen que no se han exportado. Rev" & ChrW(&HED) & "salas antes de compartirlo."
            .PandocWarnings = "Avisos de Pandoc:"
            .SavedTo = "Markdown guardado en:"
            .DialogError = "Error del cuadro Guardar como (c" & ChrW(&HF3) & "digo {code})."
            .PandocExit = "Pandoc termin" & ChrW(&HF3) & " con c" & ChrW(&HF3) & "digo {code}."
        Else
            .GetActiveDocument = "get the active document"
            .SelectDestination = "choose the destination file"
            .LocatePandoc = "locate Pandoc"
            .CreateSnapshot = "create a temporary copy of the document"
            .ConvertDocument = "convert the document with Pandoc"
            .SaveMarkdown = "save the Markdown file in the selected folder"
            .FailurePrefix = "Could not "
            .DialogTitle = "Export to Markdown"
            .ProtectedView = "The active document is in Protected View." & vbCrLf & _
                "Enable editing and try again."
            .NoDocument = "Open the document you want to export."
            .FolderConflict = "A folder with that name already exists at the destination."
            .ReplaceFile = "Do you want to replace this file?"
            .PandocMissing = "Could not find pandoc.exe. Check that Pandoc is installed and on PATH."
            .MarkdownMissing = "Pandoc did not create the Markdown file."
            .SaveError = "Windows could not save the file (error {code})."
            .ImagesWarning = "This document contains images. " & _
                "The Markdown file references image files that were not exported. Review the image links before sharing it."
            .PandocWarnings = "Pandoc warnings:"
            .SavedTo = "Markdown saved to:"
            .DialogError = "Save As dialog error (code {code})."
            .PandocExit = "Pandoc exited with code {code}."
        End If
    End With
End Sub

Private Function ShowMessage(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle) As Long
    ShowMessage = MessageBoxW(GetActiveWindow(), StrPtr(prompt), StrPtr(APP_TITLE), buttons)
End Function

Private Function IsProtectedViewActive() As Boolean
    Dim protectedWindow As Object
    On Error Resume Next
    Set protectedWindow = Application.ActiveProtectedViewWindow
    IsProtectedViewActive = Not protectedWindow Is Nothing
    On Error GoTo 0
End Function

Private Function ChooseDestination(ByVal source As Document, ByVal fs As Object, _
    ByRef ui As InterfaceMessages) As String

    Dim saveDialog As OPENFILENAME
    Dim fileBuffer As String
    Dim fileFilter As String
    Dim initialFolder As String
    Dim dialogTitle As String
    Dim defaultExtension As String
    Dim selectedPath As String
    Dim dialogError As Long

    initialFolder = DownloadsFolder(fs)
    fileFilter = "Markdown (*.md)" & vbNullChar & "*.md" & vbNullChar & vbNullChar
    dialogTitle = ui.DialogTitle & " - " & APP_TITLE
    defaultExtension = "md"
    fileBuffer = fs.GetBaseName(source.Name) & ".md"
    ' EN: Give the dialog a full path so it opens in Downloads even after a save elsewhere.
    ' ES: Usa la ruta completa para abrir Descargas incluso tras guardar en otra carpeta.
    If Len(initialFolder) > 0 Then fileBuffer = fs.BuildPath(initialFolder, fileBuffer)
    fileBuffer = fileBuffer & String$(32768, vbNullChar)

    With saveDialog
        .lStructSize = LenB(saveDialog)
        .hwndOwner = GetActiveWindow()
        .lpstrFilter = StrPtr(fileFilter)
        .nFilterIndex = 1
        .lpstrFile = StrPtr(fileBuffer)
        .nMaxFile = Len(fileBuffer)
        .lpstrTitle = StrPtr(dialogTitle)
        .lpstrDefExt = StrPtr(defaultExtension)
        .flags = OFN_EXPLORER Or OFN_PATHMUSTEXIST Or _
                 OFN_HIDEREADONLY Or OFN_NOCHANGEDIR
    End With
    If GetSaveFileNameW(saveDialog) = 0 Then
        dialogError = CommDlgExtendedError()
        If dialogError <> 0 Then
            Err.Raise vbObjectError + 107, APP_TITLE, _
                Replace(ui.DialogError, "{code}", CStr(dialogError))
        End If
        Exit Function
    End If

    selectedPath = Left$(fileBuffer, InStr(fileBuffer, vbNullChar) - 1)
    If LCase$(Right$(selectedPath, 3)) <> ".md" Then selectedPath = selectedPath & ".md"
    ChooseDestination = selectedPath
End Function

Private Function DownloadsFolder(ByVal fs As Object) As String
    Dim folderId As GUID
    Dim pathPtr As LongPtr
    Dim folderPath As String
    Dim length As Long

    ' EN: Let Windows choose the initial folder if Downloads is unavailable.
    ' ES: Windows elige la carpeta inicial si no se puede acceder a Descargas.
    On Error GoTo ReleasePath
    If CLSIDFromString(StrPtr(FOLDERID_DOWNLOADS), folderId) <> 0 Then Exit Function
    If SHGetKnownFolderPath(folderId, 0, 0, pathPtr) = 0 And pathPtr <> 0 Then
        length = lstrlenW(pathPtr)
        If length > 0 Then
            folderPath = String$(length, vbNullChar)
            CopyMemory StrPtr(folderPath), pathPtr, LenB(folderPath)
            If fs.FolderExists(folderPath) Then DownloadsFolder = folderPath
        End If
    End If
ReleasePath:
    If pathPtr <> 0 Then CoTaskMemFree pathPtr
End Function

Private Function FindPandoc(ByVal fs As Object, ByVal shell As Object) As String
    Dim directories As String
    Dim entry As Variant
    Dim candidate As String

    directories = shell.Environment("Process")("LOCALAPPDATA") & "\Pandoc;" & _
        shell.Environment("Process")("PATH") & ";" & shell.Environment("User")("PATH")
    For Each entry In Split(directories, ";")
        entry = Trim$(Replace(shell.ExpandEnvironmentStrings(CStr(entry)), Chr$(34), ""))
        If Len(entry) > 0 Then
            If Mid$(entry, 2, 2) = ":\" Or Left$(entry, 2) = "\\" Then
                candidate = fs.BuildPath(entry, "pandoc.exe")
                If fs.FileExists(candidate) Then
                    FindPandoc = candidate
                    Exit Function
                End If
            End If
        End If
    Next entry
End Function

Private Function ConvertWithPandoc(ByVal shell As Object, ByVal fs As Object, _
    ByVal pandocPath As String, ByVal docxPath As String, _
    ByVal markdownPath As String, ByVal errorLogPath As String, ByRef ui As InterfaceMessages) As String

    Dim commandLine As String
    Dim exitCode As Long
    Dim diagnostics As String

    ' EN: Keep the inner path quotes and redirect Pandoc diagnostics to a UTF-8 file.
    ' ES: Conserva las comillas de las rutas y guarda los avisos de Pandoc en un archivo UTF-8.
    commandLine = QuotePath(shell.ExpandEnvironmentStrings("%ComSpec%")) & " /d /s /c " & Chr$(34) & _
        QuotePath(pandocPath) & " -f docx -t gfm --wrap=none -o " & QuotePath(markdownPath) & _
        " " & QuotePath(docxPath) & " 2> " & QuotePath(errorLogPath) & Chr$(34)

    ' EN: Run Pandoc in a hidden window and wait for its exit code.
    ' ES: Ejecuta Pandoc en una ventana oculta y espera a que devuelva el resultado.
    exitCode = shell.Run(commandLine, WINDOW_HIDDEN, True)

    If fs.FileExists(errorLogPath) Then
        If fs.GetFile(errorLogPath).Size > 0 Then diagnostics = TrimLineBreaks(ReadUtf8(errorLogPath))
    End If
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 105, APP_TITLE, _
            Replace(ui.PandocExit, "{code}", CStr(exitCode)) & vbCrLf & diagnostics
    End If
    ConvertWithPandoc = diagnostics
End Function

Private Sub WriteUtf8(ByVal filePath As String, ByVal contents As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText contents
    stream.SaveToFile filePath
    stream.Close
End Sub

Private Function ReadUtf8(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8 = stream.ReadText
    stream.Close
End Function

Private Function TrimLineBreaks(ByVal text As String) As String
    text = Trim$(text)
    Do While Len(text) > 0 And (Right$(text, 1) = vbCr Or Right$(text, 1) = vbLf)
        text = Trim$(Left$(text, Len(text) - 1))
    Loop
    TrimLineBreaks = text
End Function

Private Function QuotePath(ByVal filePath As String) As String
    QuotePath = Chr$(34) & filePath & Chr$(34)
End Function
