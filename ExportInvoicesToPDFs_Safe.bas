Option Explicit

Public Sub ExportInvoicesToPDFs()
    ' 設定: READMEから差し替えやすい箇所だけ定数化
    Const USE_FOLDER_DIALOG As Boolean = False
    Const OUTPUT_SUBFOLDER As String = "output"
    Const DEFAULT_NAME_FIELD As String = "氏名"
    Const DEFAULT_LOG_KEY_FIELD As String = ""
    Const FILE_NAME_SUFFIX As String = ""
    Const LOG_FILE_NAME As String = "export_log.csv"
    Const DOCVAR_NAME_FIELD As String = "NameField"
    Const DOCVAR_LOG_KEY_FIELD As String = "LogKeyField"

    Dim baseDoc As Document
    Dim mergedDoc As Document
    Dim mm As MailMerge
    Dim ds As MailMergeDataSource
    Dim fso As Object
    Dim outputFolder As String
    Dim logFilePath As String
    Dim originalRecord As Long
    Dim originalFirst As Long
    Dim originalLast As Long
    Dim currentIndex As Long
    Dim processedRecords As Long
    Dim maxRecords As Long
    Dim fileName As String
    Dim recordName As String
    Dim outputDate As String
    Dim stoppedByGuard As Boolean
    Dim outputFilePath As String
    Dim nameField As String
    Dim logKeyField As String
    Dim logKeyValue As String
    Dim logKeyId As String
    Dim processedKeys As Object

    On Error GoTo ErrorHandler

    If ActiveDocument Is Nothing Then
        MsgBox "アクティブな文書が見つかりません。", vbExclamation
        Exit Sub
    End If

    Set baseDoc = ActiveDocument
    If baseDoc.MailMerge.MainDocumentType = wdNotAMergeDocument Then
        MsgBox "この文書は差し込み印刷のメイン文書ではありません。", vbExclamation
        Exit Sub
    End If

    Set mm = baseDoc.MailMerge
    If mm.State <> wdMainAndDataSource Then
        MsgBox "差し込みデータへの接続を確認してください。", vbExclamation
        Exit Sub
    End If

    Set ds = mm.DataSource
    originalRecord = ds.ActiveRecord
    originalFirst = ds.FirstRecord
    originalLast = ds.LastRecord

    ds.ActiveRecord = wdFirstRecord
    If ds.ActiveRecord = wdNoActiveRecord Then
        MsgBox "差し込みデータが見つかりません。", vbExclamation
        GoTo RestoreState
    End If

    maxRecords = ds.RecordCount
    If maxRecords < 1 Then
        maxRecords = 100000
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    ' 出力先: ダイアログ or 文書と同じ場所にoutputフォルダを自動作成
    outputFolder = GetOutputFolder(baseDoc, fso, USE_FOLDER_DIALOG, OUTPUT_SUBFOLDER)
    If Len(outputFolder) = 0 Then
        If Not USE_FOLDER_DIALOG And Len(baseDoc.Path) = 0 Then
            MsgBox "文書を保存するか、フォルダ選択に切り替えてください。", vbExclamation
        Else
            MsgBox "出力先フォルダが決まりませんでした。設定を確認してください。", vbExclamation
        End If
        GoTo RestoreState
    End If

    Application.ScreenUpdating = False

    outputDate = Format$(Date, "yyyymmdd")
    nameField = EnsureDocumentVariable(baseDoc, DOCVAR_NAME_FIELD, "ファイル名に使う差し込みフィールド名を入力してください。", DEFAULT_NAME_FIELD, False)
    If Len(nameField) = 0 Then
        MsgBox "ファイル名に使うフィールド名が未設定です。", vbExclamation
        GoTo RestoreState
    End If
    logKeyField = EnsureDocumentVariable(baseDoc, DOCVAR_LOG_KEY_FIELD, "出力済判定に使うフィールド名を入力してください。（空欄ならレコード番号を使用）", DEFAULT_LOG_KEY_FIELD, True)
    logFilePath = outputFolder & LOG_FILE_NAME
    Set processedKeys = LoadProcessedKeys(logFilePath)

    Do
        currentIndex = ds.ActiveRecord
        If currentIndex = wdNoActiveRecord Then Exit Do

        processedRecords = processedRecords + 1
        logKeyValue = ""
        If Len(logKeyField) > 0 Then
            logKeyValue = GetSafeFieldValue(ds, logKeyField)
        End If
        logKeyId = BuildLogKeyId(logKeyField, logKeyValue, currentIndex)
        If processedKeys.Exists(logKeyId) Then
            GoTo AdvanceRecord
        End If

        recordName = GetSafeFieldValue(ds, nameField)
        If Len(recordName) = 0 Then
            recordName = "未設定"
        End If
        fileName = CleanFileName(outputDate & "_" & recordName & FILE_NAME_SUFFIX, 150)
        outputFilePath = BuildUniqueFilePath(outputFolder, fileName, "pdf", fso)

        mm.Destination = wdSendToNewDocument
        mm.DataSource.FirstRecord = currentIndex
        mm.DataSource.LastRecord = currentIndex
        mm.SuppressBlankLines = True
        mm.Execute Pause:=False

        Set mergedDoc = ActiveDocument
        mergedDoc.ExportAsFixedFormat _
            OutputFileName:=outputFilePath, _
            ExportFormat:=wdExportFormatPDF, _
            OpenAfterExport:=False, _
            OptimizeFor:=wdExportOptimizeForPrint, _
            Range:=wdExportAllDocument, _
            Item:=wdExportDocumentContent, _
            IncludeDocProps:=True, _
            KeepIRM:=False, _
            CreateBookmarks:=wdExportCreateNoBookmarks, _
            DocStructureTags:=True, _
            BitmapMissingFonts:=True, _
            UseISO19005_1:=False

        mergedDoc.Close SaveChanges:=False
        Set mergedDoc = Nothing
        baseDoc.Activate
        AppendLogEntry logFilePath, baseDoc.Name, currentIndex, outputFilePath, recordName, logKeyField, logKeyValue, logKeyId
        processedKeys.Add logKeyId, True

AdvanceRecord:
        mm.DataSource.FirstRecord = originalFirst
        mm.DataSource.LastRecord = originalLast
        ds.ActiveRecord = currentIndex
        ds.ActiveRecord = wdNextRecord

        If ds.ActiveRecord = wdNoActiveRecord Then Exit Do
        ' 無限ループ対策: レコードが戻る/巡回したら停止
        If ds.ActiveRecord <= currentIndex Then
            stoppedByGuard = True
            Exit Do
        End If
        If processedRecords >= maxRecords Then
            stoppedByGuard = True
            Exit Do
        End If
    Loop

    If stoppedByGuard Then
        MsgBox "全レコード処理後の追加巡回を検出したため、処理を停止しました。", vbInformation
    Else
        MsgBox "PDFの作成が完了しました。", vbInformation
    End If

RestoreState:
    On Error Resume Next
    If Not mm Is Nothing Then
        mm.DataSource.FirstRecord = originalFirst
        mm.DataSource.LastRecord = originalLast
    End If
    If Not ds Is Nothing Then
        If originalRecord = wdNoActiveRecord Then
            ds.ActiveRecord = wdFirstRecord
        Else
            ds.ActiveRecord = originalRecord
        End If
    End If
    Application.ScreenUpdating = True
    If Not baseDoc Is Nothing Then
        baseDoc.Activate
    End If
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    If Not mergedDoc Is Nothing Then
        mergedDoc.Close SaveChanges:=False
    End If
    If Not baseDoc Is Nothing Then
        baseDoc.Activate
    End If
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    Resume RestoreState
End Sub

Public Sub SetFileNameField()
    SetDocumentVariableWithPrompt "NameField", "ファイル名に使う差し込みフィールド名を入力してください。", "氏名", False
End Sub

Public Sub SetLogKeyField()
    SetDocumentVariableWithPrompt "LogKeyField", "出力済判定に使うフィールド名を入力してください。（空欄ならレコード番号を使用）", "", True
End Sub

Private Sub SetDocumentVariableWithPrompt(ByVal varName As String, ByVal prompt As String, ByVal defaultValue As String, ByVal allowBlank As Boolean)
    Dim doc As Document
    Dim value As String

    If ActiveDocument Is Nothing Then
        MsgBox "アクティブな文書が見つかりません。", vbExclamation
        Exit Sub
    End If

    Set doc = ActiveDocument
    value = Trim$(InputBox(prompt, "設定", defaultValue))
    If Len(value) = 0 And Not allowBlank Then
        MsgBox "空欄では設定できません。", vbExclamation
        Exit Sub
    End If

    SetDocumentVariable doc, varName, value
    MsgBox "設定を保存しました。", vbInformation
End Sub

Private Function GetOutputFolder(ByVal baseDoc As Document, ByVal fso As Object, ByVal useDialog As Boolean, ByVal subfolderName As String) As String
    Dim folderPath As String
    Dim dialog As FileDialog

    If useDialog Then
        Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
        dialog.Title = "PDFの出力先フォルダを選択"
        dialog.AllowMultiSelect = False
        If dialog.Show <> -1 Then
            GetOutputFolder = ""
            Exit Function
        End If
        folderPath = dialog.SelectedItems(1)
    Else
        If Len(baseDoc.Path) = 0 Then
            GetOutputFolder = ""
            Exit Function
        End If
        folderPath = baseDoc.Path & "\" & subfolderName
    End If

    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If

    GetOutputFolder = folderPath
End Function

Private Function EnsureDocumentVariable(ByVal doc As Document, ByVal varName As String, ByVal prompt As String, ByVal defaultValue As String, ByVal allowBlank As Boolean) As String
    Dim currentValue As String
    Dim inputValue As String
    Dim hasVar As Boolean

    hasVar = HasDocumentVariable(doc, varName)
    If hasVar Then
        currentValue = GetDocumentVariableValue(doc, varName)
        If allowBlank Or Len(currentValue) > 0 Then
            EnsureDocumentVariable = currentValue
            Exit Function
        End If
    End If

    inputValue = Trim$(InputBox(prompt, "設定", defaultValue))
    If Len(inputValue) = 0 And Not allowBlank Then
        EnsureDocumentVariable = ""
        Exit Function
    End If

    SetDocumentVariable doc, varName, inputValue
    EnsureDocumentVariable = inputValue
End Function

Private Function HasDocumentVariable(ByVal doc As Document, ByVal varName As String) As Boolean
    Dim temp As Variable
    On Error Resume Next
    Set temp = doc.Variables(varName)
    HasDocumentVariable = (Err.Number = 0)
    Err.Clear
End Function

Private Function GetDocumentVariableValue(ByVal doc As Document, ByVal varName As String) As String
    On Error GoTo Missing
    GetDocumentVariableValue = Trim$(doc.Variables(varName).value)
    Exit Function
Missing:
    GetDocumentVariableValue = ""
    Err.Clear
End Function

Private Sub SetDocumentVariable(ByVal doc As Document, ByVal varName As String, ByVal value As String)
    On Error Resume Next
    doc.Variables(varName).value = value
    If Err.Number <> 0 Then
        Err.Clear
        doc.Variables.Add Name:=varName, value:=value
    End If
    On Error GoTo 0
End Sub

Private Function GetSafeFieldValue(ByVal source As MailMergeDataSource, ByVal fieldName As String) As String
    On Error GoTo MissingField
    GetSafeFieldValue = Trim$(source.DataFields(fieldName).value)
    Exit Function
MissingField:
    GetSafeFieldValue = ""
    Err.Clear
End Function

Private Function BuildLogKeyId(ByVal logKeyField As String, ByVal logKeyValue As String, ByVal recordNumber As Long) As String
    If Len(logKeyField) > 0 And Len(logKeyValue) > 0 Then
        BuildLogKeyId = logKeyField & ":" & logKeyValue
    Else
        BuildLogKeyId = "RecordNumber:" & Format$(recordNumber, "000000")
    End If
End Function

Private Function LoadProcessedKeys(ByVal logFilePath As String) As Object
    Dim dict As Object
    Dim content As String
    Dim lines As Variant
    Dim i As Long
    Dim fields As Variant
    Dim keyId As String

    Set dict = CreateObject("Scripting.Dictionary")
    If Len(Dir$(logFilePath)) = 0 Then
        Set LoadProcessedKeys = dict
        Exit Function
    End If

    content = ReadTextFileUtf8(logFilePath)
    If Len(content) = 0 Then
        Set LoadProcessedKeys = dict
        Exit Function
    End If

    lines = Split(content, vbCrLf)
    For i = 1 To UBound(lines)
        If Len(lines(i)) = 0 Then
            GoTo ContinueLine
        End If
        fields = ParseCsvLine(lines(i))
        If IsArray(fields) Then
            If UBound(fields) >= 7 Then
                keyId = fields(7)
                If Len(keyId) > 0 Then
                    dict(keyId) = True
                End If
            End If
        End If
ContinueLine:
    Next i

    Set LoadProcessedKeys = dict
End Function

Private Sub AppendLogEntry(ByVal logFilePath As String, ByVal documentName As String, ByVal recordNumber As Long, ByVal outputFilePath As String, ByVal nameFieldValue As String, ByVal logKeyField As String, ByVal logKeyValue As String, ByVal logKeyId As String)
    Dim header As String
    Dim line As String
    Dim content As String
    Dim fileExists As Boolean

    header = "Timestamp,DocumentName,RecordNumber,OutputFilePath,NameFieldValue,LogKeyField,LogKeyValue,KeyId"
    line = CsvEscape(Format$(Now, "yyyy-mm-dd HH:nn:ss")) & "," & _
           CsvEscape(documentName) & "," & _
           CsvEscape(CStr(recordNumber)) & "," & _
           CsvEscape(outputFilePath) & "," & _
           CsvEscape(nameFieldValue) & "," & _
           CsvEscape(logKeyField) & "," & _
           CsvEscape(logKeyValue) & "," & _
           CsvEscape(logKeyId)

    fileExists = (Len(Dir$(logFilePath)) > 0)
    If fileExists Then
        content = ReadTextFileUtf8(logFilePath)
        If Len(content) > 0 And Right$(content, 2) <> vbCrLf Then
            content = content & vbCrLf
        End If
        content = content & line & vbCrLf
    Else
        content = header & vbCrLf & line & vbCrLf
    End If

    WriteTextFileUtf8 logFilePath, content
End Sub

Private Function ReadTextFileUtf8(ByVal filePath As String) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadTextFileUtf8 = stream.ReadText
    stream.Close
End Function

Private Sub WriteTextFileUtf8(ByVal filePath As String, ByVal content As String)
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2
    stream.Close
End Sub

Private Function CsvEscape(ByVal value As String) As String
    Dim escaped As String

    escaped = Replace(value, """", """""")
    CsvEscape = """" & escaped & """"
End Function

Private Function ParseCsvLine(ByVal line As String) As Variant
    Dim fields() As String
    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim current As String
    Dim fieldIndex As Long

    ReDim fields(0)
    fieldIndex = 0
    current = ""
    inQuotes = False

    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If ch = """" Then
            If inQuotes And i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                current = current & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            fields(fieldIndex) = current
            fieldIndex = fieldIndex + 1
            ReDim Preserve fields(fieldIndex)
            current = ""
        Else
            current = current & ch
        End If
    Next i

    fields(fieldIndex) = current
    ParseCsvLine = fields
End Function

Private Function BuildUniqueFilePath(ByVal folderPath As String, ByVal baseName As String, ByVal extension As String, ByVal fso As Object) As String
    Dim candidate As String
    Dim counter As Long
    Dim suffix As String

    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    candidate = folderPath & baseName & "." & extension
    If Not fso.fileExists(candidate) Then
        BuildUniqueFilePath = candidate
        Exit Function
    End If

    counter = 1
    Do
        suffix = "_" & Format$(counter, "000")
        candidate = folderPath & baseName & suffix & "." & extension
        If Not fso.fileExists(candidate) Then
            BuildUniqueFilePath = candidate
            Exit Function
        End If
        counter = counter + 1
    Loop
End Function

Private Function CleanFileName(ByVal rawName As String, ByVal maxLen As Long) As String
    Dim invalidChars As Variant
    Dim ch As Variant

    CleanFileName = Trim$(rawName)
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In invalidChars
        CleanFileName = Replace(CleanFileName, ch, "_")
    Next ch
    CleanFileName = Replace(CleanFileName, vbCr, "")
    CleanFileName = Replace(CleanFileName, vbLf, "")
    CleanFileName = Replace(CleanFileName, vbTab, "")

    If Len(CleanFileName) > maxLen Then
        CleanFileName = Left$(CleanFileName, maxLen)
    End If
    If Len(CleanFileName) = 0 Then
        CleanFileName = "Invoice"
    End If
End Function


