Option Explicit

Const MAXLEN = 27

Dim isCscript : isCscript = (LCase(Right(WScript.FullName, 10)) = "cscript.exe")

Dim args : Set args = WScript.Arguments
If args.Count = 0 Then
  Dim msg
  msg = "使い方：" & vbCrLf & _
        "チェックしたい .txt（または .md）をこの 27check.vbs にドラッグ&ドロップしてください。"
  If isCscript Then
    WScript.Echo msg
  Else
    MsgBox msg, vbInformation, "27文字チェック"
  End If
  WScript.Quit 0
End If

Dim skipped : skipped = 0
Dim exitCode : exitCode = 0

Dim i
For i = 0 To args.Count - 1
  Dim inPath : inPath = args(i)

  ' 拡張子チェック: .txt または .md のみ処理
  If EndsWithLCase(inPath, ".txt") Or EndsWithLCase(inPath, ".md") Then
    ' --- メイン処理開始 ---
    Dim raw, readSuccess
    readSuccess = False

    On Error Resume Next
    raw = ReadUtf8(inPath)
    If Err.Number = 0 Then
      readSuccess = True
    Else
      ReportError "ファイルを読み込めませんでした：" & vbCrLf & inPath & vbCrLf & "（権限や文字コードを確認してください）"
      Err.Clear
      exitCode = 1
    End If
    On Error GoTo 0

    If readSuccess Then
      raw = NormalizeNewlines(raw)

      Dim lines : lines = Split(raw, vbLf, -1, vbBinaryCompare)
      ' 末尾改行由来の最終空行を落とす
      If UBound(lines) >= 0 Then
        If lines(UBound(lines)) = "" Then
          If UBound(lines) = 0 Then
            ReDim lines(-1)
          Else
            ReDim Preserve lines(UBound(lines) - 1)
          End If
        End If
      End If

      Dim totalLines : totalLines = 0
      If IsArray(lines) Then
        On Error Resume Next
        totalLines = UBound(lines) + 1
        If Err.Number <> 0 Then totalLines = 0
        Err.Clear
        On Error GoTo 0
      End If

      Dim rows : rows = ""
      Dim vCount : vCount = 0

      Dim idx
      For idx = 0 To UBound(lines)
        Dim oneLine : oneLine = lines(idx)

        Dim countTarget : countTarget = RStrip(oneLine)
        Dim charLen : charLen = Len(countTarget)

        If charLen > MAXLEN Then
          Dim wordIndex : wordIndex = Int(idx \ 3) + 1
          Dim lineInWord : lineInWord = (idx Mod 3) + 1
          Dim overBy : overBy = charLen - MAXLEN

          Dim snippet : snippet = oneLine
          snippet = Replace(snippet, "|", "｜")
          If Len(snippet) > 80 Then snippet = Left(snippet, 80) & "…"

          rows = rows & "|" & (idx + 1) & "|" & wordIndex & "|" & lineInWord & "|" & charLen & "|" & overBy & "|" & snippet & "|" & vbLf
          vCount = vCount + 1
        End If
      Next

      Dim statusText : statusText = "OK"
      If vCount > 0 Then statusText = "NG"

      Dim outPath : outPath = BuildOutputPath(inPath)

      Dim report : report = ""
      report = report & "# 27文字チェック結果" & vbLf & vbLf
      report = report & "- 対象: " & inPath & vbLf
      report = report & "- 判定: " & statusText & vbLf
      report = report & "- ルール: 1行あたり全角27文字まで（文字数 > 27 を違反）" & vbLf
      report = report & "- 総行数: " & totalLines & vbLf
      report = report & "- 違反数: " & vCount & vbLf & vbLf

      If vCount = 0 Then
        report = report & "問題ありません。" & vbLf
      Else
        report = report & "## 違反一覧" & vbLf & vbLf
        report = report & "|行|ワード|ワード内行|文字数|超過|抜粋|" & vbLf
        report = report & "|---:|---:|---:|---:|---:|---|" & vbLf
        report = report & rows & vbLf
      End If

      Dim writeSuccess : writeSuccess = False
      On Error Resume Next
      WriteUtf8 outPath, report
      If Err.Number = 0 Then
        writeSuccess = True
      Else
        ReportError "結果ファイルを書き込めませんでした：" & vbCrLf & outPath & vbCrLf & "（フォルダ権限や保護設定を確認してください）"
        Err.Clear
        exitCode = 1
      End If
      On Error GoTo 0

      If writeSuccess Then
        If isCscript Then
          WScript.Echo "OK: " & inPath & " -> " & outPath
        Else
          OpenFile outPath
        End If
      End If
    End If
    ' --- メイン処理終了 ---
  Else
    ' .txt/.md 以外はスキップ
    skipped = skipped + 1
  End If
Next

If skipped > 0 Then
  If isCscript Then
    WScript.Echo skipped & "個スキップ（.txt/.mdのみ対応）"
  Else
    MsgBox skipped & "個スキップしました（.txt/.mdのみ対応）", vbInformation, "27文字チェック"
  End If
End If

WScript.Quit exitCode


' ---------- helpers ----------

Sub ReportError(message)
  If isCscript Then
    WScript.Echo "ERROR: " & message
  Else
    MsgBox message, vbCritical, "27文字チェック"
  End If
End Sub

Function EndsWithLCase(s, suffix)
  EndsWithLCase = (LCase(Right(s, Len(suffix))) = suffix)
End Function

Function NormalizeNewlines(t)
  t = Replace(t, vbCrLf, vbLf)
  t = Replace(t, vbCr, vbLf)
  NormalizeNewlines = t
End Function

Function RStrip(t)
  Dim s : s = t
  Do While Len(s) > 0
    Dim c : c = Right(s, 1)
    If c = " " Or c = vbTab Then
      s = Left(s, Len(s) - 1)
    Else
      Exit Do
    End If
  Loop
  RStrip = s
End Function

Function Timestamp()
  Dim d : d = Now
  Timestamp = Year(d) & Pad2(Month(d)) & Pad2(Day(d)) & "_" & Pad2(Hour(d)) & Pad2(Minute(d)) & Pad2(Second(d))
End Function

Function Pad2(n)
  Dim s : s = CStr(n)
  If Len(s) = 1 Then s = "0" & s
  Pad2 = s
End Function

Function BuildOutputPath(inPath)
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  Dim folderPath : folderPath = fso.GetParentFolderName(inPath)
  Dim baseName : baseName = fso.GetBaseName(inPath)
  BuildOutputPath = folderPath & "\" & baseName & "_27check_" & Timestamp() & ".md"
End Function

Function ReadUtf8(path)
  Dim stm : Set stm = CreateObject("ADODB.Stream")
  stm.Type = 2 ' text
  stm.Charset = "utf-8"
  stm.Open
  stm.LoadFromFile path
  ReadUtf8 = stm.ReadText(-1)
  stm.Close
End Function

Sub WriteUtf8(path, textContent)
  ' UTF-8テキストとして書き込み（BOM付き）
  Dim stmUtf8 : Set stmUtf8 = CreateObject("ADODB.Stream")
  stmUtf8.Type = 2 ' text
  stmUtf8.Charset = "utf-8"
  stmUtf8.Open
  stmUtf8.WriteText textContent

  ' バイナリモードに切り替えてBOMをスキップ
  stmUtf8.Position = 0
  stmUtf8.Type = 1 ' binary
  stmUtf8.Position = 3 ' Skip BOM (3 bytes for UTF-8)

  ' BOMなしでファイルに保存
  Dim stmNoBOM : Set stmNoBOM = CreateObject("ADODB.Stream")
  stmNoBOM.Type = 1 ' binary
  stmNoBOM.Open
  stmUtf8.CopyTo stmNoBOM
  stmNoBOM.SaveToFile path, 2 ' adSaveCreateOverWrite

  stmNoBOM.Close
  stmUtf8.Close
End Sub

Sub OpenFile(path)
  On Error Resume Next
  Dim sh : Set sh = CreateObject("Shell.Application")
  sh.ShellExecute path, "", "", "open", 1
  On Error GoTo 0
End Sub
