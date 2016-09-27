Option Explicit

' VBSQLipt : 対話式 SQL 実行スクリプト
' 
' IE をコンソール画面として使用し、DB 接続し SQL を実行する。
' Oracle DB に接続するものとして作成したが、ConnectionString を変更すれば
' ODBC 接続が可能なので、その他の DBMS にも使用可能。

' ユーザ設定する項目 (DB 接続先情報)
Const HOST     = "127.0.0.1"
Const PORT     = "1521"
Const SERVICE  = "NeosService"
Const USER_ID  = "Neos21"
Const PASSWORD = "NeosPassword"



' 各処理を呼び出す
Call Confirmation()
Call OpenIE()
Call Main()
Call CloseIE()
WScript.Quit

' レコード出力時のフィールド区切り文字
Const DELIMITER = " , "

' コンソール画面として使用する IE を保持する
Dim PIE

' 実行前確認のダイアログを表示する
Sub Confirmation()
  Dim shell
  Set shell = CreateObject("WScript.Shell")
  If shell.Popup("実行しますか？", 0, "実行確認", vbOKCancel + vbQuestion) = vbCancel Then
    MsgBox "終了します。"
    Set shell = Nothing
    WScript.Quit
  End If
  Set shell = Nothing
End Sub

' コンソール画面として使用する IE を起動し、表示仕様を設定する
Sub OpenIE()
  Set PIE = CreateObject("InternetExplorer.Application")
  With PIE
    .Navigate "about:blank"
    .ToolBar = False
    .StatusBar = False
    .Width = .Document.parentWindow.screen.availWidth
    .Height = .Document.parentWindow.screen.availHeight
    .Top = 0
    .Left = 0
    .Visible = True
    .Document.Title = "VBSQLipt"
    .Document.Body.style.color = "#0c0"
    .Document.Body.style.background = "#000"
    ' 等幅フォント
    .Document.Body.style.fontFamily = "'ＭＳ ゴシック', monospace"
  End With
End Sub

' コンソール画面のメイン処理
Sub Main()
  On Error Resume Next
  
  Dim con
  Set con = CreateObject("ADODB.Connection")
  
  ' ODBC に登録してあるデータソースを使用する場合の書き方：
  '   conStr = "Provider=MSDASQL.1;Password=NeosPassword;Persist Security Info=True;User ID=Neos21;Data Source=NeosDataSource"
  ' 同じく ODBC に登録してあるデータソースを DSN で参照する場合の書き方：
  '   conStr = "DSN=NeosDataSource;UID=Neos21;PWD=NeosPassword;DBQ=NeosService;"
  ' 以下は tnsnames.ora に記載する接続文字列で接続する方法
  Dim conStr
  conStr = "Provider=OraOLEDB.Oracle;" & _
           "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & HOST & ")(PORT=" & PORT & ")))(CONNECT_DATA=(SERVICE_NAME=" & SERVICE & ")));" & _
           "User Id=" & USER_ID & ";Password=" & PASSWORD & ";"
  IEMsg("接続文字列：" & conStr & "<br>")
  
  con.ConnectionString = conStr
  con.Open
  
  If Err.Number <> 0 Then
    IEMsg("DB 接続失敗" & "<br>" & Err.Number & "<br>" & Err.Source & "<br>" & Err.Description)
    Exit Sub
  Else
    IEMsg("DB 接続成功...")
  End If
  
  ' クライアントサイドカーソルに変更する
  ' rs.RecordCount を有効にするための設定。これをしないと RecordCount が -1 になる
  con.CursorLocation = 3  ' adUseClient
  
  ' 対話式スクリプト開始
  Do While True
    IEMsg("<br>")
    
    ' SQL をユーザに入力させる
    Dim sql
    sql = SqlPrompt
    
    ' 入力文字列が exit や quit なら Do While True を抜けて終了する
    Select Case Lcase(sql)
      Case "exit"
        Exit Do
      Case "quit"
        Exit Do
    End Select
    
    ' 入力値で SQL を実行し、Recordset オブジェクトに結果を格納する
    IEMsg("&gt; " & sql & "<br>")
    Dim rs
    Set rs = con.Execute(sql)
    
    If Err.Number <> 0 Then
      IEMsg("SQL 実行失敗" & "<br>" & Err.Number & "<br>" & Err.Source & "<br>" & Err.Description)
    Else
      IEMsg("SQL 実行成功...")
      ' 結果出力処理
      printResult(rs)
    End If
    
    ' Recordset をクローズする
    rs.Close
    Set rs = Nothing
    
    ' エラーをクリアする
    Err.Clear
  Loop
  
  con.Close
  Set con = Nothing
  IEMsg("DB 接続切断")
End Sub

' SQL をユーザに入力させる
private Function SqlPrompt()
  Dim input
  
  ' 何かしら値が入力されるまでループする (空欄や「キャンセル」ボタンの押下などを許容しない)
  Do While True
    input = InputBox("SQL を入力してください。終了するときは「exit」か「quit」と入力してください。", "Prompt")
    
    If Trim(input) <> "" Then
      Exit Do
    End If
  Loop
  
  SqlPrompt = Trim(input)
End Function

' SQL 実行結果を出力する
private Sub printResult(rs)
  ' 結果件数を取得する
  Dim cnt
  cnt = rs.RecordCount
  
  If cnt = 0 Then
    ' 0件なら結果出力なし
    IEMsg("結果：0件 … ヒットしませんでした")
    Exit Sub
  End If
  
  ' 結果件数が -1 かそれ以外のときは結果出力する
  If cnt = -1 Then
    ' カーソルサービスの場所をクライアントサイドに設定できていない場合 (と思われる)
    IEMsg("件数がうまく取得できませんでした。" & "<br>")
  Else
    IEMsg(cnt & " 件のレコードを取得しました。" & "<br>")
  End If
  
  Dim shell
  Set shell = CreateObject("WScript.Shell")
  Select Case shell.Popup("結果を出力しますか？", 0, "結果出力確認", vbOKCancel + vbQuestion)
    Case vbOK
      ' Recordset の出力処理
      printRecordset(rs)
    Case vbCancel
      IEMsg("結果出力がキャンセルされました")
  End Select
  Set shell = Nothing
End Sub

' Recordset を出力する
private Sub printRecordset(rs)
  ' フィールド名を出力する
  Dim headerStr
  headerStr = ""
  
  ' 結果1行目からフィールド名を取得する
  ' (結果件数が0件の場合に呼び出してもフィールド名は取得可能)
  Dim fieldName
  For Each fieldName In rs.Fields
    headerStr = headerStr & fieldName.Name & DELIMITER
  Next
  
  ' 行末にある区切り文字を除去する
  If Right(headerStr, 3) = DELIMITER Then
    headerStr = Left(headerStr, Len(headerStr) - Len(DELIMITER))
  End If
  
  IEMsg(headerStr)
  
  ' 1行ずつ出力する
  Do Until rs.EOF = True
    Dim recordStr
    recordStr = ""
    
    Dim field
    For Each field In rs.Fields
      recordStr = recordStr & field.Value & DELIMITER
    Next
    
    If Right(recordStr, 3) = DELIMITER Then
      recordStr = Left(recordStr, Len(recordStr) - Len(DELIMITER))
    End If
    
    IEMsg(recordStr)
    rs.MoveNext
  Loop
End Sub

' IE にメッセージを出力する
Sub IEMsg(val)
  With PIE
    .Document.Body.innerHTML = .Document.Body.innerHTML & val & "<br>"
    .Document.Script.setTimeout "javascript:scrollTo(0, " & .Document.Body.ScrollHeight & ");", 0
  End With
End Sub

' IE を閉じる
Sub CloseIE()
  MsgBox "終了します"
  PIE.Quit
  Set PIE = Nothing
End Sub