Attribute VB_Name = "func"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    '32bit版
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

'---------------------------------------------------------------------------------------
' Procedure : CombineDictionaries
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : 2つのDictionaryオブジェクトを結合します。もしorginalDictがNothingの場合、
'             新しいDictionaryオブジェクトが作成されます。addDictのキーがorginalDictに
'             既に存在する場合、そのキーの値は上書きされます。
' Parameters:
'     orginalDict - Dictionaryオブジェクト (ByRef)
'                   結合される元となるDictionary。Nothingの場合、新しいDictionaryが作成されます。
'     addDict     - Dictionaryオブジェクト (ByRef)
'                   orginalDictに結合されるDictionary。
'---------------------------------------------------------------------------------------
Sub CombineDictionaries(ByRef orginalDict As Object, ByRef addDict As Object)
    
    Dim key As Variant
    
    If orginalDict Is Nothing Then
        Set orginalDict = CreateObject("Scripting.Dictionary")
    End If
    
    For Each key In addDict.Keys
        If orginalDict.Exists(key) Then
            ' キーが既に存在する場合、値を上書きまたは結合することができます
            orginalDict(key) = orginalDict(key) + addDict(key)
        Else
            ' キーが存在しない場合、新しいキーと値のペアを追加
            orginalDict.Add key, addDict(key)
        End If
    Next key
    
End Sub

'---------------------------------------------------------------------------------------
' Function  : GetAllFilePaths
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : 指定されたディレクトリ内のすべてのファイルの絶対パスを取得します。
' Parameters:
'     dirPath - String
'               ファイルのパスを取得するディレクトリのパス。
' Returns   : Collection
'               指定されたディレクトリ内のすべてのファイルの絶対パスを含むコレクション。
'---------------------------------------------------------------------------------------
Function GetAllFilePaths(dirPath As String) As Collection
    
    Dim filePath As String
    Dim fileList As Collection
    
    ' パスの終端に"\"がない場合、追加する
    If Right(dirPath, 1) <> "\" Then dirPath = dirPath & "\"
    
    ' ディレクトリ内の最初のファイルの名前を取得
    filePath = Dir(dirPath & "*.*")
    
    ' ファイルリストを格納するコレクションを初期化
    Set fileList = New Collection
    
    ' ディレクトリ内のすべてのファイルのパスを取得
    Do While filePath <> ""
        ' ディレクトリやサブディレクトリを除外
        If (GetAttr(dirPath & filePath) And vbDirectory) = 0 Then
            fileList.Add dirPath & filePath
        End If
        filePath = Dir() ' 次のファイルの名前を取得
    Loop
    
    Set GetAllFilePaths = fileList

End Function

'---------------------------------------------------------------------------------------
' Procedure : FetchAndSaveJSON
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : 指定されたURLからJSONデータを取得し、指定されたファイルパスに保存します。
' Parameters:
'     url          - String
'                    JSONデータを取得するURL。
'     saveFilePath - String
'                    JSONデータを保存するファイルの絶対パス。
'---------------------------------------------------------------------------------------
Sub FetchAndSaveJSON(url As String, saveFilePath As String)

    Dim xhr As Object
    Dim responseText As String
    Dim jsonFile As Integer
    
    On Error GoTo ErrorHandler
    
    ' XMLHTTPオブジェクトの作成
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' URLにPOSTリクエストを送信
    xhr.Open "POST", url, False
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.setOption 2, 13056 '証明書のエラーを無視するオプション
    xhr.send
    
    Call WaitResponse(xhr)
    
    ' レスポンステキストを取得
    responseText = xhr.responseText
    
    ' JSONデータをファイルに保存
    jsonFile = FreeFile
    Open saveFilePath For Output As jsonFile
    Print #jsonFile, responseText
    Close jsonFile
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
    Set xhr = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WaitResponse
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : 指定されたXMLHTTPオブジェクトの応答を待機します。指定されたタイムアウト時間を
'             超えるとエラーを発生させます。
' Parameters:
'     xhr      - Object
'                応答を待機するXMLHTTPオブジェクト。
'     timeout  - Long (Optional)
'                タイムアウトまでの秒数。デフォルトは30秒。
'     waittime - Long (Optional)
'                応答後の追加待機時間（ミリ秒）。デフォルトは3000ミリ秒（3秒）。
'---------------------------------------------------------------------------------------
Sub WaitResponse(xhr As Object, Optional timeout As Long = 30, Optional waittime As Long = 3000)
    
    Dim i As Long
    i = 0
    Do While xhr.readyState < 4
        DoEvents
        Sleep 1000
        i = i + 1
        
        If timeout < i Then
            Err.Raise vbObjectError + 1, "WaitResponse", "タイムアウトエラー: 応答がタイムアウトしました。"
            Exit Sub
        End If
    Loop
    
    Sleep waittime
End Sub

'---------------------------------------------------------------------------------------
' Function  : ConvertJSONtoDict
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : 指定されたJSONファイルの内容を読み取り、Dictionaryオブジェクトとして返します。
' Parameters:
'     jsonFilePath - Variant
'                    JSONデータを含むファイルの絶対パス。
' Returns   : Object
'                    JSONデータを表すDictionaryオブジェクト。
'---------------------------------------------------------------------------------------
Function ConvertJSONtoDict(jsonFilePath As Variant) As Object
        
    Dim jsonText As String
    Dim stream As Object

    On Error GoTo ErrorHandler

    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2 'テキストタイプ
        .Charset = "utf-8" 'UTF-8エンコーディングを指定
        .Open
        .LoadFromFile jsonFilePath
        jsonText = .ReadText
        .Close
    End With
    
    Set stream = Nothing

    ' JSONを解析する
    Set ConvertJSONtoDict = JsonConverter.ParseJson(jsonText)

    Exit Function

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
    Set ConvertJSONtoDict = Nothing

End Function

'---------------------------------------------------------------------------------------
' Function  : ExtractDataFromJson
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : 指定されたJSONデータから特定のデータタイプに関連する情報を抽出し、
'             その結果をコレクションとして返します。
' Parameters:
'     json_data  - Object
'                  解析するJSONデータ。
'     dataType   - String
'                  抽出するデータのタイプ。
'     usePrice   - Boolean
'                  価格情報を使用するかどうかを指定します。
' Returns   : Collection
'                  抽出されたデータを含むコレクション。
'---------------------------------------------------------------------------------------
Function ExtractDataFromJson(json_data As Object, dataType As String, usePrice As Boolean) As Collection
    Dim dt As Variant
    Dim data As Object
    Dim entry As Object
    Dim key As Variant
    Dim result As Collection

    Set result = New Collection

    For Each dt In json_data.Keys
        If TypeName(json_data(dt)) = "Dictionary" Then
            Set data = json_data(dt)
            If data.Count > 0 And data.Exists(dataType) Then
                Set entry = CreateObject("Scripting.Dictionary")
                entry.Add "date", dt
                If usePrice Then
                    For Each key In data(dataType)("ingod").Keys
                        entry.Add key, data(dataType)("ingod")(key)("price")
                    Next key
                Else
                    For Each key In data(dataType).Keys
                        entry.Add key, data(dataType)(key)
                    Next key
                End If
                
                result.Add entry
            End If
        End If
    Next dt

    Set ExtractDataFromJson = result
End Function

'---------------------------------------------------------------------------------------
' Procedure : WriteDataToSheet
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : 指定されたコレクションデータをExcelのワークシートに書き込みます。
' Parameters:
'     collectionData - Collection
'                       書き込むデータを含むコレクション。
'     targetSheet    - Worksheet
'                       データを書き込むターゲットのワークシート。
'---------------------------------------------------------------------------------------
Sub WriteDataToSheet(collectionData As Collection, targetSheet As Worksheet)
    Dim item As Object
    Dim key As Variant
    Dim row As Long, col As Long
    Dim dataArr() As Variant
    
    ' Excelの操作を高速化
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' ヘッダーを書き込む
    Set item = collectionData(1)
    col = 1
    For Each key In item.Keys
        targetSheet.Cells(1, col).value = key
        col = col + 1
    Next key
    
    ' データを配列に格納
    ReDim dataArr(1 To collectionData.Count, 1 To item.Count)
    row = 1
    For Each item In collectionData
        col = 1
        For Each key In item.Keys
            dataArr(row, col) = item(key)
            col = col + 1
        Next key
        row = row + 1
    Next item
    
    ' 配列のデータをワークシートに書き込む
    targetSheet.Range("A2").Resize(UBound(dataArr, 1), UBound(dataArr, 2)).value = dataArr
    
    ' Excelの操作を元に戻す
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InsertDataIntoSQLServer
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : 指定された3つのコレクションデータをSQL Serverの対応するテーブルに挿入します。
' Parameters:
'     market_data - Collection
'                   market_table に挿入するデータのコレクション。
'     nj_buy_data - Collection
'                   nj_buy_table に挿入するデータのコレクション。
'     nj_sell_data - Collection
'                    nj_sell_table に挿入するデータのコレクション。
'---------------------------------------------------------------------------------------

Sub InsertDataIntoSQLServer(market_data As Collection, nj_buy_data As Collection, nj_sell_data As Collection)

    Dim conn As Object
    Dim cmd As Object
    Dim connectionString As String
    
    ' 接続文字列の設定
    connectionString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=MyDB;Trusted_Connection=Yes;"
    
    ' ADODBオブジェクトの初期化
    Set conn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    
    ' データベースに接続
    On Error GoTo ErrorHandler
    conn.Open connectionString
    Set cmd.ActiveConnection = conn

    InsertData cmd, market_data, "InsertIntoMarketTable"
    InsertData cmd, nj_buy_data, "InsertIntoNjBuyTable"
    InsertData cmd, nj_sell_data, "InsertIntoNjSellTable"
    
    ' データベース接続を閉じる
    conn.Close
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
    If conn.State = 1 Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : InsertData
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : 指定されたコレクションデータをSQL Serverのテーブルに挿入します。
'             この関数は、指定されたストアドプロシージャを使用してデータを挿入します。
' Parameters:
'     cmd            - Object
'                      ADODB.Command オブジェクト。データベースへのクエリを実行するために使用されます。
'     dataCollection - Collection
'                      データベースに挿入するデータを含むコレクション。
'     spName         - String
'                      データを挿入するために使用するストアドプロシージャの名前。
'---------------------------------------------------------------------------------------
Sub InsertData(cmd As Object, dataCollection As Collection, spName As String)
    Dim item As Dictionary
    Dim key As Variant
    Dim value As Variant
    Dim param As Object
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = spName

    For Each item In dataCollection
        
        For Each key In item.Keys
            value = item(key)
            
            Select Case key
                Case "date"
                    value = Left(value, 4) & "-" & Mid(value, 5, 2) & "-" & Right(value, 2)
                    cmd.Parameters.Append cmd.CreateParameter("@" & key, adDate, adParamInput, , value)
                    
                Case "price_date", "price_hhmm"
                    cmd.Parameters.Append cmd.CreateParameter("@" & key, adVarChar, adParamInput, Len(value), value)
                    
                Case "ny_end", "tokyo_start", "au_buy_diff", "pt_buy_diff", "au_sell_diff", "pt_sell_diff"
                    cmd.Parameters.Append cmd.CreateParameter("@" & key, adInteger, adParamInput, , value)
                
                Case "au_ny_end", "pt_ny_end", "ag_ny_end", "ny_exchange_rate", "tokyo_exchange_rate", "au_buy", "pt_buy", "au_sell", "pt_sell", "au", "ag", "pt", "pd"
                    If InStr(1, value, ",") > 0 And IsNumeric(value) Then
                        value = Replace(value, ",", "")
                    End If
                    Set param = cmd.CreateParameter("@" & key, adNumeric, adParamInput)
                    param.Precision = 10
                    param.NumericScale = 2
                    param.value = CDbl(value)
                    cmd.Parameters.Append param
            End Select
        Next key
        
        cmd.Execute
        
        ' パラメータをクリア
        Dim i As Long
        For i = cmd.Parameters.Count - 1 To 0 Step -1
            cmd.Parameters.Delete cmd.Parameters(i).Name
        Next i

    Next item
End Sub


