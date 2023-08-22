Attribute VB_Name = "main"
Option Explicit

Private Const NETJAPANBASEURL As String = "https://www.net-japan.co.jp/system/upload/netjapan/export/price_"

Sub main()
    
    Dim jsonFolderPath As String
    Dim startDate As Date, endDate As Date, currentMonth As Date
    Dim transDate As Long, url As String, saveFilePath As String
    Dim jsonFileList As Collection, json_data As Object, add_json_data As Object
    Dim market_data As Collection, nj_buy_data As Collection, nj_sell_data As Collection
    Dim file As Variant
    
    ' 初期設定
    jsonFolderPath = ThisWorkbook.Path & "\json\"
    Set json_data = CreateObject("Scripting.Dictionary")
    
    ' ###########################################
    ' ### 1. 対象年月Jsonファイルをすべて取得 ###
    ' ###########################################
    With ThisWorkbook.Sheets("control")
        startDate = DateSerial(.Range("B2").value, .Range("C2").value, 1)
        endDate = DateSerial(.Range("B3").value, .Range("C3").value, 1)
    End With

    currentMonth = startDate
    Do While currentMonth <= endDate
        transDate = Format(currentMonth, "yyyymm")
        url = NETJAPANBASEURL & transDate & ".json"
        saveFilePath = jsonFolderPath & "price_" & transDate & ".json"
        Call FetchAndSaveJSON(url, saveFilePath)
        currentMonth = DateAdd("m", 1, currentMonth)
    Loop
    
    ' #############################################
    ' ### 2. JsonデータをDictionaryに変換し結合 ###
    ' #############################################
    Set jsonFileList = GetAllFilePaths(jsonFolderPath)
    For Each file In jsonFileList
        Set add_json_data = ConvertJSONtoDict(file)
        Call CombineDictionaries(json_data, add_json_data)
    Next file
    
    ' #####################################
    ' ### 3. JsonデータをDBへ挿入(または転記) ###
    ' #####################################
    Set market_data = ExtractDataFromJson(json_data, "market", False)
    Set nj_buy_data = ExtractDataFromJson(json_data, "nj_buy", True)
    Set nj_sell_data = ExtractDataFromJson(json_data, "nj_sell", True)
    
    InsertDataIntoSQLServer market_data, nj_buy_data, nj_sell_data
    
    ' 転記する場合以下を利用する
'    WriteDataToSheet market_data, ThisWorkbook.Sheets("MarketData")
'    WriteDataToSheet nj_buy_data, ThisWorkbook.Sheets("NjBuyData")
'    WriteDataToSheet nj_sell_data, ThisWorkbook.Sheets("NjSellData")
    
    ThisWorkbook.Save
    
End Sub

Sub RunSchedule()
    
    Dim transDate As Long, url As String, saveFilePath As String
    Dim json_data As Object
    Dim market_data As Collection, nj_buy_data As Collection, nj_sell_data As Collection
    
    transDate = Format(Date, "yyyymm")
    url = NETJAPANBASEURL & transDate & ".json"
    saveFilePath = jsonFolderPath & "price_" & transDate & ".json"
    
    Call FetchAndSaveJSON(url, saveFilePath)
    Set json_data = ConvertJSONtoDict(saveFilePath)
    
    json_data = json_data(Str(transDate))
    
    Set market_data = ExtractDataFromJson(json_data, "market", False)
    Set nj_buy_data = ExtractDataFromJson(json_data, "nj_buy", True)
    Set nj_sell_data = ExtractDataFromJson(json_data, "nj_sell", True)
    
    InsertDataIntoSQLServer market_data, nj_buy_data, nj_sell_data
    
End Sub

