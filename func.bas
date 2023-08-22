Attribute VB_Name = "func"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    '32bit��
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

'---------------------------------------------------------------------------------------
' Procedure : CombineDictionaries
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : 2��Dictionary�I�u�W�F�N�g���������܂��B����orginalDict��Nothing�̏ꍇ�A
'             �V����Dictionary�I�u�W�F�N�g���쐬����܂��BaddDict�̃L�[��orginalDict��
'             ���ɑ��݂���ꍇ�A���̃L�[�̒l�͏㏑������܂��B
' Parameters:
'     orginalDict - Dictionary�I�u�W�F�N�g (ByRef)
'                   ��������錳�ƂȂ�Dictionary�BNothing�̏ꍇ�A�V����Dictionary���쐬����܂��B
'     addDict     - Dictionary�I�u�W�F�N�g (ByRef)
'                   orginalDict�Ɍ��������Dictionary�B
'---------------------------------------------------------------------------------------
Sub CombineDictionaries(ByRef orginalDict As Object, ByRef addDict As Object)
    
    Dim key As Variant
    
    If orginalDict Is Nothing Then
        Set orginalDict = CreateObject("Scripting.Dictionary")
    End If
    
    For Each key In addDict.Keys
        If orginalDict.Exists(key) Then
            ' �L�[�����ɑ��݂���ꍇ�A�l���㏑���܂��͌������邱�Ƃ��ł��܂�
            orginalDict(key) = orginalDict(key) + addDict(key)
        Else
            ' �L�[�����݂��Ȃ��ꍇ�A�V�����L�[�ƒl�̃y�A��ǉ�
            orginalDict.Add key, addDict(key)
        End If
    Next key
    
End Sub

'---------------------------------------------------------------------------------------
' Function  : GetAllFilePaths
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : �w�肳�ꂽ�f�B���N�g�����̂��ׂẴt�@�C���̐�΃p�X���擾���܂��B
' Parameters:
'     dirPath - String
'               �t�@�C���̃p�X���擾����f�B���N�g���̃p�X�B
' Returns   : Collection
'               �w�肳�ꂽ�f�B���N�g�����̂��ׂẴt�@�C���̐�΃p�X���܂ރR���N�V�����B
'---------------------------------------------------------------------------------------
Function GetAllFilePaths(dirPath As String) As Collection
    
    Dim filePath As String
    Dim fileList As Collection
    
    ' �p�X�̏I�[��"\"���Ȃ��ꍇ�A�ǉ�����
    If Right(dirPath, 1) <> "\" Then dirPath = dirPath & "\"
    
    ' �f�B���N�g�����̍ŏ��̃t�@�C���̖��O���擾
    filePath = Dir(dirPath & "*.*")
    
    ' �t�@�C�����X�g���i�[����R���N�V������������
    Set fileList = New Collection
    
    ' �f�B���N�g�����̂��ׂẴt�@�C���̃p�X���擾
    Do While filePath <> ""
        ' �f�B���N�g����T�u�f�B���N�g�������O
        If (GetAttr(dirPath & filePath) And vbDirectory) = 0 Then
            fileList.Add dirPath & filePath
        End If
        filePath = Dir() ' ���̃t�@�C���̖��O���擾
    Loop
    
    Set GetAllFilePaths = fileList

End Function

'---------------------------------------------------------------------------------------
' Procedure : FetchAndSaveJSON
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : �w�肳�ꂽURL����JSON�f�[�^���擾���A�w�肳�ꂽ�t�@�C���p�X�ɕۑ����܂��B
' Parameters:
'     url          - String
'                    JSON�f�[�^���擾����URL�B
'     saveFilePath - String
'                    JSON�f�[�^��ۑ�����t�@�C���̐�΃p�X�B
'---------------------------------------------------------------------------------------
Sub FetchAndSaveJSON(url As String, saveFilePath As String)

    Dim xhr As Object
    Dim responseText As String
    Dim jsonFile As Integer
    
    On Error GoTo ErrorHandler
    
    ' XMLHTTP�I�u�W�F�N�g�̍쐬
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' URL��POST���N�G�X�g�𑗐M
    xhr.Open "POST", url, False
    xhr.setRequestHeader "Content-Type", "application/json"
    xhr.setOption 2, 13056 '�ؖ����̃G���[�𖳎�����I�v�V����
    xhr.send
    
    Call WaitResponse(xhr)
    
    ' ���X�|���X�e�L�X�g���擾
    responseText = xhr.responseText
    
    ' JSON�f�[�^���t�@�C���ɕۑ�
    jsonFile = FreeFile
    Open saveFilePath For Output As jsonFile
    Print #jsonFile, responseText
    Close jsonFile
    
    Exit Sub
    
ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical, "�G���["
    Set xhr = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WaitResponse
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : �w�肳�ꂽXMLHTTP�I�u�W�F�N�g�̉�����ҋ@���܂��B�w�肳�ꂽ�^�C���A�E�g���Ԃ�
'             ������ƃG���[�𔭐������܂��B
' Parameters:
'     xhr      - Object
'                ������ҋ@����XMLHTTP�I�u�W�F�N�g�B
'     timeout  - Long (Optional)
'                �^�C���A�E�g�܂ł̕b���B�f�t�H���g��30�b�B
'     waittime - Long (Optional)
'                ������̒ǉ��ҋ@���ԁi�~���b�j�B�f�t�H���g��3000�~���b�i3�b�j�B
'---------------------------------------------------------------------------------------
Sub WaitResponse(xhr As Object, Optional timeout As Long = 30, Optional waittime As Long = 3000)
    
    Dim i As Long
    i = 0
    Do While xhr.readyState < 4
        DoEvents
        Sleep 1000
        i = i + 1
        
        If timeout < i Then
            Err.Raise vbObjectError + 1, "WaitResponse", "�^�C���A�E�g�G���[: �������^�C���A�E�g���܂����B"
            Exit Sub
        End If
    Loop
    
    Sleep waittime
End Sub

'---------------------------------------------------------------------------------------
' Function  : ConvertJSONtoDict
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : �w�肳�ꂽJSON�t�@�C���̓��e��ǂݎ��ADictionary�I�u�W�F�N�g�Ƃ��ĕԂ��܂��B
' Parameters:
'     jsonFilePath - Variant
'                    JSON�f�[�^���܂ރt�@�C���̐�΃p�X�B
' Returns   : Object
'                    JSON�f�[�^��\��Dictionary�I�u�W�F�N�g�B
'---------------------------------------------------------------------------------------
Function ConvertJSONtoDict(jsonFilePath As Variant) As Object
        
    Dim jsonText As String
    Dim stream As Object

    On Error GoTo ErrorHandler

    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2 '�e�L�X�g�^�C�v
        .Charset = "utf-8" 'UTF-8�G���R�[�f�B���O���w��
        .Open
        .LoadFromFile jsonFilePath
        jsonText = .ReadText
        .Close
    End With
    
    Set stream = Nothing

    ' JSON����͂���
    Set ConvertJSONtoDict = JsonConverter.ParseJson(jsonText)

    Exit Function

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical, "�G���["
    Set ConvertJSONtoDict = Nothing

End Function

'---------------------------------------------------------------------------------------
' Function  : ExtractDataFromJson
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : �w�肳�ꂽJSON�f�[�^�������̃f�[�^�^�C�v�Ɋ֘A������𒊏o���A
'             ���̌��ʂ��R���N�V�����Ƃ��ĕԂ��܂��B
' Parameters:
'     json_data  - Object
'                  ��͂���JSON�f�[�^�B
'     dataType   - String
'                  ���o����f�[�^�̃^�C�v�B
'     usePrice   - Boolean
'                  ���i�����g�p���邩�ǂ������w�肵�܂��B
' Returns   : Collection
'                  ���o���ꂽ�f�[�^���܂ރR���N�V�����B
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
' Purpose   : �w�肳�ꂽ�R���N�V�����f�[�^��Excel�̃��[�N�V�[�g�ɏ������݂܂��B
' Parameters:
'     collectionData - Collection
'                       �������ރf�[�^���܂ރR���N�V�����B
'     targetSheet    - Worksheet
'                       �f�[�^���������ރ^�[�Q�b�g�̃��[�N�V�[�g�B
'---------------------------------------------------------------------------------------
Sub WriteDataToSheet(collectionData As Collection, targetSheet As Worksheet)
    Dim item As Object
    Dim key As Variant
    Dim row As Long, col As Long
    Dim dataArr() As Variant
    
    ' Excel�̑����������
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' �w�b�_�[����������
    Set item = collectionData(1)
    col = 1
    For Each key In item.Keys
        targetSheet.Cells(1, col).value = key
        col = col + 1
    Next key
    
    ' �f�[�^��z��Ɋi�[
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
    
    ' �z��̃f�[�^�����[�N�V�[�g�ɏ�������
    targetSheet.Range("A2").Resize(UBound(dataArr, 1), UBound(dataArr, 2)).value = dataArr
    
    ' Excel�̑�������ɖ߂�
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InsertDataIntoSQLServer
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : �w�肳�ꂽ3�̃R���N�V�����f�[�^��SQL Server�̑Ή�����e�[�u���ɑ}�����܂��B
' Parameters:
'     market_data - Collection
'                   market_table �ɑ}������f�[�^�̃R���N�V�����B
'     nj_buy_data - Collection
'                   nj_buy_table �ɑ}������f�[�^�̃R���N�V�����B
'     nj_sell_data - Collection
'                    nj_sell_table �ɑ}������f�[�^�̃R���N�V�����B
'---------------------------------------------------------------------------------------

Sub InsertDataIntoSQLServer(market_data As Collection, nj_buy_data As Collection, nj_sell_data As Collection)

    Dim conn As Object
    Dim cmd As Object
    Dim connectionString As String
    
    ' �ڑ�������̐ݒ�
    connectionString = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=MyDB;Trusted_Connection=Yes;"
    
    ' ADODB�I�u�W�F�N�g�̏�����
    Set conn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    
    ' �f�[�^�x�[�X�ɐڑ�
    On Error GoTo ErrorHandler
    conn.Open connectionString
    Set cmd.ActiveConnection = conn

    InsertData cmd, market_data, "InsertIntoMarketTable"
    InsertData cmd, nj_buy_data, "InsertIntoNjBuyTable"
    InsertData cmd, nj_sell_data, "InsertIntoNjSellTable"
    
    ' �f�[�^�x�[�X�ڑ������
    conn.Close
    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂���: " & Err.Description, vbCritical, "�G���["
    If conn.State = 1 Then conn.Close
    Set cmd = Nothing
    Set conn = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : InsertData
' Author    : kakeru suzuki
' Date      : 2023/08/20
' Purpose   : �w�肳�ꂽ�R���N�V�����f�[�^��SQL Server�̃e�[�u���ɑ}�����܂��B
'             ���̊֐��́A�w�肳�ꂽ�X�g�A�h�v���V�[�W�����g�p���ăf�[�^��}�����܂��B
' Parameters:
'     cmd            - Object
'                      ADODB.Command �I�u�W�F�N�g�B�f�[�^�x�[�X�ւ̃N�G�������s���邽�߂Ɏg�p����܂��B
'     dataCollection - Collection
'                      �f�[�^�x�[�X�ɑ}������f�[�^���܂ރR���N�V�����B
'     spName         - String
'                      �f�[�^��}�����邽�߂Ɏg�p����X�g�A�h�v���V�[�W���̖��O�B
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
        
        ' �p�����[�^���N���A
        Dim i As Long
        For i = cmd.Parameters.Count - 1 To 0 Step -1
            cmd.Parameters.Delete cmd.Parameters(i).Name
        Next i

    Next item
End Sub


