# 毎日のデータ取得をスケジュール実行
try {
    $file = "$env:USERPROFILE\MyWorkSpace\sql_server\mydb_precious_metal_table\preciout_metal_control.xlsm"
    
    # Excelの起動とファイルオープン
    $excel = New-Object -ComObject excel.application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($file)
    
    # プロシージャの実行
    $excel.Run("main.RunSchedule")
}
catch {
    Write-Output $Error[0]
}
finally {
    # 終了処理
    if ($workbook) {
        $workbook.Saveas($file)
        $workbook.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    # ガベージコレクションを強制的に実行
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
