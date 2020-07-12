#************************************************************
#　GrepTool
#
#　変更履歴
#　　・2020/07/11　新規作成
#
#************************************************************

Write-Host "**************************************************"
Write-Host "**　　　　　　　　　GrepTool　　　　　　　　　　**"
Write-Host "**************************************************"

# ps1ファイルの配置パス
$dir = Split-Path $myInvocation.MyCommand.Path -Parent

# 起動引数確認
if($args[0] -eq $null -or $args[1] -eq $null){

    Write-Host "起動引数が設定されていないため、処理を終了します。"
    return

}else{

    $path = $args[0]
    $word = $args[1]

}

if($path -eq ""){

    echo "検索パスが入力されていません。"
    return

}elseif(!(Test-Path $path)){

    echo "検索パスが存在しません。"
    return

}

if($word -eq ""){

    echo "検索ワードが入力されていません。"
    return

}

echo "検索パス：$path"
echo "検索ワード：$word"

# Excelオブジェクト生成
$excel = New-Object -ComObject Excel.Application

# Excelオブジェクト設定
$excel.Visible = $false
$excel.DisplayAlerts = $false

# 配列宣言
$msg = @()
$errMsg = @()

# 実行時間計測 開始
$watch = New-Object System.Diagnostics.Stopwatch
$watch.Start()

# ファイル数カウント
$total = (Get-ChildItem $path -Recurse -Include "*.xls*" -Name | Measure-Object).Count

# Grep検索処理
Get-ChildItem $path -Recurse -Include "*.xls*" -Name | % {

   try{

        # 処理カウント
        $cnt += 1
        $status = "{0}／$total 件処理中" -F $cnt
        Write-Progress $status -PercentComplete $cnt -CurrentOperation $currentOperation

        # サブフォルダ配下のパス
        $childPath = $_

        # Excelブック　Open
        $wb = $excel.Workbooks.Open("$path\$childPath", $false, $true, [Type]::Missing, $null)

        # シート毎に検索を実施
        $wb.Worksheets | % {
    
            $ws = $_
            $wsName = $ws.Name
            $first = $result = $ws.Cells.Find($word)

            while($result -ne $null){

                $msg += "$path\$childPath：$wsName" + "`t" + "$($result.Row), $($result.Column)" + "`t" + "$($result.Text)"

                $result = $ws.Cells.FindNext($result)

                if($result.Address() -eq $first.Address()){

                    break

                }

            }

        }

        # Excelブック　Close
        $wb.Close(0)

    }catch{
    
        # エラー発生処理
        $errMsg += "$path\$childPath" + "`t" + "ErrMsg：" + $_.Exception.Message

    }

}

# 出力処理
try{

    # Grep結果出力
    if($msg.Length -gt 0){

        echo "ファイル名（絶対パス）：シート名`tRow,Column`tText" | Out-File -Append "$dir\result.txt"
        echo $msg | Out-File -Append "$dir\result.txt"

    }else{
    
        echo "検索結果は0件です。" | Out-File -Append "$dir\result.txt"
    
    }
    

    # エラー結果出力
    if($errMsg -gt 0){

        echo $errMsg | Out-File -Append "$dir\errLog.txt"

    }

}catch{

    Write-Host "出力処理でエラーが発生しました。"
    Write-Host "ErrMsg："$_.Exception.Message

}finally{

    # 実行時間計測 終了
    $watch.Stop()
    $time = $watch.Elapsed

    Write-Host "実行時間："$time.TotalSeconds.ToString("0.000")"sec"

    # メモリ開放
    $excel.Quit()
    $ws = $null
    $wb = $null
    $excel = $null
    $time = $null

    [GC]::Collect()

}
