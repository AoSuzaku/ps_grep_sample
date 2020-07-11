﻿#************************************************************
#　GrepTool
#
#　変更履歴
#　　・2020/07/11　新規作成
#
#************************************************************

echo "**************************************************"
echo "**　　　　　　　　　GrepTool　　　　　　　　　　**"
echo "**************************************************"

# ps1ファイルの配置パス
$dir = Split-Path $myInvocation.MyCommand.Path -Parent

$path = Read-Host "検索パスを入力してください。"

if($path -eq ""){

    echo "検索パスが入力されていません。"
    return

}

$word = Read-Host "検索ワードを入力してください。"

if($word -eq ""){

    echo "検索ワードが入力されていません。"
    return

}

echo "検索パス：$path"
echo "検索ワード：$word"

# Excelオブジェクト生成
$excel = New-Object -ComObject Excel.Application

# Excelオブジェクト設定
$excel.DisplayAlerts = $false

# Grep検索処理
Get-ChildItem $path -Recurse -Include "*.xls*" -Name | % {

    # サブフォルダ配下のパス
    $childPath = $_

    # Excelブック　Open
    $wb = $excel.Workbooks.Open("$path\$childPath")

    # シート毎に検索を実施
    $wb.Worksheets | % {
    
        $ws = $_
        $wsName = $ws.Name
        $first = $result = $ws.Cells.Find($word)

        while($result -ne $null){

            echo "$path\$childPath：$wsName`t$($result.Row), $($result.Column)`t$($result.Text)" | 
                Out-File -Append "$dir\result.txt"

            $result = $ws.Cells.FindNext($result)

            if($result.Address() -eq $first.Address()){

                break

            }

        }

    } 

    # Excelブック　Close
    $wb.Close(0)

}

# 初期化
$excel.Quit()
$ws = $null
$wb = $null
$excel = $null