using namespace System.Management.Automation

Set-StrictMode -Version Latest
$ErrorActionPreference = "stop"

if ([CommandOrigin]::Runspace -eq $MyInvocation.CommandOrigin) {
    throw "This script shouldn't be run with new runspace."
}

# 今月の集計用csvファイルに前日の通信量を追記する。
function AppendUsedData($csvFullName, $usedData) {
    [array]$totalData = Get-Content $csvFullName | ConvertFrom-Csv
    $count = $totalData.Length

    $usedData | Add-Member daily ($usedData.inMonth - $totalData[$count - 1].inMonth)
    $totalData += $usedData
    $totalData | ConvertTo-Csv | Select-Object -Skip 1 | Out-File $csvFullName -Force
}

# 先月分のファイルに対して月末のデータを追記する。
function AppendUsedDataToPreMonth($csvFullNamePreMonth, $usedData) {
    [array]$totalData = Get-Content $csvFullNamePreMonth | ConvertFrom-Csv
    $totalData += $usedData
    $count = $totalData.Length
    $totalData[$count - 1].inMonth = $usedData.toToday - $totalData[$count - 2].daily
    $totalData[$count - 1].daily = $totalData[$count - 1].inMonth - $totalData[$count - 2].inMonth
    $totalData | ConvertTo-Csv | Select-Object -Skip 1 | Out-File $csvFullNamePreMonth -Force
}

# csvへの算出方式に従って出力する。
function WriteCsvWithUsedData($csvDir, $usedData) {
    $csvFullName = Join-Path $csvDir ("wimax_" + [datetime]::Today.ToString("yyyyMM") + ".csv")
    if (!(Test-Path $csvFullName)) {
        if ([datetime]::Today.ToString("dd") -eq "01") {
            # 1日の処理
            $csvFullNamePreMonth = Join-Path $settings.CsvDir ("wimax_" + [datetime]::Today.AddDays(-1).ToString("yyyyMM") + ".csv")
            if (!(Test-Path $csvFullNamePreMonth)) {
                # 先月分のファイルがない場合は先月末日の一日の使用量は算出しない。
                $usedData | ConvertTo-Csv | Select-Object -Skip 1 | Out-File $csvFullNamePreMonth
            }
            else {
                AppendUsedDataToPreMonth -csvFullNamePreMonth $csvFullNamePreMonth -usedData $usedData
            }
        }
        else {
            # 2日、または、月の途中から集計を開始する場合
            if ([datetime]::Today.ToString("dd") -eq "02") {
                # 2日実行の場合、前日までの月累計が1日の使用量と等しくなる。
                $usedData | Add-Member daily $usedData.inMonth
            }
            # 月の途中から集計する場合、開始日の一日の使用量は空白（不明）となる。
            $usedData | ConvertTo-Csv | Select-Object -Skip 1 | Out-File $csvFullName
        }
    }
    else {
        AppendUsedData -csvFullName $csvFullName -usedData $usedData
    }
}