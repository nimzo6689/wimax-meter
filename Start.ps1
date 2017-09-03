<#
author     : nizmo6689
version    : 0.0.1

WiMAXの通信料自動取得スクリプト

必要となる環境設定：
  PowerShell: 5.0以上
    管理者権限で「Set-ExecutionPolicy RemoteSigned」を実行しておくこと。
    バージョンの確認方法：$PSVersionTable
  ブラウザ: Chromeがインストール済みであること。
  Selenium: 3.0以上（WebDriver.dll, chromedriver.exe）

タスクスケジューラへの登録方法
$SchedulerParams = @{
    Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "$(AbsolutePath)\Start.ps1"
    Trigger = New-ScheduledTaskTrigger -DaysInterval 1 -Daily -At "00:00 AM"
    Settings = New-ScheduledTaskSettingsSet -Hidden -StartWhenAvailable
}
Register-ScheduledTask -TaskPath \ -TaskName WiMAX-Meta @SchedulerParams -Force
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = "stop"

$settings = . $(Join-Path $PSScriptRoot _Settings.ps1)
. $(Join-Path $PSScriptRoot Libs | Join-Path -ChildPath _XPathLogic.ps1)

Initialize -webDriver $settings.Selenium.WebDriver -chromeDriver $settings.Selenium.ChromeDriver

# 「Speed Wi-Fi NEXT」へアクセスし、通信量を取得する。
Login -userType $settings.Auth.UserType -encodedPass $settings.Auth.EncodedPass
$usedData = GetUsedData
CloseWebDriver

# csvへの算出方式に従って出力、または、上書きする。
$csvFullName = Join-Path $settings.CsvDir ("wimax_" + [datetime]::Today.ToString("yyyyMM") + ".csv")
if (Test-Path $csvFullName) {
    [array]$totalData = Get-Content $csvFullName | ConvertFrom-Csv
    $count = $totalData.Length
    if ($totalData[$count - 1].day -ne $usedData.day) {
        $usedData | Add-Member daily ($usedData.inMonth - $totalData[$count - 1].inMonth)
        $totalData += $usedData
        $totalData | ConvertTo-Csv | Select-Object -Skip 1 | Out-File $csvFullName -Force
    }
}
else {
    if ([datetime]::Today.ToString("dd") -eq "01") {
        # 月初の処理
        $csvFullNamePreMonth = Join-Path $settings.CsvDir ("wimax_" + [datetime]::Today.AddDays(-1).ToString("yyyyMM") + ".csv")
        $totalData = Get-Content $csvFullNamePreMonth | ConvertFrom-Csv
        $totalData += $usedData
        $count = $totalData.Length
        $totalData[$count - 1].inMonth = $usedData.toToday - $totalData[$count - 2].daily
        $totalData[$count - 1].daily = $totalData[$count - 1].inMonth - $totalData[$count - 2].inMonth
        $totalData | ConvertTo-Csv | Select-Object -Skip 1 | Out-File $csvFullNamePreMonth -Force
    }
    else {
        # 2日の処理
        $usedData | Add-Member daily $usedData.inMonth
        $usedData | ConvertTo-Csv | Select-Object -Skip 1 | Select-Object -Last 2 | Out-File $csvFullName
    }
}