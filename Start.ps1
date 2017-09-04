<#
author     : nizmo6689
version    : 0.1.0

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
Get-ChildItem $(Join-Path $PSScriptRoot Libs) -Include *.ps1 -Recurse -Force | ForEach-Object {. $_.FullName}

. $(Join-Path $PSScriptRoot Libs | Join-Path -ChildPath _XPathLogic.ps1)
. $(Join-Path $PSScriptRoot Libs | Join-Path -ChildPath _CsvLogic.ps1)

Initialize -webDriver $settings.Selenium.WebDriver -chromeDriver $settings.Selenium.ChromeDriver

# 「Speed Wi-Fi NEXT」へアクセスし、通信量を取得する。
Login -userType $settings.Auth.UserType -encodedPass $settings.Auth.EncodedPass
$usedData = GetUsedData
CloseWebDriver

# csvファイルへ使用量を書き込む。
WriteCsvWithUsedData -csvDir $settings.CsvDir -usedData $usedData