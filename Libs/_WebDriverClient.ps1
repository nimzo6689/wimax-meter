using namespace System
using namespace System.Management.Automation

Set-StrictMode -Version Latest
$ErrorActionPreference = "stop"

if ([CommandOrigin]::Runspace -eq $MyInvocation.CommandOrigin) {
    throw "This script shouldn't be run with new runspace."
}

# SingletonオブジェクトとしてDriverを定義
function InitializeWebDriver($webDriver, $chromeDriver) {
    if (!(Test-Path Variable:Script:Driver) -or $Script:Driver -eq $null) {
        Add-Type -Path $webDriver
        $chromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
        $Script:Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver(
            (Split-Path $chromeDriver -Parent), $chromeOptions)
    }
    if (!(Test-Path Variable:Script:IsLogined)) {
        $Script:IsLogined = $false
    }
}

# ページがロードされるまで待機する
function SleepUntilLoaded($title) {
    $loadedTitle = ""
    do {
        Start-Sleep -Milliseconds 300
        $loadedTitle = $Script:Driver.Title
    } until($loadedTitle.Contains($title))
}

# WebDriverを閉じる。
function CloseWebDriver() {
    if ($Script:Driver -eq $null) {
        return
    }
    $Script:Driver.Quit()
    $Script:Driver = $null
    $Script:IsLogined = $false
}
