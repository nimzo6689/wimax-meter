using namespace System.Management.Automation
using namespace System.Text
using namespace System.Windows.Forms

Set-StrictMode -Version Latest
$ErrorActionPreference = "stop"

if ([CommandOrigin]::Runspace -eq $MyInvocation.CommandOrigin) {
    throw "This script shouldn't be run with new runspace."
}

. $(Join-Path $PSScriptRoot _WebDriverClient.ps1)

# Spped Wi-Fi NEXT へログインする
# 既にログイン済みの場合は何もしない。
# return -> [boolean] ログイン成功の場合、true。
function Login($userType, $encodedPass) {
    if ($Script:IsLogined -eq $false) {
        $Script:Driver.Url = "http://speedwifi-next.home/html/login.htm"
        SleepUntilLoaded -title "Speed"

        $Script:Driver.FindElementsByXPath('//*[@id="user_type"]').SendKeys($userType)
        $decodedPass = [Encoding]::UTF8.GetString([Convert]::FromBase64String($encodedPass))
        $Script:Driver.FindElementsByXPath('//*[@id="input_password"]').SendKeys($decodedPass)
        $Script:Driver.FindElementsByXPath('//*[@id="login"]').Click()
        SleepUntilLoaded -title "Speed"

        $Script:IsLogined = $Script:Driver.Url -notcontains "login"
        if (!$Script:IsLogined) {
            [void][MessageBox]::Show("ログインできませんでした。`nログインIDかパスワードが間違っています。")
        }
    }
}

# 各通信料を取得する
# １ヵ月、前日までの３日間、本日までの３日間
function GetUsedData() {
    $Script:Driver.Url = "http://speedwifi-next.home/html/statistics.htm"
    SleepUntilLoaded -title "Speed"

    return [PSCustomObject]@{
        day         = [datetime]::Today.AddDays(-1).ToString("yyyy/MM/dd")
        inMonth     = ($Script:Driver.FindElementsByXPath('//*[@id="label_usedData"]').Text -split " ")[0]
        toYesterday = ($Script:Driver.FindElementsByXPath('//*[@id="label_usedData_yesterday"]').Text -split " ")[0]
        toToday     = ($Script:Driver.FindElementsByXPath('//*[@id="label_usedData_today"]').Text -split " ")[0]
    }
}