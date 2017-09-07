using namespace System
using namespace System.Management.Automation
using namespace System.Runtime.Interopservices

Set-StrictMode -Version Latest
$ErrorActionPreference = "stop"

if ([CommandOrigin]::Runspace -eq $MyInvocation.CommandOrigin) {
    throw "This script shouldn't be run with new runspace."
}

# enum型の名前はシート名として使用します。
enum Month {
    Jan = 1; Feb = 2; Mar = 3; Apr = 4; May = 5; Jun = 6
    Jul = 7; Aug = 8; Sep = 9; Oct = 10; Nov = 11; Dec = 12
}

function MonthOf($number) {
    return $(switch ($number) {
            1 { [Month]::Jan }
            2 { [Month]::Feb }
            3 { [Month]::Mar }
            4 { [Month]::Apr }
            5 { [Month]::May }
            6 { [Month]::Jun }
            7 { [Month]::Jul }
            8 { [Month]::Aug }
            9 { [Month]::Sep }
            10 { [Month]::Oct }
            11 { [Month]::Nov }
            12 { [Month]::Dec }
        })
}

# SingletonオブジェクトとしてExcelを定義
function InitializeExcel($excelFullName) {
    if (!(Test-Path Variable:Script:Excel) -or $Script:Excel -eq $null) {
        $Script:Excel = New-Object -ComObject Excel.Application
        # $Script:Excel.Visible = $true
        # $Script:Excel.DisplayAlerts = $true
        $Script:Book = $Script:Excel.Workbooks.Open($excelFullName)
        $month = MonthOf -number ([datetime]::Today.Month)
        $Script:Sheet = $Script:Book.Worksheets.Item($month.ToString())
        $Script:Sheet.Activate()
    }
}

# セルへ書き込みする。
function SetUsedDataInRow($row, $daily, $inMonth, $toYesterday, $toToday) {
    $Script:Sheet.Cells.Item($row, $daily)
    $Script:Sheet.Cells.Item($row, $inMonth)
    $Script:Sheet.Cells.Item($row, $toYesterday)
    $Script:Sheet.Cells.Item($row, $toToday)
}

# Excelへの算出方式に従って出力する。
function WriteExcelWithUsedData($usedData) {
    $day = [datetime]::Today.Day
    switch ($day) {
        1 {
            $month = MonthOf -number ([datetime]::Today.Month - 1)
            $Script:Sheet = $Script:Book.Worksheets.Item($month.ToString())
            $Script:Sheet.Activate()
            $row = [datetime]::Today.AddDays(-1).Day + 2
            $daily = $usedData.toToday - $Script:Sheet.Cells.Item($row - 1, 2).Text
            SetUsedDataInRow -row $row -daily $daily -inMonth ($Script:Sheet.Cells.Item($row - 1, 3).Text + $daily) \
            -toYesterday $usedData.toYesterday -toToday $usedData.toToday
        }
        2 {
            $row = $day + 2
            SetUsedDataInRow -row $row -daily $usedData.inMonth -inMonth $usedData.inMonth \
            -toYesterday $usedData.toYesterday -toToday $usedData.toToday
        }
        default {
            $row = $day + 2
            $daily = $usedData.inMonth - $Script:Sheet.Cells.Item($row - 1, 3).Text
            SetUsedDataInRow -row $row -daily $daily -inMonth $usedData.inMonth \
            -toYesterday $usedData.toYesterday -toToday $usedData.toToday
        }
    }
}

# ExcelのComObjectを閉じる。
function CloseExcel() {
    if ($Script:Excel -eq $null) {
        return
    }
    $Script:Excel.ActiveWorkbook.Save()    
    $Script:Excel.Quit()
    $Script:Excel, $Script:Book, $Script:Sheet | 
        ForEach-Object {[Marshal]::ReleaseComObject($_)}
    $Script:Excel = $null
}

