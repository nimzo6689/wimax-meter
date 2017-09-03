[PSCustomObject]@{
    Selenium = [PSCustomObject]@{
        # Path of WebDriver.dll
        WebDriver    = "Input absolute path of WebDriver.dll"
        # Path of chromedriver.exe
        ChromeDriver = "Input absolute path of chromedriver.exe"
    }
    Auth     = [PSCustomObject]@{
        UserType    = "admin"
        # [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($cmd))
        EncodedPass = "Input your password."
    }
    # Path of wimax_yyyyMMdd.csv
    CsvDir   = "Input any directory on your computer."
}