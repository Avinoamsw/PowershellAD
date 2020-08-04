import-module ActiveDirectory
Install-Module -Name DotNetVersionLister

$Logfile = "C:\Temp\ADComputers.log"
$mtx = New-Object System.Threading.Mutex($false, "MyMutex")

Function LogWrite
{
   Param ([string]$logstring)
   $mtx.WaitOne()
   Add-content -path $Logfile -value "$(Get-Date -format 'u'): $logstring" -Force
   $mtx.ReleaseMutex()
}

Function InvokeWithMutex
{
   Param ([string]$Command)
   $mtx.WaitOne()
   Invoke-Expression $Command
   $mtx.ReleaseMutex()
}


LogWrite "Starting Log..."

$Lastyear = (Get-Date).AddYears(-1)

#getting list of active win 10 computers that created in the last year
LogWrite "Getting computers list..."
Try{$Computers = Get-ADComputer -Properties Name -Filter {enabled -eq "true" -and OperatingSystem -like "*Windows 10*" -and whenCreated -ge $Lastyear }}
catch{
    LogWrite "Failed to get Comupters list."
    exit
}
if ($Computers.count -eq 0)
{
    LogWrite "No computers returned."
    exit
}
LogWrite "Computers list created successFully."
LogWrite "Gethring Data about the computers..."
$Records = @{}


$Computers | ForEach-Object -Parallel{
    InvokeWithMutex '$Records.Add($_.Name,@{})'

    #Get OS
    LogWrite "Start Getting OS info for $_.Name..."
    $OS = Get-WmiObject -Computer $_.Name -Class Win32_OperatingSystem
    InvokeWithMutex '$Records[$_.Name].add("OS",@{ FullName = $OS.Caption; Version = $OS.Version; BuildNumber = $OS.BuildNumber})'
    LogWrite "Getting OS info for $_.Name Completed."

    #Get Disks Info
    LogWrite "Start Getting Disks info for $_.Name..."
    $Disks = get-WmiObject -Computer $_.Name -Class win32_logicaldisk
    InvokeWithMutex '$Records[$_.Name].add("Disks",@{})'
    Foreach ($Disk in $Disks)
    {
        InvokeWithMutex '$Records[$_.Name]["Disks"].Add($Disk.DeviceID,@{ Size = $Disk.Size; Type = $Disk.DriveType})'
    }
    LogWrite "Getting Disks info for $_.Name Completed."

    #Get IP Address
    LogWrite "Start Getting IP info for $_.Name..."
    $IP = Test-Connection $_.Name -count 1 | select Ipv4Address
    InvokeWithMutex '$Records[$_.Name].Add("IP", $IP.IPV4Address)'
    LogWrite "Getting IP info for $_.Name Completed."

    #Get .Net Frameworks installed
    LogWrite "Start Getting .NET Frameworks info for $_.Name..." 
    $DotNet = gGet-STDotNetVersion -ComputerName $_.Name
    InvokeWithMutex '$Records[$_.Name].Add("DotNetVersions", $DotNet)'
    LogWrite "Getting .NET Frameworks info for $_.Name Completed."

    #Get Hotfixes Number & Latest installed
    LogWrite "Start Getting HotFixes info for $_.Name..."
    $HF = Get-HotFix -ComputerName $_.Name
    InvokeWithMutex '$Records[$_.Name].Add("HotFixes", @{NumberOfHotfixes = $HF.count; LastInstalledDate = ($HF | Sort-Object -Property InstalledOn)[-1].InstalledOn})'
    LogWrite "Getting HotFixes info for $_.Name Completed."


    #Get Time & Time zone
    LogWrite "Start Getting Time info for $_.Name..."
    $timeZone=Get-WmiObject -Class win32_timezone -ComputerName $_.Name
    $localTime = Get-WmiObject -Class win32_localtime -ComputerName $_.Name
    $Time = (Get-Date -Day $localTime.Day -Month $localTime.Month -Year $localTime.Year -Hour $localTime.Hour -Minute $localTime.Minute -Second $localTime.Second);
    InvokeWithMutex '$Records[$_.Name].Add("Time", @{TimeZone = $timeZone; Time = $Time})'
    LogWrite "Getting Time info for $_.Name Completed."


    #Get Admin users
    LogWrite "Start Getting Admins info for $_.Name..."
    $admins = Get-WmiObject win32_groupuser –computer $_.Name   
    $admins = $admins |where {$_.groupcomponent –like '*"Administrators"'}
    $admins = $admins |% {  
    $_.partcomponent –match “.+Domain\=(.+)\,Name\=(.+)$” > $nul  
    $matches[1].trim('"') + “\” + $matches[2].trim('"')}
    InvokeWithMutex '$Records[$_.Name].Add("Admins", $admins)'
    LogWrite "Getting Admins info for $_.Name Completed."


}

$Records | Export-Csv "C:\Temp\ADComputers.csv"
# SIG # Begin signature block
# MIIFdgYJKoZIhvcNAQcCoIIFZzCCBWMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUTQhrz7VK+LAI/LKBhPgz7LpI
# 2aagggMOMIIDCjCCAfKgAwIBAgIQHc4TXkNA3KZBr2CkD9GEPjANBgkqhkiG9w0B
# AQsFADAdMRswGQYDVQQDDBJBdmlzdyBDb2RlIFNpZ25pbmcwHhcNMjAwODA0MTMz
# NTU1WhcNMjEwODA0MTM1NTU1WjAdMRswGQYDVQQDDBJBdmlzdyBDb2RlIFNpZ25p
# bmcwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCkjUv8WyHFwy9L51rG
# eiEe8PiUNUOSqXM3l3xUaQfyhjfVLSmVAg2X3ZbvyeJbzF1r3+2BB6sVT1OEbdP/
# Ts0kFShEt71CMikTAl54P8CTWvEy8WdN6zeW6tRRMKigXZ+so85+Bph7nHjO21sG
# p5tT8AR6x/IgOxnzZLDMw4GG16Lz6K+ORQdaNwf5UrU9ktOSXeL4nc/VxRX+ZS6n
# Xovb7dmZmuyNZ3y+RLdnCzN+0yATyndrYHFaNQEpXKKh8W6hq8MrKwuaHzqG7tvE
# u+lmUCAbaRy/PmR5NPd0XwaVytVMiAx8WORWNvdeqyarZsa1iXthxEHci8BeRU62
# wDWZAgMBAAGjRjBEMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcD
# AzAdBgNVHQ4EFgQU09lcvhYjUHSoWsOENFTlCrM95aAwDQYJKoZIhvcNAQELBQAD
# ggEBAEhqpvjV067cXOjnshYV3Wzs33wO1HKQmS2gLTDsaFUPreDvlCkJ2pQ829v8
# QgWh2d3io9lIIz7nOQSbDFhPknRJOMcvaTUPqjvP0JvysGW/MDip0gtiwKBXoHma
# 9ZTpbBbPfUQD2TcJ6e62zh43bw7CQTHcUEaASrYF/P4ztQYMh8wsdegYZR3Xzaid
# 4ZZqMzqtndiPGvVc8Uk7Cxx5eO35kVrCL4OyFRqK2xMWtoHbdvFaPjAgFV5hE+/w
# u5kcdmIKglUtahKoWKmXKMf5NaE56ZpOIO7c1S68ZxJI0CwiSFKuKkq5D+l2YBnJ
# jzcOjg7CbkT0VLXvOkhkuxNyKLQxggHSMIIBzgIBATAxMB0xGzAZBgNVBAMMEkF2
# aXN3IENvZGUgU2lnbmluZwIQHc4TXkNA3KZBr2CkD9GEPjAJBgUrDgMCGgUAoHgw
# GAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQx
# FgQUAlnPPdwq3kJS6Ow9pe4SiM24Y48wDQYJKoZIhvcNAQEBBQAEggEAjAgh0b/M
# SslCcPEMYQOddywT6wbXcpi80bqHZFPf1Kxj//tbIdYLJv4SV4LDEz3n2uXy4nVs
# V6DzT013eSuCpjs59Cn2KHbPWOqNkZNGl1JO0qbv22fHueJ6pj/gypINXDH3E5X0
# NCcxVJOBZjyfFhPtj7XajjU+lT9KBGl/JIAXdknajpwSHvSraciHYfY+PRd8lEjn
# VKesUbkXMv1BysFj7y7z1F1ADA57M3xxbQ7k5mwUjBFSGDbKHxT77MaTGhWSJ5rL
# Sb8ApPNYdIz5nBkGnAaaVI0DGkkidKFBgg4MXlH/cmtWnx0hGKZFhcO4CgAzJl+s
# 1dNqatC58dQAzA==
# SIG # End signature block
