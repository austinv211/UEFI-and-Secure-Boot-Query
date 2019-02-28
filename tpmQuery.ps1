<#
Name: tpmQuery.ps1
Author: Austin Vargason
Description: queries devices in batches to get uefi properties 
#>

Function Get-Systems {
    param (
        [Parameter(Mandatory=$true)]
        [String]$FilePath
    )

    #try to import the csv
    Try { 
        $data = Import-Excel -Path $FilePath
    }
    Catch {
        Write-Error "Could not import sheet at: $FilePath"
        return
    }

    $data
}


Function Get-SystemData {
    param (
        [Parameter(Mandatory=$true)]
        [String]$ComputerName
    )

    Try {

        if (Test-WSMan -ComputerName $ComputerName -ErrorAction SilentlyContinue) {
            $ScriptBlock = {
                $test = Confirm-SecureBootUEFI -ErrorVariable eVar

                $isUEFI = $true
                $isSecureBoot = $false

                if ($eVar) {
                    $isUEFI = $false
                }
                else {
                    $isSecureBoot = $test
                }

                $objectProperty = [ordered]@{
                    isUEFI = $isUEFI
                    isSecureBoot = $isSecureBoot
                    SystemName = [System.Net.Dns]::GetHostName()
                }

                $obj = New-Object -TypeName psobject -Property $objectProperty

                Write-Output $obj
            }

            Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock -ErrorAction Stop
        }
        else {
            Write-Debug "WSMAN Error on Asset: $ComputerName"
        }

    }
    Catch {
        Write-Debug "Could Not Ivoke command on Asset: $ComputerName"
    }
}

Function Start-SystemQuery {
    param (
        [Parameter(Mandatory=$true)]
        [psobject[]]$SystemData
    )

    $i = 0
    $count = $SystemData.Count

    #loop through the system data
    foreach ($row in $SystemData) {
        $systemName = $row.NetBios_Name0

        if (($row.UEFI0 -eq "NULL") -or ($row.SecureBoot0 -eq "NULL")) {

            #wait for there to be enough space to create a new job
            Wait-Jobs

            #start a job to get the data from the computer
            Start-Job -Name $row.Netbios_Name0 -ScriptBlock ${function:Get-SystemData} -ArgumentList $row.Netbios_Name0 | Out-Null
        }

        $i++

        Write-Progress -Activity "Starting System Jobs" -Status "Started Job on System: $systemName" -PercentComplete (($i / $count) * 100)
    }
}

Function Get-QueryJobResults {

    #wait for all jobs to complete
    $data = Get-Job | Wait-Job | Receive-Job

    Write-Output $data | Select-Object SystemName, isUEFI, isSecureBoot
}

Function Wait-Jobs {
    while ((Get-Job | Where-Object {$_.State -eq "Running"} | Measure-Object | Select-Object -ExpandProperty Count) -ge 16) {
        # wait
    }
}

Function Export-UEFIData {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [psobject[]]$SystemData,
        [Parameter(Mandatory=$true)]
        [psobject[]]$FileData
    )

    Begin {
        $exportData = @()
    }
    Process {
        foreach ($v in $SystemData) {
            $res = $FileData | Where-Object {$_.Netbios_Name0 -eq $v.SystemName}
            
            if ($null -ne $res) {
                if ($v.isUEFI) {
                    $res.UEFI0 = 1
                }
                else {
                    $res.UEFI0 = 0
                }

                if ($v.isSecureBoot) {
                    $res.SecureBoot0 = 1
                }
                else {
                    $res.SecureBoot0 = 0
                }

                $exportData += $v
            }
        }
    }
    End {
        $exportData | Export-Csv -Path .\testOutput.csv
    }
}


Get-Job | Wait-Job | Remove-Job
$test = Get-Systems -FilePath .\SystemResults.xlsx
Start-SystemQuery -SystemData $test
Get-QueryJobResults | Export-UEFIData -FileData $test
$test | Export-Excel -Path .\SystemResults.xlsx -WorksheetName "Results" -TableName "SystemDataTable" -AutoSize
