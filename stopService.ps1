Function Get-FileName($initialDirectory){   
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
	Out-Null

	$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
	$OpenFileDialog.initialDirectory = $initialDirectory
	$OpenFileDialog.filter = "All files (*.*)| *.*"
	$OpenFileDialog.ShowDialog() | Out-Null
	$OpenFileDialog.filename
}

function Save-File([string] $initialDirectory ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() |  Out-Null
	
	$nameWithExtension = "$($OpenFileDialog.filename).csv"
	return $nameWithExtension
}

#Open a file dialog window to get the source file
$serverList = Get-Content -Path (Get-FileName)

#open a file dialog window to save the output
$fileName = Save-File $fileName

#define "i" for progress bar
$i = 0
$data = @()

$ErrorActionPreference= 'silentlycontinue'

#define the service being manipulated
$service = Read-Host "What is the service name?"

foreach($server in $serverList){
$serviceStatus = (get-service $service -computername $server).status
	Try{
		if($serviceStatus -eq "Running"){
			(Get-service -computername $server -name $service).stop()
			set-service -computername $server -name $service -startuptype Disabled
				$props = [ordered]@{
				'Server' = $server
				'Service' = $service
				'Status' = ''
				'Details' = (get-service $service -computername $server).status
				'Startup' = (Get-WmiObject -computername $server -Query "Select StartMode From Win32_Service Where Name='$service'").startmode
				}
			$obj = New-Object -TypeName PSObject -Property $props
		}
		else{
			$pingIt = Test-Connection -ComputerName $server -quiet -count 1
			if($pingIt -eq "True"){
				$ErrorMessage = $_.Exception.Message
				$props = [ordered]@{
				'Server' = $server
				'Service' = $service
				'Status' = $ErrorMessage
				'Details' = $serviceStatus
				'Startup' = (Get-WmiObject -Query "Select StartMode From Win32_Service Where Name='$service'").startmode
				}
			}
			else{
				$props = [ordered]@{
				'Server' = $server
				'Service' = $service
				'Status' = 'Server not responding'
				'Details' = ''
				'Startup' = ''
				}
			}
		}
	
	}
	Catch{
		$pingIt = Test-Connection -ComputerName $server -quiet -count 1
		if($pingIt -eq "True"){
			$ErrorMessage = $_.Exception.Message
			$props = [ordered]@{
			'Server' = $server
			'Service' = $service
			'Status' = $ErrorMessage
			'Details' = $serviceStatus
			'Startup' = (Get-WmiObject -Query "Select StartMode From Win32_Service Where Name='$service'").startmode
			}
		else{
			$props = [ordered]@{
			'Server' = $server
			'Service' = $service
			'Status' = 'Server not responding'
			'Details' = $serviceStatus
			'Startup' = (Get-WmiObject -Query "Select StartMode From Win32_Service Where Name='$service'").startmode
			}
		}
		}
	}
	
	$obj = New-Object -TypeName PSObject -Property $props
	$data += $obj
	$i++
	Write-Progress -activity "Stopping agent $service on server $i of $($serverList.count)" -percentComplete ($i / $serverList.Count*100)
}
$data | Where-Object {$_} | Export-Csv $filename -noTypeInformation