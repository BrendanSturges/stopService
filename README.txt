<#
Author: github.com/brendansturges
Description: stopService.ps1

1) Copy script to local workstation
2) Create a .txt file containing every server you want to run this script against.  Servers must be on a seperate line from each other.  ie, the contents of your serverlist will look like so

SERVER1
SERVER2
SERVER3
etc...

3) Run the .ps1 file from a powershell prompt
4) when prompted for service name, enter it
5) next prompt, choose your server list
6) next prompt, give it a name for the CSV file that is being generated.  This does not need an extension as that is dynamically generated.
6) Once the script completes you can check the CSV file you named for the status

#>