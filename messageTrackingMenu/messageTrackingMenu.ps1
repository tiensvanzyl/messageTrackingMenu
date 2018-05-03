###########################################################################################
#   A message tracking script with a menu to assist users with easy message tracking      #
#   Author: Tiens van Zyl                                                                 #
#   Email: tiens.vanzyl@gmail.com                                                         #
#   Date:   31 April 2018                                                                 #
#   Updated by:                                                                           #
#   Updates:                                                                              #
#   Version: 0.1                                                                          #
#   External link:                                                                        #
#   Description:                                                                          #        
#       This script was written as an easy way to track messages in Microsoft Exchange    #
#        for users and junior administrators.                                             #
#                                                                                         #
###########################################################################################

function Show-Menu() {

    Write-Host "================ Message Tracking - Choose your preferred option below ================`n" -ForegroundColor Green
    
    Write-Host "1: Track mail with no subject full. (Output to console).`n" 
    Write-Host "2: Track mail with no subject full. (Output to text file).`n" 
    Write-Host "3: Track mail with subject full. (Output to console).`n"
    Write-Host "4: Track mail with subject full. (Output to text file).`n" 
    Write-Host "5: Track mail with no subject delivered to recipient only. (Output to console).`n"
    Write-Host "6: Track mail with no subject delivered to recipient only. (Output to text file).`n" 
    Write-Host "7: Track mail with no subject sender only. (Output to console).`n"
    Write-Host "8: Track mail with no subject sender only. (Output to text file).`n"
    Write-Host "9: Track mail with no subject recipient only. (Output to console).`n" 
    Write-Host "10: Track mail with no subject recipient only. (Output to text file).`n" 
    Write-Host "Q: Press 'Q' to quit.`n"
    }
    
    do {
        Show-Menu
        $input = Read-Host "Please make a selection"
        switch ($input) {
            '1' {
                cls
                $Server = Read-Host -Prompt "Enter the server  name or server wildcard for multiple servers i.e server0*"
                $Sender = Read-Host -Prompt "Enter the sender email address."
                $Recipient = Read-Host -Prompt "Enter the recipient email address."
                $StartDate = Read-Host -Prompt "Enter the start date (mm/dd/yy)."
                $EndDate = Read-Host -Prompt "Enter the end date (mm/dd/yy)."
                
                Get-TransportServer $Server | Get-MessageTrackingLog -Sender $Sender -recipient $Recipient -start $StartDate -End $EndDate -resultsize unlimited | select-object Timestamp,EventID,ServerHostname,Totalbytes,MessageSubject,Sender,ReturnPath,@{Name="Recipients";Expression={$_.recipients}}
                
            } '2' {
                cls
                $Server = Read-Host -Prompt "Enter the server  name or server wildcard for multiple servers i.e server0*"
                $Sender = Read-Host -Prompt "Enter the sender email address."
                $Recipient = Read-Host -Prompt "Enter the recipient email address."
                $StartDate = Read-Host -Prompt "Enter the start date (mm/dd/yy)."
                $EndDate = Read-Host -Prompt "Enter the end date (mm/dd/yy)."
                $Output = Read-Host -Prompt "Enter the path for the text file output. (c:\temp\tracking.txt)"
                
                Get-TransportServer $Server | Get-MessageTrackingLog -Sender $Sender -recipient $Recipient -start $StartDate -End $EndDate -resultsize unlimited | select-object Timestamp,EventID,ServerHostname,Totalbytes,MessageSubject,Sender,ReturnPath,@{Name="Recipients";Expression={$_.recipients}} > $Output

                Write-Host "The results have been saved to $Output" -ForegroundColor Green

            } '3' {
                cls
                $Server = Read-Host -Prompt "Enter the server  name or server wildcard for multiple servers i.e server0*"
                $Sender = Read-Host -Prompt "Enter the sender email address."
                $Recipient = Read-Host -Prompt "Enter the recipient email address."
                $StartDate = Read-Host -Prompt "Enter the start date (mm/dd/yy)."
                $EndDate = Read-Host -Prompt "Enter the end date (mm/dd/yy)."
                $Subject = Read-Host -Prompt "Enter the subject."

                Get-TransportServer $Server | Get-MessageTrackingLog -Sender $Sender -recipient $Recipient -MessageSubject $Subject -start $StartDate -End $EndDate -resultsize unlimited | select-object Timestamp,EventID,ServerHostname,Totalbytes,MessageSubject,Sender,ReturnPath,@{Name="Recipients";Expression={$_.recipients}}

            }'4' {
                cls
                $Server = Read-Host -Prompt "Enter the server  name or server wildcard for multiple servers i.e server0*"
                $Sender = Read-Host -Prompt "Enter the sender email address."
                $Recipient = Read-Host -Prompt "Enter the recipient email address."
                $StartDate = Read-Host -Prompt "Enter the start date (mm/dd/yy)."
                $EndDate = Read-Host -Prompt "Enter the end date (mm/dd/yy)."
                $Subject = Read-Host -Prompt "Enter the subject."
                $Output = Read-Host -Prompt "Enter the path for the text file output. (c:\temp\tracking.txt)"

                Get-TransportServer $Server | Get-MessageTrackingLog -Sender $Sender -recipient $Recipient -MessageSubject $Subject -start $StartDate -End $EndDate -resultsize unlimited | select-object Timestamp,EventID,ServerHostname,Totalbytes,MessageSubject,Sender,ReturnPath,@{Name="Recipients";Expression={$_.recipients}} > $Output

                Write-Host "The results have been saved to $Output" -ForegroundColor Green

            }'5' {
                cls
                $Server = Read-Host -Prompt "Enter the server  name or server wildcard for multiple servers i.e server0*"
                $Recipient = Read-Host -Prompt "Enter the recipient email address."
                $StartDate = Read-Host -Prompt "Enter the start date (mm/dd/yy)."
                $EndDate = Read-Host -Prompt "Enter the end date (mm/dd/yy)."
                
                Get-TransportServer $Server | Get-MessageTrackingLog -recipient $Recipient -start $StartDate -End $EndDate -resultsize unlimited  | where {$_.eventid –eq “deliver”} | select-object Timestamp,EventID,ServerHostname,Totalbytes,MessageSubject,Sender,ReturnPath,@{Name="Recipients";Expression={$_.recipients}}

            }'6' {
                cls
                $Server = Read-Host -Prompt "Enter the server  name or server wildcard for multiple servers i.e server0*"
                $Recipient = Read-Host -Prompt "Enter the recipient email address."
                $StartDate = Read-Host -Prompt "Enter the start date (mm/dd/yy)."
                $EndDate = Read-Host -Prompt "Enter the end date (mm/dd/yy)."
                $Output = Read-Host -Prompt "Enter the path for the text file output. (c:\temp\tracking.txt)"

                Get-TransportServer $Server | Get-MessageTrackingLog -recipient $Recipient -start $StartDate -End $EndDate -resultsize unlimited  | where {$_.eventid –eq “deliver”} | select-object Timestamp,EventID,ServerHostname,Totalbytes,MessageSubject,Sender,ReturnPath,@{Name="Recipients";Expression={$_.recipients}} > $Output

                Write-Host "The results have been saved to $Output" -ForegroundColor Green

            }'7' {
                cls
                $Server = Read-Host -Prompt "Enter the server  name or server wildcard for multiple servers i.e server0*"
                $Sender = Read-Host -Prompt "Enter the sender email address."
                $StartDate = Read-Host -Prompt "Enter the start date (mm/dd/yy)."
                $EndDate = Read-Host -Prompt "Enter the end date (mm/dd/yy)."
                $Output = Read-Host -Prompt "Enter the path for the text file output. (c:\temp\tracking.txt)"

                Get-TransportServer $Server | Get-MessageTrackingLog -Sender $Sender -start $StartDate -End $EndDate -resultsize unlimited  | select-object Timestamp,EventID,ServerHostname,Totalbytes,MessageSubject,Sender,ReturnPath,@{Name="Recipients";Expression={$_.recipients}}

            }'8' {
                cls
                $Server = Read-Host -Prompt "Enter the server  name or server wildcard for multiple servers i.e server0*"
                $Sender = Read-Host -Prompt "Enter the sender email address."
                $StartDate = Read-Host -Prompt "Enter the start date (mm/dd/yy)."
                $EndDate = Read-Host -Prompt "Enter the end date (mm/dd/yy)."
                $Output = Read-Host -Prompt "Enter the path for the text file output. (c:\temp\tracking.txt)"

                Get-TransportServer $Server | Get-MessageTrackingLog -Sender $Sender -start $StartDate -End $EndDate -resultsize unlimited  | select-object Timestamp,EventID,ServerHostname,Totalbytes,MessageSubject,Sender,ReturnPath,@{Name="Recipients";Expression={$_.recipients}} > $Output

                Write-Host "The results have been saved to $Output" -ForegroundColor Green

            }'9' {
                cls
                $Server = Read-Host -Prompt "Enter the server  name or server wildcard for multiple servers i.e server0*"
                $Recipient = Read-Host -Prompt "Enter the recipient email address."
                $StartDate = Read-Host -Prompt "Enter the start date (mm/dd/yy)."
                $EndDate = Read-Host -Prompt "Enter the end date (mm/dd/yy)."
                
                Get-TransportServer $Server | Get-MessageTrackingLog -Recipient $Recipient -start $StartDate -End $EndDate -resultsize unlimited  | select-object Timestamp,EventID,ServerHostname,Totalbytes,MessageSubject,Sender,ReturnPath,@{Name="Recipients";Expression={$_.recipients}}

            }'10' {
                cls
                $Server = Read-Host -Prompt "Enter the server  name or server wildcard for multiple servers i.e server0*"
                $Recipient = Read-Host -Prompt "Enter the recipient email address."
                $StartDate = Read-Host -Prompt "Enter the start date (mm/dd/yy)."
                $EndDate = Read-Host -Prompt "Enter the end date (mm/dd/yy)."
                $Output = Read-Host -Prompt "Enter the path for the text file output. (c:\temp\tracking.txt)"

                Get-TransportServer $Server | Get-MessageTrackingLog -Recipient $Recipient -start $StartDate -End $EndDate -resultsize unlimited  | select-object Timestamp,EventID,ServerHostname,Totalbytes,MessageSubject,Sender,ReturnPath,@{Name="Recipients";Expression={$_.recipients}} > $Output

                Write-Host "The results have been saved to $Output" -ForegroundColor Green

            } 'q' {
                return
            }
        }
        pause
    }
        until ($input -eq 'q')
