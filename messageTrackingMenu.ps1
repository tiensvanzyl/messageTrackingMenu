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

    Write-Host "================ Message Tracking - Choose your preferred option below ================" #-ForegroundColor Green
    
    Write-Host "1: Track mail with no subject full. (Output to console)."
    Write-Host "2: Track mail with no subject full. (Output to text file)."
    Write-Host "3: Track mail with subject full. (Output to console)."
    Write-Host "4: Track mail with subject full. (Output to text file)."
    Write-Host "5: Track mail with no subject delivered only. (Output to console)."
    Write-Host "6: Track mail with no subject delivered only. (Output to text file)."
    Write-Host "Q: Press 'Q' to quit."
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
                
                #Message tracking script for full detail here:
            } '2' {
                cls
                'You chose option #2'
                Get-ChildItem
            } '3' {
                cls
                'You chose option #3'
            } 'q' {
                return
            }
        }
        pause
    }
        until ($input -eq 'q')
