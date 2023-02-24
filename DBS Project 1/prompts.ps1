# Christopher Jeys
Function GetUserSelection {
    <# 
    .Synopsis
        Function to display a menu and collect the users output
    .Description
        This function uses StringBuilder to build and display a menu to the user.
        Then the results are returned based on the users input
    #>
    $Menu = New-Object -TypeName System.Text.StringBuilder 
        $Menu.AppendLine('CJEYS - C916 Task 1')
        $Menu.AppendLine('...................')
        $Menu.AppendLine('1) List log files')
        $Menu.AppendLine('2) List all files')
        $Menu.AppendLine('3) Show CPU & Memory consumption')
        $Menu.AppendLine('4) Show all running processes')
        $Menu.AppendLine('5) Exit')
    Write-Output -ForegroundColor Cyan $Menu.ToString()
    return $(Write-Output '>> Choose your selection (1-5)';Read-Host)
}
# Variable to hold the users selection
$UserSelection = 0
Try {
    while ($UserSelection -ne 5) {
        #Display the menu to the user
        Write-Output $Menu.ToString
        {
            #$UserSelection = GetUserSelection
            $UserSelection = Read-Host 'Please select a number'
            switch ($UserSelection) {
                1{
                    # User selected option 1
                    # Listing all files within the Requirements1 folder, with the .log file extension.  Results are redirected to a new file called “DailyLog.txt”
                    Write-Output Listing todays log files
                    'TIMESTAMP: ' + (Get-Date) | Out-File -FilePath $PSScriptRoot\DailyLog.txt -Append
                    Get-ChildItem -Path $PSScriptRoot -Filter *.log | Out-File -FilePath $PSScriptRoot\DailyLog.txt -Append
                }
                2{
                    # User selected option 2
                    # Listing all files inside the “Requirements1” folder in tabular format, sorted in ascending alphabetical order. Output is redirected into a new file called “C916contents.txt”.
                    Write-Output Listing all files
                    Get-ChildItem $PSScriptRoot * | Sort-Object Name | Format-Table -AutoSize -Wrap | Out-File -FilePath $PSScriptRoot\C916contents.txt 
                }
                3{
                    # User selected option 3
                    # The current CPU and memory usage will be displayed
                    Write-Output Showing CPU and Memory consumption
                    Get-Counter -Counter '\Memory\Committed Bytes' -SampleInterval 2 -MaxSamples 3
                    Get-Counter -Counter '\Processor(_Total)\% Processor Time' -SampleInterval 2 -MaxSamples 3
                }
                4{
                    # Listing all the running processes with the output sorted by virtual size used least to greatest, and displayed in grid format.
                    Write-Output Showing all running processes in a pop-up window
                    Get-Process | Select-Object ID, Name, VM | Sort-Object -Property VM | Out-GridView
                }
                5{
                    # User selected option 5
                    # Exiting the script
                    Write-Output Exiting now...
                    Exit-PSSession
                }
            }
        }
    }
}
#Error handling to catch System.OutOfMemoryException errors
Catch [System.OutOfMemoryException] {  
    Write-Output 'An Out of Memory Exception Occurred'
}
Finally {   
    # Closes open resources
}