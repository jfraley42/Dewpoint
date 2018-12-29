#*** VARIABLES / CONFIG SETTINGS **********************************************
# change this to be directory where script file is stored and will be run from
$myScriptPath  =  "C:\Scripts\Powershell\OutOfOffice"
 
# start and end times as text strings. used in calendar, auto replies, email etc. use date + time
# WOULD LIKE TO USE CALENDAR FUNCTION BELOW TO GET DATES
$startTime     =  "12/03/2018 5:00 PM"    # format MM/dd/yyyy hh:mm AMPM
$endTime       =  "01/09/2019 8:00 AM"    # format MM/dd/yyyy hh:mm AMPM
 
# leave this be:
$startDt = [DateTime]::Parse($startTime)
$endDt = [DateTime]::Parse($endTime)
 
# calendar appointment subject and location
$apptCreate    =  $true
$apptSubject   =  "Out of Office"
$apptLocation  =  "Away"
 
# email address used for out of office auto reply and From for out of office email
$emailAddress  =  "my_email@mycompany.com"
$myName = "Geoff"
 
# internal and external messages for out of office automatic replies. Supports HTML
$autoReplySet  =  $true
$internalMsg   =  "I will be out of the office from &lt;font color='blue'&gt;" + $startDt.DayOfWeek.ToString() + " " + $startTime + "&lt;/font&gt; to &lt;font color='blue'&gt;" + $endDt.DayOfWeek.ToString() + " " + $endTime + "&lt;/font&gt;"
$externalMsg   =  $internalMsg
 
# this is who you want to email to notify ahead of time that you will be out of office
$emailSend     =  $true
# comma separate multiple addresses
$emailTo       =  "AppsTeam@mycompany.com, business_person@mycompany.com"
#$emailSubject =  [string]::Format("Out of office {0} - {1}", $startDt.ToShortDateString(), $endDt.ToShortDateString())
$emailSubject  =  "Out of office"
$emailBody     =  $internalMsg + ". Please see me before then if you need anything.&lt;br/&gt;&lt;br/&gt;Thanks,&lt;br/&gt;&lt;br/&gt;" + $myName + "&lt;br/&gt;&lt;br/&gt;Sent by OutOfOffice.ps1"
 
#*** CONSTANTS ****************************************************************
if (!(test-path variable:olFolderCalendar))
{ 
    New-Variable -Option constant -Name olFolderCalendar -Value 9
}    
 
if (!(test-path variable:olAppointmentItem))     
{
    New-Variable -Option constant -Name olAppointmentItem  -Value 1
}    
     
if (!(test-path variable:olOutOfOffice))         
{
    New-Variable -Option constant -Name olOutOfOffice  -Value 3
}    

#*** CREATE APPOINTMENT **********************************************************
if ($apptCreate)
{
    $outlook = new-object -com Outlook.Application
 
    #*** CREATE CALENDAR APPT *****************************************************
    $calendar = $outlook.Session.GetDefaultFolder($olFolderCalendar)
    $appt = $calendar.Items.Add($olAppointmentItem)
    $appt.Start = $startDt
    $appt.End = $endDt
    $appt.Subject = $apptSubject
    $appt.Location = $apptLocation
    $appt.BusyStatus = $olFree
    $appt.Save()
}

#*** SET OOO

#Need help getting this function to run
#Set-MailboxAutoReplyConfiguration -AutoReplyState Scheduled -StartTime $startTime -EndTime $endTime -InternalMessage "Internal auto-reply message" -ExternalMessage "External auto-reply message."


#*** CALENDAR FUNCTION **************************************************************

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object Windows.Forms.Form

$form.Text = 'Select a Date'
$form.Size = New-Object Drawing.Size @(243,230)
$form.StartPosition = 'CenterScreen'

$calendar = New-Object System.Windows.Forms.MonthCalendar
$calendar.ShowTodayCircle = $false
$calendar.MaxSelectionCount = 1
$form.Controls.Add($calendar)

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(38,165)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(113,165)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $date = $calendar.SelectionStart
    Write-Host "Date selected: $($date.ToShortDateString())"
}

