[Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll")

function Create-RandomString()
{
  $aChars = @()
  $aChars = "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "A", "C", "B", "D", "E", "F", "G", "H", "J", "K", "M", "N", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "2", "3", "4", "5", "6", "7", "8", "9", "_", "+"
  $intUpperLimit = Get-Random -minimum 12 -maximum 13

  $x = 0
  $strString = ""
  while ($x -lt $intUpperLimit)
  {
     $a = Get-Random -minimum 0 -maximum $aChars.getupperbound(0)
     $strString += $aChars[$a]
     $x += 1
  }

  return $strString
}

$email = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)
$email.Credentials = New-Object Net.NetworkCredential('svc-o365-pass-reset@zocdoc.onmicrosoft.com', 'i2m@f\SR]*]qc6&j')
$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
$email.Url = $uri
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($email,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
if ($inbox.UnreadCount -gt 0)
{
 $sendMailCredentials = New-Object System.Management.Automation.PsCredential 'svc-o365-pass-reset@zocdoc.onmicrosoft.com',(ConvertTo-SecureString -String 'i2m@f\SR]*]qc6&j' -AsPlainText -Force)
 $PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
 $PropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;
  # Set search criteria - unread only
 $SearchForUnRead = New-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $false) 
 $items = $inbox.FindItems($SearchForUnRead,20)  #return only 20 unread mail items
 foreach ($item in $items.Items)
 {
   # load the property set to allow us to view the body
  $item.load($PropertySet)
  if($item.HasAttachments -eq "True")
   {
    $item.Attachments[0].Load()
    $samplepath = "C:\PassReset\done.txt"
    set-content -value $item.Attachments.Content -enc byte -path $samplepath
    $done = Get-Content $samplepath
   }
  if($item.Body.text -Like "*Reset my password*" -or $done -like "*Reset my password*")
   {
   $Phone = $item.From.address.substring(0, $item.From.address.IndexOf("@"))
   if ($Phone.substring(0, 1) -eq "1")
   {
    $Phone = $Phone.substring(1)
   }
   if ($Phone.substring(0, 1) -eq "+")
   {
    $Phone = $Phone.substring(2)
   }
   $Phone = "{0: +1 (###) ###-####}" -f [double]$Phone
   $user = get-aduser -filter "mobilePhone -eq '$Phone'" -Properties PasswordNeverExpires
   If ($user -ne $null)
   {
    $PW = Create-RandomString
    if ($PW.length -gt 11)
    {
     Set-ADAccountPassword -identity $user.samaccountname -NewPassword (ConvertTo-SecureString -AsPlainText $PW -Force)
     Unlock-ADAccount -identity $user.samaccountname
     if ($user.passwordneverexpires -eq 0)
        {
            Set-ADUser -identity $user.samaccountname -ChangePasswordAtLogon $true
        }
     send-mailmessage -to "reset-pass@zocdoc.com" -from "reset-pass@zocdoc.com" -subject "Password Reset" -body "Password reset for $($user.SamAccountName) - $($user.DistinguishedName)" -SmtpServer smtp.office365.com -UseSsl -Port 587 -Credential $sendMailCredentials
     send-mailmessage -to $item.From.Address -from "reset-pass@zocdoc.com" -subject "Password Reset Success" -body "Your temporary password is $PW and you will be asked to update this upon logging in." -SmtpServer smtp.office365.com -UseSsl -Port 587 -Credential $sendMailCredentials
    }
   else
    {
     send-mailmessage -to "reset-pass@zocdoc.com" -from "reset-pass@zocdoc.com" -subject "Invalid Phone number" -body "Phone number $Phone not found" -SmtpServer smtp.office365.com -UseSsl -Port 587 -Credential $sendMailCredentials
     send-mailmessage -to $item.From.Address -from "reset-pass@zocdoc.com" -subject "Password Reset Failure" -body "Your phone was not found. Please contact Action-IT@zocdoc.com to update your Outlook mobile number." -SmtpServer smtp.office365.com -UseSsl -Port 587 -Credential $sendMailCredentials
    }
   }
  }
  $item.Isread = $true
  $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
  Remove-Item $samplepath
 }
}