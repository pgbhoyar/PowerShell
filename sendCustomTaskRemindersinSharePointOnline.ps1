[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")


Function Get-SPOContext([string]$Url,[string]$UserName,[string]$Password)
{
    $SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)
    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName, $SecurePassword)
    return $context
}

Function Get-ListItems([Microsoft.SharePoint.Client.ClientContext]$Context, [String]$ListTitle) {
    $todaysDate = (Get-Date);
    
    $list = $Context.Web.Lists.GetByTitle($ListTitle)
    $queryRowLimit = 200;
    
    $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery;
    $camlQuery.ViewXml = "<View><RowLimit>$queryRowLimit</RowLimit></View>";
    
    $items = @();
    #Get the list items in the batches of queryRowLimit till we reach the end
    do 
    {
        $listItems = $list.GetItems( $camlQuery );
        $Context.Load($list);
        $Context.Load($listItems); 
        #execute query 
        $Context.ExecuteQuery();
        $camlQuery.ListItemCollectionPosition = $listItems.ListItemCollectionPosition;
        foreach($item in $listItems)
        {
            Try
            {
                # If you want to send the email reminder for the tasks which are overdue for more than 7 days use the following condition
                #if (($item["Status"] -eq 'Not Started') -AND (($todaysDate - $item["DueDate"]).TotalDays -gt 7))
                # If you want to send email reminders for all the incomplete tasks
                if ($item["Status"] -eq 'Not Started')
                {
                    Write-Host 'Item ID is ' $item["ID"]  ' and due date is '  $item["DueDate"]  ;
                    Add-Content $logFileLocation ('Item ID is ' + $item["ID"]  + ' and due date is '  + $item["DueDate"] );
                    # Add the item to the collection if the condition is met
                    $items += $item
                }
            }
            Catch [System.Exception]
            {
                # In case of exception, log the message
                Write-Host $_.Exception.Message
                Add-Content $logFileLocation ('ERROR : ' + $_.Exception.Message) ;
            }
        }
    }
    While($camlQuery.ListItemCollectionPosition -ne $null)
return $items 
}

Function sendMail($subject, $emailBody, $emailTo, $credential)
{
     Write-Host "Sending Email"
     Add-Content $logFileLocation ('Sending reminder email to ' + $emailTo + ' for task ' + $item["ID"]  + ' and due date is '  + $item["DueDate"] ) ;
     Send-MailMessage -To $emailTo -from $UserName -Subject $subject -Body $emailBody -BodyAsHtml -smtpserver smtp.office365.com -usessl -Credential $credential -Port 587
}


$logFileLocation = "C:\PGB\customTaskReminderinSharePointOnlineLog.txt";
Add-Content $logFileLocation ('Execution Started ' + (Get-Date));
$UserName = "userName@domainName.onmicrosoft.com";
#$Password = Read-Host -Prompt 'Enter the password';
$Password = "yourSecurePassword";
$Url = "https://domainName.sharepoint.com";
$listUrl = '<a href="listUrl">List Title</a>';;
$listTitle = "List Title";
$secureStringPwd = $Password | ConvertTo-SecureString -AsPlainText -Force 
$context = Get-SPOContext -Url $Url -UserName $UserName -Password $Password
$credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $secureStringPwd




$items = Get-ListItems -Context $context -ListTitle $listTitle;

Write-Host $items.Count;
 if ( $items.Count -gt 0){
     foreach($item in $items)
     {
         $id = $item["ID"]
         $title = $item["Title"]
         $status = $item["Status"]
         $owner = $item["AssignedTo"].Email
         $dueDate = $item["DueDate"];
         #Write-Host $item["Body"]
         #$body = $item["Body"]
         
         #Write-Host 'We need to send the reminder for this task ' $id, $title, $status, $owner, $dueDate
         $taskUrl = '<a href="' + $Url + '/Lists/' + $listTitle + '/DispForm.aspx?ID=' + $id + '&Source=' + $Url + '">this link</a>';
         #Write-Host $taskUrl
         $emailBody = "The " + $title + " has been submitted for your approval.</br></br>
To take action on this task, click on " + $taskUrl + " and perform any of the following steps:</br></br>
1. Please review and approve, click on the <b>'Approved'</b> button </br>
2. If you want to reject the request, click on the <b>'Rejected'</b> button.</br></br>

To track the status of this request go to the site at " + $listUrl + ".</br></br>

Thank You,</br>

This is automatic email. Please do not respond."
             
         $emailSubject = 'Friendly reminder to approve a ' +  $title;
             
         sendMail $emailSubject $emailBody $owner $credential;
            
        }
    }
Add-Content $logFileLocation ('Execution ended ' + (Get-Date));
