param(
    # Path to EWS-.dll, which is required
    $ewsDll = "<Path to the Microsoft.Exchange.WebServices.dll>",
    # Mailbox, which we go through
    $mailboxName = "<username of mailbox>",
    # Password for mailbox mentioned previously
    $mailboxPassword = "<password of mailbox>",
    # Limit mails beeing processed (throttling)
    $resultLimit = 100,
    # Path to the folder in mailbox, which contains the mails beeing processed. Stored as string array. Example: $folderPath = @("_demo", "_demodemo"),
    $folderPath = @("<Folder1>", "<Folder2>"),
    # Location in which the pdf-files getting downloaded
    $downloadPath = "<Path where you want do download the files>",
    # EWS URL
    $ewsUrl = "https://<Exchange-Url>/EWS/Exchange.asmx",
    # Content type to process
    $contentType = "application/pdf"
)

# Load the EWS Managed API
Add-Type -Path $ewsDll
$Exchange2013SP1 = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1 #Nicht wundern, seither keine Änderungen mehr

# Create EWS Service object for the target mailbox name
$exchangeService = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $Exchange2013SP1
$exchangeService.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($mailboxName, $mailboxPassword)
$exchangeService.Url = $ewsUrl

# Create a PropertySet with the Attachments metadata
$ItemPropetySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new(
    [Microsoft.Exchange.Webservices.Data.BasePropertySet]::IdOnly,
    [Microsoft.Exchange.WebServices.Data.ItemSchema]::Attachments,
    [Microsoft.Exchange.WebServices.Data.ItemSchema]::HasAttachments
)

# Bind to the Inbox folder of the target mailbox
$inboxFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
$inboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService,$inboxFolderName)

# Optional: reduce the query overhead by viewing the inbox $resultLimit items at a time
$itemView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $resultLimit
$folderView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderView -ArgumentList $resultLimit

# Going through the folder path
Write-Host "Navigating to: $($inboxFolder.DisplayName)/" -NoNewline
$curFolders = $exchangeService.FindFolders($inboxFolder.Id, $folderView)
$curIndex = 0
$highestPathIndex = $folderPath.Count - 1

while ( $curIndex -le $highestPathIndex){
    $foundFolder = $false
    # we go through each folder in the current folder
    foreach ($folder in $curFolders){
        # ... and check, if the folder is matching the name we look for
        if ($folder.DisplayName -eq $folderPath[$curIndex]){
            # if we come here, the current folder seems to match the required folder
            if ($curIndex -lt $highestPathIndex){
                # not the last folder, so we have to receive the folders in the folder and start all over
                $curFolders = $exchangeService.FindFolders($folder.Id, $folderView)
                Write-Host "$($folder.DisplayName)/" -NoNewline
            }
            else {
                # final folder - in this case $folder contains the correct folder, in which we have to look
                Write-Host "$($folder.DisplayName)" -NoNewline -ForegroundColor Green
            }
            $curIndex++
            $foundFolder = $true
            break
        }
    }
    # we checked each subfolders name. if we didnt found the expected name, we throw an error
    if ($foundFolder -eq $false) {
        throw "Folder ""$($folderPath[$curIndex])"" doesn't exist."
    }
}
Write-Host " (done)" -ForegroundColor Green
Write-Host ""

# Fetching all mails in the folder
Write-Host "Fetching Mails (max: $($resultLimit)): " -NoNewline
$allMails = $exchangeService.FindItems($folder.Id, $itemView)
Write-Host "$($allMails.Items.Count)" -ForegroundColor Yellow
Write-Host ""

if ($allMails.Items.Count -gt 0){
    Write-Host "Processing Mails:"
    Write-Host ""
    $mailIndex = 1
    $unprocessedMails = 0

    # ... and go through them
    foreach ($mail in $allMails){
        Write-Host "[$($mailIndex)/$($allMails.Items.Count)] - $($mail.Subject)"
        Write-Host "  From:        $($mail.From)"
        Write-Host "  To:          $($mail.DisplayTo)"
        Write-Host "  Attachments: " -NoNewline
        # checking if it has attachments
        if ($mail.HasAttachments)
        {
            $hadRequiredAttachment = $false
            $content = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService, $mail.Id, $ItemPropetySet)
            Write-Host "$($content.Attachments.Count)" -ForegroundColor Yellow
            foreach ($attachment in $content.Attachments) {
                Write-Host "   - $($attachment.Name) => " -NoNewline
                # only if it is an pdf, we will download it
                if ($attachment -is [Microsoft.Exchange.WebServices.Data.FileAttachment] -and $attachment.ContentType -eq $contentType) {
                    $FilePath = Join-Path $downloadPath $attachment.Name
                    $attachment.Load($filePath)
                    $hadRequiredAttachment = $true
                    Write-Host "Downloaded" -ForegroundColor Green
                }
                else {
                    
                    Write-Host "Not processed" -ForegroundColor Yellow
                }
            }
            if ($hadRequiredAttachment){
                Write-Host "  => Mail was processed and will be deleted: " -NoNewline
                $mail.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
                Write-Host "success" -ForegroundColor Green
            }
            else {
                $unprocessedMails++
                Write-Host "No attachment could be processed!" -ForegroundColor Red
            }
        }
        else {
            Write-Host "(Has no attachments)" -ForegroundColor Red
            $unprocessedMails++
        }
        $mailIndex++
        Write-Host ""
    }

    Write-Host "Finished processing mails"
    Write-Host "Processed: $($allMails.Items.Count) | Succeed: " -NoNewline
    Write-Host "$($allMails.Items.Count - $unprocessedMails)" -NoNewline -ForegroundColor Green
    Write-Host " | Failed: " -NoNewline
    Write-Host "$($unprocessedMails)" -NoNewline -ForegroundColor Red
}
