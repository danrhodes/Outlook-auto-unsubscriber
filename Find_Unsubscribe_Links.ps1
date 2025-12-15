# Create Outlook COM Object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNameSpace("MAPI")

# Choose the folder you want to check in your mailbox
$rootFolder = $namespace.PickFolder()

# Create hashtables for tracking unsubscribe links and processed domains (to avoid duplicates)
$unsubscribeLinks = @{}
$processedDomains = @{}

# Get the directory of the current script
$scriptPath = Split-Path -Parent $PSCommandPath

# Initialize HTML output with a modern CSS styling
$script:htmlOutput = @"
<html>
<head>
    <style>
        body {
            font-family: Arial, sans-serif;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 20px;
        }
        th, td {
            border: 1px solid #dddddd;
            text-align: left;
            padding: 8px;
        }
        th {
            background-color: #4CAF50;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
    </style>
</head>
<body>
    <h1>Unsubscribe Links</h1>
    <table>
        <tr>
            <th>Unsubscribe Link</th>
            <th>Sent From Domain</th>
        </tr>
"@

# Recursive function to process emails in a folder and its subfolders
function ProcessFolder($folder) {
    # Get all email items
    $emails = $folder.Items

    # Calculate the total number of emails for the progress bar
    $totalEmails = $emails.Count

    # Initialize progress bar
    $i = 0

    # Loop over each email
    foreach ($email in $emails) {
        # Increment progress bar
        $i++
        Write-Progress -Activity "Processing Emails in $($folder.Name)" -Status "$i out of $totalEmails Emails Processed" -PercentComplete (($i / $totalEmails) * 100)
        ProcessEmail $email
    }

    # Recurse into subfolders
    foreach ($subfolder in $folder.Folders) {
        ProcessFolder $subfolder
    }
}

# Function to process each email
function ProcessEmail($email) {
    try {
        # Use regex to extract hyperlinks containing "Unsubscribe"
        $matches = [regex]::Matches($email.HTMLBody, '<a href="(.*?)".*?>.*?(Unsubscribe|click here to unsubscribe).*?</a>', 'IgnoreCase')

        # Check if any matches were found
        if ($matches.Count -gt 0) {
            Write-Host "Found $($matches.Count) Unsubscribe link(s) in email: $($email.Subject)" -ForegroundColor Green

            # Extract the domain from the sender's email address
            $senderDomain = $email.SenderEmailAddress -replace '.*@'

            foreach ($match in $matches) {
                # Add link to hashtable and directly write to HTML output if it's not already there and the domain hasn't been processed yet
                if (!$unsubscribeLinks.ContainsKey($match.Groups[1].Value) -and !$processedDomains.ContainsKey($senderDomain)) {
                    $unsubscribeLinks[$match.Groups[1].Value] = $true
                    $processedDomains[$senderDomain] = $true
                    $script:htmlOutput += "<tr><td><a href=`"$($match.Groups[1].Value)`">Click Here to Unsubscribe</a></td><td>$senderDomain</td></tr>"
                }
            }
        } else {
            Write-Host "No Unsubscribe links found in email: $($email.Subject)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error processing email: $($email.Subject) - $_" -ForegroundColor Red
    }
}

# Process the root folder and all its subfolders
ProcessFolder $rootFolder

# Close HTML tags
$script:htmlOutput += "</table></body></html>"

# Write HTML output to file
$script:htmlOutput | Out-File -FilePath "$scriptPath\output.html"

# Report completion
Write-Host "Completed processing all emails. Unsubscribe links have been exported to $scriptPath\output.html" -ForegroundColor Cyan
