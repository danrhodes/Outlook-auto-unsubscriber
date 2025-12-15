# Find Unsubscribe Links - Enhanced Edition
# Features: Email counts, priority detection, sorting, click tracking, one-click unsubscribe, and statistics

# Check if Microsoft.Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Mail)) {
    Write-Host "Microsoft Graph PowerShell SDK is not installed." -ForegroundColor Yellow
    Write-Host "Installing Microsoft.Graph.Mail module..." -ForegroundColor Cyan
    Install-Module Microsoft.Graph.Mail -Scope CurrentUser -Force
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}

# Import required modules
Import-Module Microsoft.Graph.Mail
Import-Module Microsoft.Graph.Authentication

# Prompt for user email address
Write-Host "Enter your email address (the account you want to scan for unsubscribe links):" -ForegroundColor Cyan
$userEmail = Read-Host "Email address"

# Connect to Microsoft Graph
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan
Write-Host "You will be prompted to sign in with your Microsoft 365 account." -ForegroundColor Yellow

try {
    # Connect with necessary permissions
    Connect-MgGraph -Scopes "Mail.Read", "Mail.ReadBasic" -NoWelcome
    Write-Host "Successfully connected to Microsoft Graph!" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    exit
}

# Prompt the user for the number of days to check
do {
    $daysToCheck = Read-Host "`nEnter the number of days to check back on emails"
} while ($daysToCheck -notmatch '^\d+$' -or $daysToCheck -le 0)

# Prompt for folder selection
Write-Host "`nAvailable folder options:" -ForegroundColor Cyan
Write-Host "1. Inbox (default)"
Write-Host "2. Sent Items"
Write-Host "3. Junk Email"
Write-Host "4. All Mail (entire mailbox)"
$folderChoice = Read-Host "Select folder option (1-4, default: 1)"

$folderName = "Inbox"
switch ($folderChoice) {
    "2" { $folderName = "SentItems" }
    "3" { $folderName = "JunkEmail" }
    "4" { $folderName = "AllMail" }
    default { $folderName = "Inbox" }
}

# Ask about one-click unsubscribe feature
Write-Host "`n=== ONE-CLICK UNSUBSCRIBE FEATURE ===" -ForegroundColor Cyan
Write-Host "Some emails support RFC 8058 List-Unsubscribe-Post which allows automatic unsubscription" -ForegroundColor Yellow
Write-Host "without opening a browser. This feature can automatically unsubscribe you from these emails." -ForegroundColor Yellow
Write-Host "WARNING: This will immediately unsubscribe you without confirmation!" -ForegroundColor Red
do {
    $enableOneClick = Read-Host "`nEnable automatic one-click unsubscribe? (y/n)"
} while ($enableOneClick -notmatch '^(y|yes|n|no)$')
$enableOneClick = $enableOneClick -match '^y'

# Create a date object for the specified number of days ago
$startDate = (Get-Date).AddDays(-[int]$daysToCheck)
$filterDate = $startDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

# Enhanced tracking with domain info
$domainInfo = @{}  # Will store: domain -> @{ Count, UnsubLink, Subject, ReceivedDate, ListUnsubPost }
$oneClickUnsubscribed = @()

# Get the directory of the current script
$scriptPath = Split-Path -Parent $PSCommandPath
if (-not $scriptPath) {
    $scriptPath = Get-Location
}

# Function to process each email
function ProcessEmail($email) {
    try {
        # Extract sender information first
        $senderEmail = $null
        if ($email.From.EmailAddress.Address) {
            $senderEmail = $email.From.EmailAddress.Address
        } elseif ($email.Sender.EmailAddress.Address) {
            $senderEmail = $email.Sender.EmailAddress.Address
        }

        if (-not $senderEmail) {
            return
        }

        # Extract the domain from the sender's email address
        $senderEmailParts = $senderEmail -split "@"
        if ($senderEmailParts.Count -lt 2) {
            return
        }

        $senderDomain = $senderEmailParts[-1].ToLower()

        # Initialize or update domain info
        if (-not $script:domainInfo.ContainsKey($senderDomain)) {
            $script:domainInfo[$senderDomain] = @{
                Count = 0
                UnsubLink = $null
                Subject = $null
                ReceivedDate = $null
                ListUnsubPost = $null
            }
        }

        # Increment email count for this domain
        $script:domainInfo[$senderDomain].Count++

        # Store the most recent subject and date if not already set
        if (-not $script:domainInfo[$senderDomain].Subject) {
            $script:domainInfo[$senderDomain].Subject = $email.Subject
            $script:domainInfo[$senderDomain].ReceivedDate = $email.ReceivedDateTime
        }

        # Skip further processing if we already have an unsubscribe link for this domain
        if ($script:domainInfo[$senderDomain].UnsubLink) {
            return
        }

        # Get email body
        $emailBody = ""
        if ($email.Body.Content) {
            $emailBody = $email.Body.Content
        }

        if (-not $emailBody) {
            return
        }

        # Collection to store found links
        $foundLinks = @()
        $listUnsubPost = $null

        # Method 1: Check for List-Unsubscribe header (RFC 2369) - Most reliable
        if ($email.InternetMessageHeaders) {
            $unsubHeader = $email.InternetMessageHeaders | Where-Object { $_.Name -eq "List-Unsubscribe" }
            if ($unsubHeader) {
                # Extract URLs from List-Unsubscribe header
                $headerMatches = [regex]::Matches($unsubHeader.Value, '<(https?://[^>]+)>')
                foreach ($headerMatch in $headerMatches) {
                    $foundLinks += $headerMatch.Groups[1].Value
                }
            }

            # Check for List-Unsubscribe-Post header (RFC 8058 - One-Click Unsubscribe)
            $unsubPostHeader = $email.InternetMessageHeaders | Where-Object { $_.Name -eq "List-Unsubscribe-Post" }
            if ($unsubPostHeader -and $unsubHeader) {
                # Store the List-Unsubscribe URL and Post data for one-click unsubscribe
                $listUnsubPost = @{
                    Url = $foundLinks[0]
                    PostData = $unsubPostHeader.Value
                }
                $script:domainInfo[$senderDomain].ListUnsubPost = $listUnsubPost
            }
        }

        # Method 2: Look for <a> tags with unsubscribe-related text
        $pattern1 = '<a[^>]+href\s*=\s*["' + "'" + ']([^"' + "'" + ']+)["' + "'" + '][^>]*>(?:[^<]*(?:unsubscribe|opt[\s\-]?out|remove[\s]?me|manage[\s]?preferences|email[\s]?preferences|update[\s]?preferences|subscription[\s]?preferences)[^<]*)</a>'
        $matches1 = [regex]::Matches($emailBody, $pattern1, 'IgnoreCase')
        foreach ($match in $matches1) {
            $foundLinks += $match.Groups[1].Value
        }

        # Method 3: Look for links that contain "unsubscribe" in the URL itself
        $pattern2 = '<a[^>]+href\s*=\s*["' + "'" + ']([^"' + "'" + ']*(?:unsubscribe|optout|opt-out|remove|unsub)[^"' + "'" + ']*)["' + "'" + ']'
        $matches2 = [regex]::Matches($emailBody, $pattern2, 'IgnoreCase')
        foreach ($match in $matches2) {
            $foundLinks += $match.Groups[1].Value
        }

        # Method 4: Look for any href that appears near unsubscribe text
        $pattern3 = '(?:unsubscribe|opt[\s\-]?out|remove[\s]?me).{0,200}?href\s*=\s*["' + "'" + ']([^"' + "'" + ']+)["' + "'" + ']'
        $matches3 = [regex]::Matches($emailBody, $pattern3, 'IgnoreCase')
        foreach ($match in $matches3) {
            $foundLinks += $match.Groups[1].Value
        }

        # Method 5: Reverse - href before unsubscribe text
        $pattern4 = 'href\s*=\s*["' + "'" + ']([^"' + "'" + ']+)["' + "'" + '][^>]{0,200}?(?:unsubscribe|opt[\s\-]?out|remove[\s]?me)'
        $matches4 = [regex]::Matches($emailBody, $pattern4, 'IgnoreCase')
        foreach ($match in $matches4) {
            $foundLinks += $match.Groups[1].Value
        }

        # Clean and validate links
        $validLinks = @()
        foreach ($link in $foundLinks) {
            # Decode HTML entities
            $link = [System.Web.HttpUtility]::HtmlDecode($link)

            # Skip if empty or not a valid HTTP/HTTPS URL
            if (-not $link -or $link -notmatch '^https?://') {
                continue
            }

            # Skip common tracking pixels and non-unsubscribe links
            if ($link -match '\.(gif|jpg|jpeg|png|bmp)(\?|$)') {
                continue
            }

            # Skip mailto links
            if ($link -match '^mailto:') {
                continue
            }

            # Add to valid links if not already there
            if ($validLinks -notcontains $link) {
                $validLinks += $link
            }
        }

        # If we found valid unsubscribe links
        if ($validLinks.Count -gt 0) {
            Write-Host "Found $($validLinks.Count) Unsubscribe link(s) in email from $senderDomain" -ForegroundColor Green

            # Use the first valid link
            $unsubLink = $validLinks[0]
            $script:domainInfo[$senderDomain].UnsubLink = $unsubLink

            # Attempt one-click unsubscribe if enabled and supported
            if ($script:enableOneClick -and $listUnsubPost) {
                try {
                    Write-Host "  -> Attempting one-click unsubscribe for $senderDomain..." -ForegroundColor Yellow
                    $response = Invoke-WebRequest -Uri $listUnsubPost.Url -Method POST -Body $listUnsubPost.PostData -UseBasicParsing -ErrorAction Stop
                    if ($response.StatusCode -eq 200 -or $response.StatusCode -eq 204) {
                        Write-Host "  -> Successfully unsubscribed from $senderDomain!" -ForegroundColor Green
                        $script:oneClickUnsubscribed += $senderDomain
                    }
                } catch {
                    Write-Host "  -> One-click unsubscribe failed for $senderDomain (use manual link)" -ForegroundColor Red
                }
            }
        }
    } catch {
        Write-Verbose "Error processing email: $_"
    }
}

# Retrieve emails using Microsoft Graph API
Write-Host "`nFetching emails from $folderName..." -ForegroundColor Cyan

try {
    # Build the filter query for date range
    $filter = "receivedDateTime ge $filterDate"

    Write-Host "Retrieving emails (this may take a moment)..." -ForegroundColor Yellow

    # Fetch messages with expanded properties (note: 'size' is not available in basic properties)
    $messages = Get-MgUserMessage -UserId $userEmail -Filter $filter -Top 999 -All -Property "id,subject,from,sender,body,internetMessageHeaders,receivedDateTime"

    $totalEmails = $messages.Count
    Write-Host "Found $totalEmails emails to process." -ForegroundColor Green

    if ($totalEmails -eq 0) {
        Write-Host "No emails found in the specified date range." -ForegroundColor Yellow
    } else {
        # Process emails with progress bar
        $i = 0
        foreach ($email in $messages) {
            $i++
            Write-Progress -Activity "Processing Emails" -Status "$i out of $totalEmails Emails Processed" -PercentComplete (($i / $totalEmails) * 100)
            ProcessEmail $email
        }

        Write-Progress -Activity "Processing Emails" -Completed
    }

} catch {
    Write-Host "Error retrieving emails: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit
}

# Calculate statistics
$domainsWithUnsubscribe = ($domainInfo.GetEnumerator() | Where-Object { $_.Value.UnsubLink }).Count
$totalEmailsFromDomains = ($domainInfo.Values | Measure-Object -Property Count -Sum).Sum
$highPriorityCount = ($domainInfo.GetEnumerator() | Where-Object { $_.Value.Count -ge 10 }).Count
$mediumPriorityCount = ($domainInfo.GetEnumerator() | Where-Object { $_.Value.Count -ge 3 -and $_.Value.Count -lt 10 }).Count
$lowPriorityCount = ($domainInfo.GetEnumerator() | Where-Object { $_.Value.Count -lt 3 }).Count

Write-Host "`n=== STATISTICS ===" -ForegroundColor Cyan
Write-Host "Total emails processed: $totalEmails" -ForegroundColor Green
Write-Host "Unique domains found: $($domainInfo.Count)" -ForegroundColor Green
Write-Host "Domains with unsubscribe links: $domainsWithUnsubscribe" -ForegroundColor Green
Write-Host "High priority (10+ emails): $highPriorityCount" -ForegroundColor Red
Write-Host "Medium priority (3-9 emails): $mediumPriorityCount" -ForegroundColor Yellow
Write-Host "Low priority (1-2 emails): $lowPriorityCount" -ForegroundColor Green
if ($oneClickUnsubscribed.Count -gt 0) {
    Write-Host "One-click unsubscribed: $($oneClickUnsubscribed.Count) domains" -ForegroundColor Green
}

# Generate HTML Report with enhanced features
$script:htmlOutput = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Unsubscribe Links Report - Enhanced</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 40px 20px;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px;
            position: relative;
            overflow: hidden;
        }

        .header::before {
            content: '';
            position: absolute;
            top: -50%;
            right: -50%;
            width: 200%;
            height: 200%;
            background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 70%);
        }

        .header h1 {
            font-size: 2.5em;
            font-weight: 700;
            margin-bottom: 10px;
            position: relative;
            z-index: 1;
        }

        .header p {
            font-size: 1.1em;
            opacity: 0.9;
            position: relative;
            z-index: 1;
        }

        .stats-dashboard {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px 40px;
            background: #f8f9fa;
        }

        .stat-card {
            background: white;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.07);
            border-left: 4px solid #667eea;
            transition: transform 0.2s, box-shadow 0.2s;
        }

        .stat-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(0, 0, 0, 0.1);
        }

        .stat-card strong {
            display: block;
            color: #667eea;
            font-size: 0.85em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }

        .stat-card span {
            font-size: 1.5em;
            font-weight: 600;
            color: #2d3748;
        }

        .controls {
            padding: 20px 40px;
            background: white;
            border-bottom: 2px solid #e2e8f0;
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
            align-items: center;
        }

        .controls input[type="text"] {
            flex: 1;
            min-width: 250px;
            padding: 10px 15px;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            font-size: 1em;
        }

        .controls select, .controls button {
            padding: 10px 20px;
            border: 2px solid #667eea;
            border-radius: 8px;
            background: white;
            color: #667eea;
            font-size: 1em;
            cursor: pointer;
            transition: all 0.2s;
        }

        .controls button:hover {
            background: #667eea;
            color: white;
        }

        .table-container {
            padding: 40px;
            overflow-x: auto;
        }

        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 20px;
        }

        th, td {
            border: 1px solid #e2e8f0;
            text-align: left;
            padding: 16px;
        }

        th {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.85em;
            letter-spacing: 0.5px;
            cursor: pointer;
            user-select: none;
        }

        th:hover {
            opacity: 0.9;
        }

        tr:nth-child(even) {
            background-color: #f8f9fa;
        }

        tr:hover {
            background-color: #e6f2ff;
            transition: background-color 0.2s;
        }

        tr.clicked {
            opacity: 0.5;
            background-color: #e0e0e0 !important;
        }

        tr.one-click-unsubscribed {
            background-color: #d4edda !important;
        }

        .priority-badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 0.85em;
            font-weight: 600;
        }

        .priority-high {
            background: #fee;
            color: #c00;
        }

        .priority-medium {
            background: #ffeaa7;
            color: #d63031;
        }

        .priority-low {
            background: #dfe6e9;
            color: #2d3436;
        }

        .unsub-link {
            color: #667eea;
            text-decoration: none;
            font-weight: 500;
            padding: 8px 16px;
            background: #e6f2ff;
            border-radius: 6px;
            display: inline-block;
            transition: all 0.2s;
        }

        .unsub-link:hover {
            background: #667eea;
            color: white;
            transform: translateY(-1px);
            box-shadow: 0 2px 8px rgba(102, 126, 234, 0.3);
        }

        .email-subject {
            color: #718096;
            font-size: 0.9em;
            font-style: italic;
            margin-top: 5px;
        }

        .footer {
            padding: 30px 40px;
            background: #f8f9fa;
            border-top: 1px solid #e2e8f0;
            text-align: center;
            color: #718096;
            font-size: 0.9em;
        }

        .one-click-badge {
            display: inline-block;
            padding: 4px 8px;
            background: #d4edda;
            color: #155724;
            border-radius: 6px;
            font-size: 0.75em;
            font-weight: 600;
            margin-left: 8px;
        }

        @media (max-width: 768px) {
            .header h1 {
                font-size: 1.8em;
            }

            .stats-dashboard {
                grid-template-columns: 1fr;
            }

            .controls {
                flex-direction: column;
            }

            .controls input[type="text"] {
                width: 100%;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Unsubscribe Links Report</h1>
            <p>Enhanced with email counts, priority detection, and smart tracking</p>
        </div>

        <div class="stats-dashboard">
            <div class="stat-card">
                <strong>Total Emails</strong>
                <span>$totalEmails</span>
            </div>
            <div class="stat-card">
                <strong>Unique Domains</strong>
                <span>$($domainInfo.Count)</span>
            </div>
            <div class="stat-card">
                <strong>With Unsubscribe</strong>
                <span>$domainsWithUnsubscribe</span>
            </div>
            <div class="stat-card">
                <strong>High Priority</strong>
                <span style="color: #c00;">$highPriorityCount</span>
            </div>
            <div class="stat-card">
                <strong>Medium Priority</strong>
                <span style="color: #d63031;">$mediumPriorityCount</span>
            </div>
            <div class="stat-card">
                <strong>Total from Domains</strong>
                <span>$totalEmailsFromDomains</span>
            </div>
"@

if ($oneClickUnsubscribed.Count -gt 0) {
    $script:htmlOutput += @"
            <div class="stat-card">
                <strong>Auto-Unsubscribed</strong>
                <span style="color: #155724;">$($oneClickUnsubscribed.Count)</span>
            </div>
"@
}

$script:htmlOutput += @"
        </div>

        <div class="controls">
            <input type="text" id="searchBox" placeholder="Search domains or subjects..." onkeyup="filterTable()">
            <select id="sortSelect" onchange="sortTable()">
                <option value="priority">Sort by Priority</option>
                <option value="domain">Sort by Domain</option>
                <option value="count">Sort by Email Count</option>
            </select>
            <select id="filterSelect" onchange="filterTable()">
                <option value="all">Show All</option>
                <option value="high">High Priority Only</option>
                <option value="medium">Medium Priority Only</option>
                <option value="low">Low Priority Only</option>
                <option value="unclicked">Unclicked Only</option>
            </select>
            <button onclick="resetTracking()">Reset Tracking</button>
        </div>

        <div class="table-container">
            <table id="dataTable">
                <thead>
                    <tr>
                        <th onclick="sortByColumn(0)">Priority</th>
                        <th onclick="sortByColumn(1)">Domain</th>
                        <th onclick="sortByColumn(2)">Email Count</th>
                        <th onclick="sortByColumn(3)">Recent Subject</th>
                        <th onclick="sortByColumn(4)">Unsubscribe Link</th>
                    </tr>
                </thead>
                <tbody>
"@

# Sort domains by priority (count) descending
$sortedDomains = $domainInfo.GetEnumerator() | Where-Object { $_.Value.UnsubLink } | Sort-Object { $_.Value.Count } -Descending

foreach ($domain in $sortedDomains) {
    $domainName = $domain.Key
    $info = $domain.Value
    $count = $info.Count
    $unsubLink = $info.UnsubLink
    $subject = if ($info.Subject) { $info.Subject } else { "N/A" }
    $truncatedSubject = if ($subject.Length -gt 60) { $subject.Substring(0, 57) + "..." } else { $subject }

    # Determine priority
    $priorityClass = "priority-low"
    $priorityLabel = "LOW"
    $priorityEmoji = "&#x1F7E2;"  # Green circle
    if ($count -ge 10) {
        $priorityClass = "priority-high"
        $priorityLabel = "HIGH"
        $priorityEmoji = "&#x1F534;"  # Red circle
    } elseif ($count -ge 3) {
        $priorityClass = "priority-medium"
        $priorityLabel = "MEDIUM"
        $priorityEmoji = "&#x1F7E1;"  # Yellow circle
    }

    $truncatedUrl = if ($unsubLink.Length -gt 60) { $unsubLink.Substring(0, 57) + "..." } else { $unsubLink }

    # Check if one-click unsubscribed
    $rowClass = ""
    $oneClickBadge = ""
    if ($oneClickUnsubscribed -contains $domainName) {
        $rowClass = " class='one-click-unsubscribed'"
        $oneClickBadge = "<span class='one-click-badge'>&#x2713; AUTO-UNSUBSCRIBED</span>"
    }

    $script:htmlOutput += @"
                    <tr data-domain="$domainName" data-priority="$priorityLabel" data-count="$count"$rowClass>
                        <td><span class="priority-badge $priorityClass">$priorityEmoji $priorityLabel</span></td>
                        <td><strong>$domainName</strong></td>
                        <td><strong>$count</strong> emails</td>
                        <td><div class="email-subject">$truncatedSubject</div></td>
                        <td>
                            <a href="$unsubLink" class="unsub-link" target="_blank" onclick="markClicked(this)">Click to Unsubscribe</a>$oneClickBadge
                            <br><small style='color: #718096; font-size: 0.85em;'>$truncatedUrl</small>
                        </td>
                    </tr>
"@
}

$script:htmlOutput += @"
                </tbody>
            </table>
        </div>

        <div class="footer">
            <p><strong>Generated:</strong> $(Get-Date -Format "MMM dd, yyyy HH:mm")</p>
            <p style="margin-top: 10px;">Generated using Microsoft Graph API | Date Range: Last $daysToCheck days</p>
        </div>
    </div>

    <script>
        // Load clicked state from localStorage
        window.addEventListener('DOMContentLoaded', function() {
            loadClickedState();
        });

        function markClicked(link) {
            const row = link.closest('tr');
            const domain = row.getAttribute('data-domain');
            row.classList.add('clicked');

            // Save to localStorage
            let clicked = JSON.parse(localStorage.getItem('clickedDomains') || '[]');
            if (!clicked.includes(domain)) {
                clicked.push(domain);
                localStorage.setItem('clickedDomains', JSON.stringify(clicked));
            }
        }

        function loadClickedState() {
            const clicked = JSON.parse(localStorage.getItem('clickedDomains') || '[]');
            clicked.forEach(domain => {
                const row = document.querySelector('tr[data-domain="' + domain + '"]');
                if (row && !row.classList.contains('one-click-unsubscribed')) {
                    row.classList.add('clicked');
                }
            });
        }

        function resetTracking() {
            if (confirm('Reset all click tracking? This will mark all links as unclicked.')) {
                localStorage.removeItem('clickedDomains');
                document.querySelectorAll('tr.clicked').forEach(row => {
                    if (!row.classList.contains('one-click-unsubscribed')) {
                        row.classList.remove('clicked');
                    }
                });
            }
        }

        function filterTable() {
            const searchValue = document.getElementById('searchBox').value.toLowerCase();
            const filterValue = document.getElementById('filterSelect').value;
            const rows = document.querySelectorAll('#dataTable tbody tr');
            const clicked = JSON.parse(localStorage.getItem('clickedDomains') || '[]');

            rows.forEach(row => {
                const domain = row.getAttribute('data-domain').toLowerCase();
                const priority = row.getAttribute('data-priority');
                const subject = row.textContent.toLowerCase();

                let showRow = true;

                // Search filter
                if (searchValue && !domain.includes(searchValue) && !subject.includes(searchValue)) {
                    showRow = false;
                }

                // Priority filter
                if (filterValue === 'high' && priority !== 'HIGH') showRow = false;
                if (filterValue === 'medium' && priority !== 'MEDIUM') showRow = false;
                if (filterValue === 'low' && priority !== 'LOW') showRow = false;
                if (filterValue === 'unclicked' && (clicked.includes(row.getAttribute('data-domain')) || row.classList.contains('one-click-unsubscribed'))) {
                    showRow = false;
                }

                row.style.display = showRow ? '' : 'none';
            });
        }

        function sortTable() {
            const sortValue = document.getElementById('sortSelect').value;
            const tbody = document.querySelector('#dataTable tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));

            rows.sort((a, b) => {
                if (sortValue === 'domain') {
                    return a.getAttribute('data-domain').localeCompare(b.getAttribute('data-domain'));
                } else if (sortValue === 'count') {
                    return parseInt(b.getAttribute('data-count')) - parseInt(a.getAttribute('data-count'));
                } else if (sortValue === 'priority') {
                    const priorityOrder = { 'HIGH': 0, 'MEDIUM': 1, 'LOW': 2 };
                    return priorityOrder[a.getAttribute('data-priority')] - priorityOrder[b.getAttribute('data-priority')];
                }
            });

            rows.forEach(row => tbody.appendChild(row));
        }

        function sortByColumn(n) {
            // Trigger the appropriate sort
            const select = document.getElementById('sortSelect');
            if (n === 0) select.value = 'priority';
            else if (n === 1) select.value = 'domain';
            else if (n === 2) select.value = 'count';
            sortTable();
        }
    </script>
</body>
</html>
"@

# Write HTML output to file
$outputFile = Join-Path $scriptPath "unsubscribe-links-enhanced.html"
$script:htmlOutput | Out-File -FilePath $outputFile -Encoding UTF8

# Disconnect from Microsoft Graph
Disconnect-MgGraph | Out-Null

# Report completion
Write-Host "`n=== COMPLETED ===" -ForegroundColor Green
Write-Host "Report saved to: $outputFile" -ForegroundColor Cyan

# Optionally open the report
$openReport = Read-Host "`nWould you like to open the report now? (y/n)"
if ($openReport -match '^y') {
    Start-Process $outputFile
}
