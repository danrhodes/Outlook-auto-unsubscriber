/**
 * Main taskpane controller
 */

let scanner = null;
let currentResults = null;

// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Office.js ready');
        initializeApp();
    }
});

// Initialize the application
async function initializeApp() {
    try {
        // Initialize authentication
        const isAuthenticated = await AuthModule.initializeMsal();

        if (isAuthenticated) {
            showMainContent();
        } else {
            showAuthSection();
        }

        // Set up event listeners
        setupEventListeners();

        // Initialize scanner
        scanner = new EmailScanner();
        scanner.setProgressCallback(updateProgress);

    } catch (error) {
        console.error('Error initializing app:', error);
        showError('Failed to initialize application: ' + error.message);
    }
}

// Set up event listeners
function setupEventListeners() {
    // Sign in button
    document.getElementById('signin-button').addEventListener('click', async () => {
        try {
            await AuthModule.signIn();
            showMainContent();
        } catch (error) {
            showError('Sign in failed: ' + error.message);
        }
    });

    // Sign out button
    document.getElementById('signout-button').addEventListener('click', () => {
        AuthModule.signOut();
        location.reload();
    });

    // Scan all emails button
    document.getElementById('scan-all-button').addEventListener('click', async () => {
        await scanAllEmails();
    });

    // Scan selected email button
    document.getElementById('scan-selected-button').addEventListener('click', async () => {
        await scanSelectedEmail();
    });

    // Search box
    document.getElementById('search-box').addEventListener('input', filterResults);

    // Priority filter
    document.getElementById('priority-filter').addEventListener('change', filterResults);

    // Sort select
    document.getElementById('sort-select').addEventListener('change', sortResults);
}

// Show authentication section
function showAuthSection() {
    document.getElementById('auth-section').style.display = 'block';
    document.getElementById('main-content').style.display = 'none';
}

// Show main content
function showMainContent() {
    document.getElementById('auth-section').style.display = 'none';
    document.getElementById('main-content').style.display = 'block';

    const user = AuthModule.getCurrentUser();
    if (user) {
        console.log('Signed in as:', user.username);
    }
}

// Scan all emails
async function scanAllEmails() {
    try {
        showProgress('Initializing scan...');
        hideError();

        // Get access token
        const accessToken = await AuthModule.getAccessToken();

        // Initialize Graph client
        GraphService.initializeGraphClient(accessToken);

        // Get scan parameters
        const daysBack = parseInt(document.getElementById('days-input').value) || 30;
        const folder = document.getElementById('folder-select').value;

        // Fetch messages
        showProgress(`Fetching emails from ${folder}...`);
        const messages = await GraphService.getMessages(folder, daysBack);

        if (messages.length === 0) {
            showError('No emails found in the specified date range');
            hideProgress();
            return;
        }

        // Scan messages
        showProgress(`Scanning ${messages.length} emails...`);
        const results = await scanner.scanEmails(messages);

        // Display results
        displayResults(results);

        hideProgress();

    } catch (error) {
        console.error('Error scanning emails:', error);
        showError('Error scanning emails: ' + error.message);
        hideProgress();
    }
}

// Scan selected email
async function scanSelectedEmail() {
    try {
        showProgress('Scanning selected email...');
        hideError();

        // Get access token
        const accessToken = await AuthModule.getAccessToken();

        // Initialize Graph client
        GraphService.initializeGraphClient(accessToken);

        // Get current message
        const messages = await GraphService.getCurrentMessage();

        if (messages.length === 0) {
            showError('No email selected or unable to access selected email');
            hideProgress();
            return;
        }

        // Scan messages
        const results = await scanner.scanEmails(messages);

        // Display results
        displayResults(results);

        hideProgress();

    } catch (error) {
        console.error('Error scanning selected email:', error);
        showError('Error scanning selected email: ' + error.message);
        hideProgress();
    }
}

// Display results
function displayResults(results) {
    currentResults = results;

    // Update statistics
    document.getElementById('stat-total-emails').textContent = results.stats.totalEmails;
    document.getElementById('stat-unique-domains').textContent = results.stats.uniqueDomains;
    document.getElementById('stat-with-unsubscribe').textContent = results.stats.domainsWithUnsubscribe;
    document.getElementById('stat-high-priority').textContent = results.stats.highPriority;

    // Show sections
    document.getElementById('stats-section').style.display = 'block';
    document.getElementById('filters-section').style.display = 'block';
    document.getElementById('results-section').style.display = 'block';

    // Render results table
    renderResultsTable(results.domains);

    // Load clicked state from localStorage
    loadClickedState();
}

// Render results table
function renderResultsTable(domains) {
    const tbody = document.getElementById('results-body');
    tbody.innerHTML = '';

    if (domains.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align: center; padding: 40px;">No unsubscribe links found</td></tr>';
        return;
    }

    domains.forEach(domain => {
        const row = document.createElement('tr');
        row.setAttribute('data-domain', domain.domain);
        row.setAttribute('data-priority', domain.priority);
        row.setAttribute('data-count', domain.count);

        const priorityInfo = scanner.getPriorityInfo(domain.priority);
        const truncatedSubject = domain.subject.length > 50
            ? domain.subject.substring(0, 47) + '...'
            : domain.subject;
        const truncatedUrl = domain.unsubLink.length > 50
            ? domain.unsubLink.substring(0, 47) + '...'
            : domain.unsubLink;

        row.innerHTML = `
            <td><span class="priority-badge ${priorityInfo.class}">${priorityInfo.emoji} ${priorityInfo.label}</span></td>
            <td><strong>${escapeHtml(domain.domain)}</strong></td>
            <td><strong>${domain.count}</strong> emails</td>
            <td><div class="email-subject">${escapeHtml(truncatedSubject)}</div></td>
            <td>
                <a href="${escapeHtml(domain.unsubLink)}" class="unsub-link" target="_blank" onclick="markClicked(this)">
                    Unsubscribe
                </a>
                <span class="link-preview">${escapeHtml(truncatedUrl)}</span>
            </td>
        `;

        tbody.appendChild(row);
    });
}

// Mark link as clicked
window.markClicked = function(link) {
    const row = link.closest('tr');
    const domain = row.getAttribute('data-domain');
    row.classList.add('clicked');

    // Save to localStorage
    let clicked = JSON.parse(localStorage.getItem('clickedDomains') || '[]');
    if (!clicked.includes(domain)) {
        clicked.push(domain);
        localStorage.setItem('clickedDomains', JSON.stringify(clicked));
    }
};

// Load clicked state from localStorage
function loadClickedState() {
    const clicked = JSON.parse(localStorage.getItem('clickedDomains') || '[]');
    clicked.forEach(domain => {
        const row = document.querySelector(`tr[data-domain="${domain}"]`);
        if (row) {
            row.classList.add('clicked');
        }
    });
}

// Filter results
function filterResults() {
    if (!currentResults) return;

    const searchValue = document.getElementById('search-box').value.toLowerCase();
    const priorityFilter = document.getElementById('priority-filter').value;

    const filteredDomains = currentResults.domains.filter(domain => {
        // Search filter
        if (searchValue) {
            const matchesDomain = domain.domain.toLowerCase().includes(searchValue);
            const matchesSubject = domain.subject.toLowerCase().includes(searchValue);
            if (!matchesDomain && !matchesSubject) {
                return false;
            }
        }

        // Priority filter
        if (priorityFilter !== 'all' && domain.priority !== priorityFilter) {
            return false;
        }

        return true;
    });

    renderResultsTable(filteredDomains);
    loadClickedState();
}

// Sort results
function sortResults() {
    if (!currentResults) return;

    const sortBy = document.getElementById('sort-select').value;
    const sortedDomains = [...currentResults.domains];

    switch (sortBy) {
        case 'domain':
            sortedDomains.sort((a, b) => a.domain.localeCompare(b.domain));
            break;
        case 'count':
            sortedDomains.sort((a, b) => b.count - a.count);
            break;
        case 'priority':
            const priorityOrder = { high: 0, medium: 1, low: 2 };
            sortedDomains.sort((a, b) => priorityOrder[a.priority] - priorityOrder[b.priority]);
            break;
    }

    renderResultsTable(sortedDomains);
    loadClickedState();
}

// Show progress
function showProgress(message) {
    document.getElementById('progress-section').style.display = 'block';
    document.getElementById('progress-text').textContent = message;
    document.getElementById('progress-fill').style.width = '0%';
}

// Update progress
function updateProgress(percentage, current, total) {
    document.getElementById('progress-text').textContent = `Scanning emails... ${current} of ${total}`;
    document.getElementById('progress-fill').style.width = percentage + '%';
}

// Hide progress
function hideProgress() {
    document.getElementById('progress-section').style.display = 'none';
}

// Show error
function showError(message) {
    document.getElementById('error-section').style.display = 'block';
    document.getElementById('error-text').textContent = message;
}

// Hide error
function hideError() {
    document.getElementById('error-section').style.display = 'none';
}

// Escape HTML
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
