/**
 * Microsoft Graph API service
 */

let graphClient = null;

// Initialize Graph client
function initializeGraphClient(accessToken) {
    graphClient = MicrosoftGraph.Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}

// Get user's email address
async function getUserEmail() {
    try {
        const user = await graphClient
            .api('/me')
            .select('mail,userPrincipalName')
            .get();
        return user.mail || user.userPrincipalName;
    } catch (error) {
        console.error('Error getting user email:', error);
        throw error;
    }
}

// Get messages from a specific folder
async function getMessages(folderName, daysBack) {
    try {
        const startDate = new Date();
        startDate.setDate(startDate.getDate() - daysBack);
        const filterDate = startDate.toISOString();

        let endpoint = '';
        switch (folderName.toLowerCase()) {
            case 'inbox':
                endpoint = '/me/mailFolders/inbox/messages';
                break;
            case 'sentitems':
                endpoint = '/me/mailFolders/sentitems/messages';
                break;
            case 'junkemail':
                endpoint = '/me/mailFolders/junkemail/messages';
                break;
            case 'allmail':
                endpoint = '/me/messages';
                break;
            default:
                endpoint = '/me/mailFolders/inbox/messages';
        }

        const filter = `receivedDateTime ge ${filterDate}`;

        let messages = [];
        let nextLink = null;

        // Fetch messages with pagination
        do {
            const response = await graphClient
                .api(nextLink || endpoint)
                .filter(filter)
                .select('id,subject,from,sender,body,receivedDateTime,internetMessageHeaders')
                .top(100)
                .get();

            messages = messages.concat(response.value);
            nextLink = response['@odata.nextLink'];

            // Limit to prevent excessive API calls (max 1000 emails)
            if (messages.length >= 1000) {
                break;
            }
        } while (nextLink);

        return messages;
    } catch (error) {
        console.error('Error fetching messages:', error);
        throw error;
    }
}

// Get currently selected message in Outlook
async function getCurrentMessage() {
    return new Promise((resolve, reject) => {
        if (!Office.context.mailbox.item) {
            reject(new Error('No email selected'));
            return;
        }

        const item = Office.context.mailbox.item;

        // Get full message via Graph API
        const messageId = item.itemId;

        // Convert Outlook REST ID to Graph API format
        const graphMessageId = Office.context.mailbox.convertToRestId(
            messageId,
            Office.MailboxEnums.RestVersion.v2_0
        );

        graphClient
            .api(`/me/messages/${graphMessageId}`)
            .select('id,subject,from,sender,body,receivedDateTime,internetMessageHeaders')
            .get()
            .then(message => resolve([message]))
            .catch(error => reject(error));
    });
}

// Export functions
window.GraphService = {
    initializeGraphClient,
    getUserEmail,
    getMessages,
    getCurrentMessage
};
