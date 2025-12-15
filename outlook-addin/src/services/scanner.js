/**
 * Email scanning service - finds unsubscribe links in emails
 * Ported from PowerShell script logic
 */

class EmailScanner {
    constructor() {
        this.domainInfo = new Map();
        this.progressCallback = null;
    }

    // Set progress callback function
    setProgressCallback(callback) {
        this.progressCallback = callback;
    }

    // Main scan function
    async scanEmails(messages) {
        this.domainInfo.clear();
        const totalEmails = messages.length;

        for (let i = 0; i < messages.length; i++) {
            await this.processEmail(messages[i]);

            // Update progress
            if (this.progressCallback) {
                const progress = ((i + 1) / totalEmails) * 100;
                this.progressCallback(progress, i + 1, totalEmails);
            }
        }

        return this.generateReport();
    }

    // Process a single email
    async processEmail(email) {
        try {
            // Extract sender information
            let senderEmail = null;
            if (email.from && email.from.emailAddress && email.from.emailAddress.address) {
                senderEmail = email.from.emailAddress.address;
            } else if (email.sender && email.sender.emailAddress && email.sender.emailAddress.address) {
                senderEmail = email.sender.emailAddress.address;
            }

            if (!senderEmail) {
                return;
            }

            // Extract domain from sender email
            const emailParts = senderEmail.split('@');
            if (emailParts.length < 2) {
                return;
            }

            const senderDomain = emailParts[1].toLowerCase();

            // Initialize or update domain info
            if (!this.domainInfo.has(senderDomain)) {
                this.domainInfo.set(senderDomain, {
                    count: 0,
                    unsubLink: null,
                    subject: null,
                    receivedDate: null,
                    listUnsubPost: null
                });
            }

            const info = this.domainInfo.get(senderDomain);
            info.count++;

            // Store the most recent subject and date if not already set
            if (!info.subject) {
                info.subject = email.subject || 'N/A';
                info.receivedDate = email.receivedDateTime;
            }

            // Skip further processing if we already have an unsubscribe link for this domain
            if (info.unsubLink) {
                return;
            }

            // Get email body
            const emailBody = email.body && email.body.content ? email.body.content : '';
            if (!emailBody) {
                return;
            }

            // Find unsubscribe links
            const foundLinks = this.findUnsubscribeLinks(email, emailBody);

            if (foundLinks.length > 0) {
                info.unsubLink = foundLinks[0];
                console.log(`Found ${foundLinks.length} unsubscribe link(s) in email from ${senderDomain}`);
            }

        } catch (error) {
            console.error('Error processing email:', error);
        }
    }

    // Find unsubscribe links using multiple methods
    findUnsubscribeLinks(email, emailBody) {
        const foundLinks = [];

        // Method 1: Check List-Unsubscribe header (RFC 2369) - Most reliable
        if (email.internetMessageHeaders) {
            const unsubHeader = email.internetMessageHeaders.find(h => h.name === 'List-Unsubscribe');
            if (unsubHeader) {
                const headerMatches = unsubHeader.value.match(/<(https?:\/\/[^>]+)>/g);
                if (headerMatches) {
                    headerMatches.forEach(match => {
                        const url = match.replace(/[<>]/g, '');
                        foundLinks.push(url);
                    });
                }
            }

            // Check for List-Unsubscribe-Post header (RFC 8058 - One-Click)
            const unsubPostHeader = email.internetMessageHeaders.find(h => h.name === 'List-Unsubscribe-Post');
            if (unsubPostHeader && unsubHeader) {
                // Store for potential one-click unsubscribe implementation
                console.log('One-click unsubscribe supported for this email');
            }
        }

        // Method 2: Look for <a> tags with unsubscribe-related text
        const pattern1 = /<a[^>]+href\s*=\s*["']([^"']+)["'][^>]*>(?:[^<]*(?:unsubscribe|opt[\s\-]?out|remove[\s]?me|manage[\s]?preferences|email[\s]?preferences)[^<]*)<\/a>/gi;
        const matches1 = emailBody.matchAll(pattern1);
        for (const match of matches1) {
            foundLinks.push(match[1]);
        }

        // Method 3: Look for links that contain "unsubscribe" in the URL
        const pattern2 = /<a[^>]+href\s*=\s*["']([^"']*(?:unsubscribe|optout|opt-out|remove|unsub)[^"']*)["']/gi;
        const matches2 = emailBody.matchAll(pattern2);
        for (const match of matches2) {
            foundLinks.push(match[1]);
        }

        // Method 4: Look for any href near unsubscribe text
        const pattern3 = /(?:unsubscribe|opt[\s\-]?out|remove[\s]?me).{0,200}?href\s*=\s*["']([^"']+)["']/gi;
        const matches3 = emailBody.matchAll(pattern3);
        for (const match of matches3) {
            foundLinks.push(match[1]);
        }

        // Method 5: Reverse - href before unsubscribe text
        const pattern4 = /href\s*=\s*["']([^"']+)["'][^>]{0,200}?(?:unsubscribe|opt[\s\-]?out|remove[\s]?me)/gi;
        const matches4 = emailBody.matchAll(pattern4);
        for (const match of matches4) {
            foundLinks.push(match[1]);
        }

        // Clean and validate links
        return this.cleanAndValidateLinks(foundLinks);
    }

    // Clean and validate unsubscribe links
    cleanAndValidateLinks(links) {
        const validLinks = [];
        const seen = new Set();

        for (let link of links) {
            // Decode HTML entities
            link = this.decodeHtmlEntities(link);

            // Skip if not a valid HTTP/HTTPS URL
            if (!link || !link.match(/^https?:\/\//i)) {
                continue;
            }

            // Skip common tracking pixels and non-unsubscribe links
            if (link.match(/\.(gif|jpg|jpeg|png|bmp)(\?|$)/i)) {
                continue;
            }

            // Skip mailto links
            if (link.match(/^mailto:/i)) {
                continue;
            }

            // Add to valid links if not already there
            if (!seen.has(link)) {
                validLinks.push(link);
                seen.add(link);
            }
        }

        return validLinks;
    }

    // Decode HTML entities
    decodeHtmlEntities(text) {
        const textarea = document.createElement('textarea');
        textarea.innerHTML = text;
        return textarea.value;
    }

    // Generate report from scanned data
    generateReport() {
        const domainsArray = Array.from(this.domainInfo.entries())
            .filter(([domain, info]) => info.unsubLink)
            .map(([domain, info]) => ({
                domain,
                count: info.count,
                unsubLink: info.unsubLink,
                subject: info.subject,
                receivedDate: info.receivedDate,
                priority: this.calculatePriority(info.count)
            }))
            .sort((a, b) => b.count - a.count);

        const stats = {
            totalEmails: Array.from(this.domainInfo.values()).reduce((sum, info) => sum + info.count, 0),
            uniqueDomains: this.domainInfo.size,
            domainsWithUnsubscribe: domainsArray.length,
            highPriority: domainsArray.filter(d => d.priority === 'high').length,
            mediumPriority: domainsArray.filter(d => d.priority === 'medium').length,
            lowPriority: domainsArray.filter(d => d.priority === 'low').length
        };

        return {
            domains: domainsArray,
            stats
        };
    }

    // Calculate priority based on email count
    calculatePriority(count) {
        if (count >= 10) return 'high';
        if (count >= 3) return 'medium';
        return 'low';
    }

    // Get priority display info
    getPriorityInfo(priority) {
        const info = {
            high: { label: 'HIGH', emoji: '🔴', class: 'priority-high' },
            medium: { label: 'MEDIUM', emoji: '🟡', class: 'priority-medium' },
            low: { label: 'LOW', emoji: '🟢', class: 'priority-low' }
        };
        return info[priority] || info.low;
    }
}

// Export
window.EmailScanner = EmailScanner;
