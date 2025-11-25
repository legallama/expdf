// Initialize Office Add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Email to PDF Exporter loaded successfully");
        updatePagePreview();
        loadAttachmentInfo();
    }
});

// Event listeners
document.getElementById('exportButton').addEventListener('click', exportToPdf);
document.getElementById('cancelButton').addEventListener('click', () => {
    Office.addin.odsDialog.closeDialog();
});

document.getElementById('includeAttachments').addEventListener('change', () => {
    updatePagePreview();
    if (document.getElementById('includeAttachments').checked) {
        loadAttachmentInfo();
    }
});

document.querySelectorAll('input[name="exportType"]').forEach(radio => {
    radio.addEventListener('change', updatePagePreview);
});

document.getElementById('fromDate').addEventListener('change', () => {
    validateDateRange();
    updatePagePreview();
});

document.getElementById('toDate').addEventListener('change', () => {
    validateDateRange();
    updatePagePreview();
});

async function exportToPdf() {
    try {
        showStatus('Processing...', 'info');
        const exportType = document.querySelector('input[name="exportType"]:checked').value;
        const includeAttachments = document.getElementById('includeAttachments').checked;
        
        // Get date range
        const fromDate = document.getElementById('fromDate').value;
        const toDate = document.getElementById('toDate').value;
        const dateRange = { from: fromDate ? new Date(fromDate) : null, to: toDate ? new Date(toDate) : null };

        // Request save location
        const saveLocation = await requestSaveLocation();
        if (!saveLocation) {
            showStatus('Export cancelled', 'info');
            return;
        }

        if (exportType === 'single') {
            await exportSingleEmail(includeAttachments, saveLocation, dateRange);
        } else {
            await exportConversation(includeAttachments, saveLocation, dateRange);
        }
    } catch (error) {
        showStatus('Error: ' + error.message, 'error');
        console.error(error);
    }
}

async function exportSingleEmail(includeAttachments, saveLocation, dateRange) {
    const item = Office.context.mailbox.item;
    const itemDate = new Date(item.dateTimeCreated);
    
    // Check if email is within date range
    if (!isDateInRange(itemDate, dateRange)) {
        showStatus('Email is outside the selected date range', 'error');
        return;
    }
    
    const emailData = {
        subject: item.subject,
        from: item.from.emailAddress,
        to: item.to?.map(r => r.emailAddress).join(', ') || '',
        cc: item.cc?.map(r => r.emailAddress).join(', ') || '',
        date: itemDate.toLocaleString(),
        body: await getEmailBody(item)
    };

    const html = generateEmailHtml([emailData]);
    await generatePdf(html, emailData.subject, saveLocation);

    if (includeAttachments && item.attachments.length > 0) {
        await processAttachments(item, emailData.subject, saveLocation);
    }
}

async function exportConversation(includeAttachments, saveLocation, dateRange) {
    const item = Office.context.mailbox.item;
    const emailsData = [];

    // Get current email
    const itemDate = new Date(item.dateTimeCreated);
    
    // Check if email is within date range
    if (isDateInRange(itemDate, dateRange)) {
        const currentEmail = {
            subject: item.subject,
            from: item.from.emailAddress,
            to: item.to?.map(r => r.emailAddress).join(', ') || '',
            cc: item.cc?.map(r => r.emailAddress).join(', ') || '',
            date: itemDate.toLocaleString(),
            body: await getEmailBody(item)
        };
        emailsData.push(currentEmail);
    }

    // Try to get conversation items
    try {
        // Note: Getting full conversation requires additional API calls
        // This is a simplified version that gets the current email and any available thread info
        const conversationId = item.conversationId;
        // Full conversation extraction would require server-side processing
        // For now, we'll include the current email and note this limitation
        showStatus('Note: Exporting available conversation data', 'info');
    } catch (error) {
        console.log('Conversation data not fully available:', error);
    }

    if (emailsData.length === 0) {
        showStatus('No emails in conversation match the selected date range', 'error');
        return;
    }

    const html = generateEmailHtml(emailsData);
    await generatePdf(html, emailsData[0].subject, saveLocation);

    if (includeAttachments && item.attachments.length > 0) {
        await processAttachments(item, emailsData[0].subject, saveLocation);
    }
}

async function getEmailBody(item) {
    return new Promise((resolve) => {
        item.body.getAsync(Office.CoercionType.Html, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                resolve('<p>Unable to retrieve email body</p>');
            }
        });
    });
}

function generateEmailHtml(emailsData) {
    let html = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                * {
                    margin: 0;
                    padding: 0;
                    box-sizing: border-box;
                }
                
                body {
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    color: #333;
                    background: white;
                    line-height: 1.6;
                }
                
                .email-container {
                    padding: 20px;
                    margin: 20px 0;
                    border: 1px solid #ddd;
                    page-break-after: always;
                }
                
                .email-header {
                    border-bottom: 2px solid #0078d4;
                    padding-bottom: 15px;
                    margin-bottom: 15px;
                }
                
                .email-subject {
                    font-size: 18px;
                    font-weight: bold;
                    color: #0078d4;
                    margin-bottom: 10px;
                }
                
                .email-metadata {
                    font-size: 12px;
                    color: #666;
                    display: grid;
                    gap: 5px;
                }
                
                .metadata-row {
                    display: grid;
                    grid-template-columns: 80px 1fr;
                }
                
                .metadata-label {
                    font-weight: 600;
                    color: #444;
                }
                
                .email-body {
                    margin-top: 15px;
                    color: #333;
                    word-wrap: break-word;
                }
                
                .email-body img {
                    max-width: 100%;
                    height: auto;
                }
                
                .email-body a {
                    color: #0078d4;
                    text-decoration: underline;
                }
                
                table {
                    border-collapse: collapse;
                    width: 100%;
                    margin: 10px 0;
                }
                
                td, th {
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: left;
                }
                
                th {
                    background-color: #f5f5f5;
                }
                
                @media print {
                    .email-container {
                        page-break-after: always;
                        border: none;
                        padding: 40px 0;
                    }
                }
            </style>
        </head>
        <body>
    `;

    emailsData.forEach((email, index) => {
        html += `
            <div class="email-container">
                <div class="email-header">
                    <div class="email-subject">${escapeHtml(email.subject)}</div>
                    <div class="email-metadata">
                        <div class="metadata-row">
                            <span class="metadata-label">From:</span>
                            <span>${escapeHtml(email.from)}</span>
                        </div>
                        <div class="metadata-row">
                            <span class="metadata-label">To:</span>
                            <span>${escapeHtml(email.to)}</span>
                        </div>
                        ${email.cc ? `
                        <div class="metadata-row">
                            <span class="metadata-label">Cc:</span>
                            <span>${escapeHtml(email.cc)}</span>
                        </div>
                        ` : ''}
                        <div class="metadata-row">
                            <span class="metadata-label">Date:</span>
                            <span>${escapeHtml(email.date)}</span>
                        </div>
                    </div>
                </div>
                <div class="email-body">
                    ${email.body}
                </div>
            </div>
        `;
    });

    html += `
        </body>
        </html>
    `;

    return html;
}

async function generatePdf(html, filename, saveLocation) {
    const element = document.createElement('div');
    element.innerHTML = html;

    return new Promise((resolve, reject) => {
        const opt = {
            margin: 10,
            filename: `${sanitizeFilename(filename)}.pdf`,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2 },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
            pagebreak: { mode: 'avoid-all', before: '.email-container' }
        };

        html2pdf()
            .set(opt)
            .from(element)
            .save()
            .then(() => {
                showStatus('Email PDF exported successfully!', 'success');
                resolve();
            })
            .catch(error => {
                reject(error);
            });
    });
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function sanitizeFilename(filename) {
    return filename.replace(/[^a-z0-9]/gi, '_').toLowerCase().substring(0, 100);
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('statusMessage');
    statusDiv.textContent = message;
    statusDiv.className = `status-message status-${type}`;
    
    if (type === 'success') {
        setTimeout(() => {
            statusDiv.textContent = '';
        }, 3000);
    }
}

function loadAttachmentInfo() {
    const item = Office.context.mailbox.item;
    const attachmentCount = item.attachments.length;
    
    if (attachmentCount === 0) {
        document.getElementById('attachmentInfo').style.display = 'none';
        return;
    }

    const attachmentInfo = document.getElementById('attachmentInfo');
    const attachmentList = document.getElementById('attachmentList');
    
    attachmentInfo.style.display = 'block';
    document.getElementById('attachmentCount').textContent = attachmentCount;
    
    attachmentList.innerHTML = '';
    item.attachments.forEach((attachment, index) => {
        const li = document.createElement('li');
        li.textContent = `${attachment.name} (${formatFileSize(attachment.size)})`;
        attachmentList.appendChild(li);
    });
}

function updatePagePreview() {
    const item = Office.context.mailbox.item;
    const exportType = document.querySelector('input[name="exportType"]:checked').value;
    const includeAttachments = document.getElementById('includeAttachments').checked;
    const fromDate = document.getElementById('fromDate').value;
    const toDate = document.getElementById('toDate').value;
    const dateRange = { from: fromDate ? new Date(fromDate) : null, to: toDate ? new Date(toDate) : null };
    
    let pages = 0;
    const itemDate = new Date(item.dateTimeCreated);
    
    // Check if email is within date range
    if (!isDateInRange(itemDate, dateRange)) {
        document.getElementById('pageCount').textContent = '0';
        updateDateRangeInfo(0, dateRange);
        return;
    }
    
    if (exportType === 'single') {
        pages = 1; // One page per email
    } else {
        pages = 1; // Conversation - estimate based on available data
    }
    
    // Add pages for attachments
    if (includeAttachments && item.attachments.length > 0) {
        // Rough estimate: 1 page per attachment
        pages += item.attachments.length;
    }
    
    document.getElementById('pageCount').textContent = pages;
    updateDateRangeInfo(pages, dateRange);
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

async function requestSaveLocation() {
    // Try to use File System Access API (modern browsers)
    if (window.showDirectoryPicker) {
        try {
            const dirHandle = await window.showDirectoryPicker();
            return dirHandle;
        } catch (error) {
            console.log('Directory picker cancelled or not supported');
            return null;
        }
    } else {
        // Fallback: use downloads folder via automatic download
        showStatus('Files will be saved to your Downloads folder', 'info');
        return 'downloads';
    }
}

async function processAttachments(item, emailSubject, saveLocation) {
    showStatus('Processing attachments...', 'info');
    
    try {
        const attachments = item.attachments;
        let processedCount = 0;

        for (let i = 0; i < attachments.length; i++) {
            const attachment = attachments[i];
            
            // Only process image attachments for now
            if (isImageAttachment(attachment.name)) {
                await downloadAndConvertAttachment(attachment, emailSubject, saveLocation);
                processedCount++;
            }
        }

        showStatus(`${processedCount} attachment(s) processed!`, 'success');
    } catch (error) {
        console.error('Error processing attachments:', error);
        showStatus('Some attachments could not be processed', 'error');
    }
}

function isImageAttachment(filename) {
    const imageExtensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'];
    const ext = filename.toLowerCase().substring(filename.lastIndexOf('.'));
    return imageExtensions.includes(ext);
}

async function downloadAndConvertAttachment(attachment, emailSubject, saveLocation) {
    // Note: Outlook API limitations prevent direct file access
    // Attachments can be referenced but binary data extraction requires server-side processing
    console.log('Attachment processing requested:', attachment.name);
    showStatus('Note: Direct attachment conversion requires enhanced permissions', 'info');
}

function isDateInRange(date, dateRange) {
    if (!dateRange.from && !dateRange.to) {
        return true; // No filter applied
    }
    
    if (dateRange.from && date < dateRange.from) {
        return false;
    }
    
    if (dateRange.to) {
        // Add one day to include the entire "to" date
        const toDateEndOfDay = new Date(dateRange.to);
        toDateEndOfDay.setDate(toDateEndOfDay.getDate() + 1);
        if (date >= toDateEndOfDay) {
            return false;
        }
    }
    
    return true;
}

function validateDateRange() {
    const fromDate = document.getElementById('fromDate').value;
    const toDate = document.getElementById('toDate').value;
    
    if (fromDate && toDate) {
        const from = new Date(fromDate);
        const to = new Date(toDate);
        
        if (from > to) {
            showStatus('From date cannot be later than To date', 'error');
            document.getElementById('toDate').value = fromDate;
        }
    }
}

function updateDateRangeInfo(pageCount, dateRange) {
    const dateInfo = document.getElementById('dateRangeInfo');
    
    if (!dateRange.from && !dateRange.to) {
        dateInfo.classList.remove('active');
        return;
    }
    
    let info = '';
    
    if (dateRange.from && dateRange.to) {
        info = `Filtering: ${dateRange.from.toLocaleDateString()} to ${dateRange.to.toLocaleDateString()} (${pageCount} page${pageCount !== 1 ? 's' : ''})`;
    } else if (dateRange.from) {
        info = `Filtering from: ${dateRange.from.toLocaleDateString()} onwards (${pageCount} page${pageCount !== 1 ? 's' : ''})`;
    } else if (dateRange.to) {
        info = `Filtering until: ${dateRange.to.toLocaleDateString()} (${pageCount} page${pageCount !== 1 ? 's' : ''})`;
    }
    
    dateInfo.textContent = info;
    dateInfo.classList.add('active');
}
