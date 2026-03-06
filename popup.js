document.addEventListener('DOMContentLoaded', () => {
    const scrapeBtn = document.getElementById('scrapeBtn');
    const scrapeAllBtn = document.getElementById('scrapeAllBtn');
    const statusDiv = document.getElementById('status');

    function setStatus(message, type = 'info') {
        statusDiv.textContent = message;
        statusDiv.className = type;
    }

    scrapeBtn.addEventListener('click', async () => {
        setStatus('Checking current tab...', 'info');
        scrapeBtn.disabled = true;

        try {
            // Get the active tab
            let [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

            if (!tab.url || !tab.url.includes('amazon.com')) {
                setStatus('Please navigate to Amazon "Your Orders" page first.', 'error');
                scrapeBtn.disabled = false;
                return;
            }
            if (!tab.url.includes('your-orders') && !tab.url.includes('order-history')) {
                setStatus('Warning: Not heavily tested on this Amazon page.', 'error');
                // Don't disable immediately, let them try it.
            }

            setStatus('Injecting script...', 'info');

            // Inject the content script
            await chrome.scripting.executeScript({
                target: { tabId: tab.id },
                files: ['content.js']
            });

            setStatus('Scraping data...', 'info');

            // Listen for the response from the content script
            chrome.tabs.sendMessage(tab.id, { action: "scrape_orders" }, (response) => {
                if (chrome.runtime.lastError) {
                    console.error(chrome.runtime.lastError);
                    setStatus('Error communicating with page. Try refreshing.', 'error');
                    scrapeBtn.disabled = false;
                    return;
                }

                if (response && response.success) {
                    if (!response.data || response.data.length === 0) {
                        setStatus('No orders found on this page.', 'error');
                        scrapeBtn.disabled = false;
                        return;
                    }

                    setStatus(`Found ${response.data.length} orders. Downloading...`, 'success');
                    downloadXLSX(response.data);

                    setTimeout(() => {
                        setStatus('Done!', 'success');
                        scrapeBtn.disabled = false;
                    }, 1500);
                } else {
                    setStatus(response?.error || 'Unknown error occurred.', 'error');
                    scrapeBtn.disabled = false;
                }
            });

        } catch (error) {
            console.error('Scraping error:', error);
            setStatus('An error occurred. See console.', 'error');
            scrapeBtn.disabled = false;
        }
    });

    function parseDateToFormattedString(dateStr) {
        if (!dateStr || dateStr === 'N/A') return dateStr;
        const parsedDate = new Date(dateStr);
        if (isNaN(parsedDate)) return dateStr;
        const mm = String(parsedDate.getMonth() + 1).padStart(2, '0');
        const dd = String(parsedDate.getDate()).padStart(2, '0');
        const yyyy = parsedDate.getFullYear();
        return `${mm}/${dd}/${yyyy}`;
    }

    function computeDaysUntilDeadline(returnDateStr) {
        if (!returnDateStr || returnDateStr === 'N/A') return '';
        const returnDate = new Date(returnDateStr);
        if (isNaN(returnDate)) return '';
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        returnDate.setHours(0, 0, 0, 0);
        return Math.ceil((returnDate - today) / (1000 * 60 * 60 * 24));
    }

    function computeStatus(returnEligible, returnDateStr, orderStatus) {
        if (orderStatus === 'Refunded') return 'Returned';
        if (orderStatus === 'Not Yet Arrived') return 'Not Yet Arrived';
        if (returnEligible === 'Yes') {
            const days = computeDaysUntilDeadline(returnDateStr);
            if (days === '') return 'Unknown';
            if (days <= 7) return 'Urgent - Return Soon';
            return 'Eligible';
        } else if (returnDateStr && returnDateStr !== 'N/A') {
            return 'Window Closed';
        }
        return 'Unknown';
    }

    async function downloadXLSX(data) {
        if (!data || !data.length) return;

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Amazon Orders');

        const rawRows = data.map(row => [
            computeStatus(row.returnEligible, row.returnDate, row.orderStatus || ''),
            computeDaysUntilDeadline(row.returnDate),
            parseDateToFormattedString(row.date),
            row.total,
            row.orderId,
            row.items,
            row.asins || '',
            row.shipTo || '',
            row.link ? { text: row.link, hyperlink: row.link, tooltip: row.link } : '',
            row.returnEligible,
            parseDateToFormattedString(row.returnDate),
            ''
        ]);

        worksheet.addTable({
            name: 'OrdersTable',
            ref: 'A1',
            headerRow: true,
            totalsRow: false,
            style: {
                theme: 'TableStyleMedium2',
                showRowStripes: true,
            },
            columns: [
                { name: 'Status' },
                { name: 'Days Until Deadline' },
                { name: 'Order Date' },
                { name: 'Order Total' },
                { name: 'Order Number' },
                { name: 'Items' },
                { name: 'ASINs' },
                { name: 'Ship To' },
                { name: 'Order Link' },
                { name: 'Return Eligible' },
                { name: 'Return Date' },
                { name: 'Notes' }
            ],
            rows: rawRows
        });

        // Column widths
        const widths = [18, 20, 15, 15, 22, 60, 30, 20, 50, 15, 15, 30];
        worksheet.columns.forEach((col, i) => { col.width = widths[i]; });

        // Color-code the Status column (col A = index 1)
        data.forEach((row, i) => {
            const cell = worksheet.getCell(i + 2, 1); // +2: row 1 is header
            const status = computeStatus(row.returnEligible, row.returnDate, row.orderStatus || '');
            if (status === 'Eligible') {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
                cell.font = { color: { argb: 'FF276221' }, bold: true };
            } else if (status === 'Urgent - Return Soon') {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
                cell.font = { color: { argb: 'FF9C5700' }, bold: true };
            } else if (status === 'Window Closed') {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
                cell.font = { color: { argb: 'FF9C0006' }, bold: true };
            } else if (status === 'Returned') {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
                cell.font = { color: { argb: 'FF595959' }, bold: true };
            } else if (status === 'Not Yet Arrived') {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDAE8FC' } };
                cell.font = { color: { argb: 'FF1F4E79' }, bold: true };
            }
        });

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);

        const timestamp = new Date().toISOString().slice(0, 10);

        chrome.downloads.download({
            url: url,
            filename: `amazon_orders_${timestamp}.xlsx`,
            saveAs: true
        });
    }

    function resetButtons() {
        scrapeBtn.disabled = false;
        if (scrapeAllBtn) scrapeAllBtn.disabled = false;
    }

    scrapeAllBtn.addEventListener('click', async () => {
        setStatus('Checking current tab...', 'info');
        scrapeBtn.disabled = true;
        scrapeAllBtn.disabled = true;

        try {
            let [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

            if (!tab.url || !tab.url.includes('amazon.com')) {
                setStatus('Please navigate to Amazon "Your Orders" page first.', 'error');
                resetButtons();
                return;
            }

            setStatus('Starting auto-scrape...', 'info');

            // Initialize storage and reload page to start the content script auto-scrape loop
            chrome.storage.local.set({ isAutoScraping: true, scrapedData: [] }, () => {
                chrome.tabs.reload(tab.id);
                // Status for the popup before it closes
                setStatus('Auto-scraping started. You can close this popup.', 'success');
            });

        } catch (error) {
            console.error('Auto-scrape init error:', error);
            setStatus('Error starting auto-scrape.', 'error');
            resetButtons();
        }
    });

});
