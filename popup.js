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

    async function downloadXLSX(data) {
        if (!data || !data.length) return;

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Amazon Orders');

        // Define columns and widths
        worksheet.columns = [
            { header: 'Order Date', key: 'date', width: 15 },
            { header: 'Order Total', key: 'total', width: 15 },
            { header: 'Order Number', key: 'orderId', width: 22 },
            { header: 'Items', key: 'items', width: 60 },
            { header: 'Order Link', key: 'link', width: 50 },
            { header: 'Return Eligible', key: 'returnEligible', width: 15 },
            { header: 'Return Date', key: 'returnDate', width: 15 },
            { header: 'Notes', key: 'notes', width: 30 }
        ];

        // Style the header row manually isn't needed if we use a Table, but we need the raw rows
        const rawRows = data.map(row => [
            parseDateToFormattedString(row.date),
            row.total,
            row.orderId,
            row.items,
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
                { name: 'Order Date' },
                { name: 'Order Total' },
                { name: 'Order Number' },
                { name: 'Items' },
                { name: 'Order Link' },
                { name: 'Return Eligible' },
                { name: 'Return Date' },
                { name: 'Notes' }
            ],
            rows: rawRows
        });

        // The addTable method doesn't auto-apply column widths, so we still set them
        worksheet.columns.forEach((col, i) => {
            const widths = [15, 15, 22, 60, 50, 15, 15, 30];
            col.width = widths[i];
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
