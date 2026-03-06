importScripts('exceljs.min.js');

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "download_xlsx") {
        const data = request.data;
        if (!data || !data.length) return;

        (async () => {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Amazon Orders');

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

            function computeStatus(returnEligible, returnDateStr) {
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

            const rawRows = data.map(row => [
                computeStatus(row.returnEligible, row.returnDate),
                computeDaysUntilDeadline(row.returnDate),
                parseDateToFormattedString(row.date),
                row.total,
                row.orderId,
                row.items,
                row.asins || '',
                row.link ? { text: row.link, hyperlink: row.link, tooltip: row.link } : '',
                row.returnEligible,
                parseDateToFormattedString(row.returnDate),
                ''
            ]);

            worksheet.addTable({
                name: 'AutoOrdersTable',
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
                    { name: 'Order Link' },
                    { name: 'Return Eligible' },
                    { name: 'Return Date' },
                    { name: 'Notes' }
                ],
                rows: rawRows
            });

            // Column widths
            const widths = [18, 20, 15, 15, 22, 60, 30, 50, 15, 15, 30];
            worksheet.columns.forEach((col, i) => { col.width = widths[i]; });

            // Color-code the Status column (col A = index 1)
            data.forEach((row, i) => {
                const cell = worksheet.getCell(i + 2, 1);
                const status = computeStatus(row.returnEligible, row.returnDate);
                if (status === 'Eligible') {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
                    cell.font = { color: { argb: 'FF276221' }, bold: true };
                } else if (status === 'Urgent - Return Soon') {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
                    cell.font = { color: { argb: 'FF9C5700' }, bold: true };
                } else if (status === 'Window Closed') {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
                    cell.font = { color: { argb: 'FF9C0006' }, bold: true };
                }
            });

            const buffer = await workbook.xlsx.writeBuffer();

            // Convert ArrayBuffer to Base64 in Service Worker context where we can't create Blob URLs
            let binary = '';
            const bytes = new Uint8Array(buffer);
            const len = bytes.byteLength;
            for (let i = 0; i < len; i++) {
                binary += String.fromCharCode(bytes[i]);
            }
            const base64 = btoa(binary);

            const dataUrl = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + base64;
            const timestamp = new Date().toISOString().slice(0, 10);

            chrome.downloads.download({
                url: dataUrl,
                filename: `amazon_orders_auto_${timestamp}.xlsx`,
                saveAs: true
            });
        })();
    }
});
