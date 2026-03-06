importScripts('exceljs.min.js');

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "download_xlsx") {
        const data = request.data;
        if (!data || !data.length) return;

        (async () => {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Amazon Orders');

            // Helper to format date strings (e.g., "March 2, 2026") into mm/dd/yyyy
            function parseDateToFormattedString(dateStr) {
                if (!dateStr || dateStr === 'N/A') return dateStr;
                const parsedDate = new Date(dateStr);
                if (isNaN(parsedDate)) return dateStr; // Fallback if parsing fails
                const mm = String(parsedDate.getMonth() + 1).padStart(2, '0');
                const dd = String(parsedDate.getDate()).padStart(2, '0');
                const yyyy = parsedDate.getFullYear();
                return `${mm}/${dd}/${yyyy}`;
            }

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
                name: 'AutoOrdersTable',
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

            // Set column widths
            worksheet.columns.forEach((col, i) => {
                const widths = [15, 15, 22, 60, 50, 15, 15, 30];
                col.width = widths[i];
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
