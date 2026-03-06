chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "scrape_orders") {
        try {
            const orders = scrapeOrders();
            sendResponse({ success: true, data: orders });
        } catch (error) {
            console.error("Scraping error:", error);
            sendResponse({ success: false, error: error.message });
        }
    }
    return true; // Indicates async response
});

function scrapeOrders() {
    const scrapedData = [];
    // Various Amazon order card selectors
    const orderCards = document.querySelectorAll('.js-order-card, .order-card, .order, .yohtmlc-order');

    if (!orderCards || orderCards.length === 0) {
        console.warn('Amazon Scraper: No order cards found on the page.');
        return [];
    }

    orderCards.forEach(card => {
        try {
            const orderInfo = { date: 'N/A', total: 'N/A', orderId: 'N/A', items: '', asins: '', link: '', returnEligible: 'No', returnDate: 'N/A', orderStatus: '' };

            // 1. Order ID
            const orderIdMatch = card.innerHTML.match(/\d{3}-\d{7}-\d{7}/);
            if (orderIdMatch) orderInfo.orderId = orderIdMatch[0];

            // 2. Order Date & Total

            // Strategy A: Look at list items for 'Order placed' and 'Total' (Newer layout)
            const listItems = card.querySelectorAll('.order-header__header-list-item, div.a-column');
            listItems.forEach(item => {
                const text = item.textContent.replace(/\s+/g, ' ').trim(); // Normalize spaces

                // Date extraction
                if (text.includes('Order placed')) {
                    // Extract if text looks like "Order placed March 2, 2026"
                    const dateMatch = text.match(/Order placed\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})/i);
                    if (dateMatch) {
                        orderInfo.date = dateMatch[1].trim();
                    } else {
                        // Fallback: look for a specific span that looks like a date
                        const spans = item.querySelectorAll('span');
                        spans.forEach(s => {
                            if (s.textContent.trim().match(/^[A-Za-z]+\s+\d{1,2},\s+\d{4}$/)) {
                                orderInfo.date = s.textContent.trim();
                            }
                        });
                    }
                }

                // Total extraction
                if (text.includes('Total') && typeof orderInfo.total === 'string' && orderInfo.total === 'N/A') {
                    const valMatch = text.match(/Total\s*(?:\n)?\s*(\$[\d,]+\.\d{2})/i);
                    if (valMatch) {
                        orderInfo.total = valMatch[1].trim();
                    } else {
                        const spans = item.querySelectorAll('span');
                        spans.forEach(s => {
                            if (s.textContent.trim().match(/^\$[\d,]+\.\d{2}$/)) {
                                orderInfo.total = s.textContent.trim();
                            }
                        });
                    }
                }
            });

            // Strategy B: Old layout (.order-info .a-color-secondary.value)
            if (orderInfo.date === 'N/A' || orderInfo.total === 'N/A') {
                const infoValues = card.querySelectorAll('.order-info .a-color-secondary.value, .yohtmlc-order-header .a-color-secondary.value');
                if (infoValues.length >= 2) {
                    if (orderInfo.date === 'N/A') orderInfo.date = infoValues[0].textContent.trim();
                    if (orderInfo.total === 'N/A') orderInfo.total = infoValues[1].textContent.trim();
                } else {
                    // Fallback to simpler searches
                    if (orderInfo.date === 'N/A') {
                        const dateEl = card.querySelector('.order-date-invoice-item .a-color-secondary, .yohtmlc-order-date');
                        if (dateEl) orderInfo.date = dateEl.textContent.trim();
                    }
                    if (orderInfo.total === 'N/A') {
                        const totalEl = card.querySelector('.yo-ac-order-total');
                        if (totalEl) {
                            orderInfo.total = totalEl.textContent.trim();
                        } else {
                            const priceEls = Array.from(card.querySelectorAll('.a-color-price, .a-size-base.a-color-base'));
                            const priceMatch = priceEls.find(el => el.textContent.trim().match(/^\$?[\d,]+\.\d{2}$/));
                            if (priceMatch) orderInfo.total = priceMatch.textContent.trim();
                        }
                    }
                }
            }

            // 3. Items + ASINs
            const items = [];
            const asins = [];
            const itemLinks = card.querySelectorAll('a.a-link-normal[href*="/product/"], a.a-link-normal[href*="/dp/"]');
            const seenItems = new Set();
            itemLinks.forEach(link => {
                const title = link.textContent.trim();
                // Filter out empty strings or brief linking text, we want actual product titles
                if (title && title.length > 5 && !seenItems.has(title)) {
                    seenItems.add(title);
                    items.push(title);
                    const asinMatch = link.href.match(/\/(?:dp|product)\/([A-Z0-9]{10})/i);
                    if (asinMatch) asins.push(asinMatch[1].toUpperCase());
                }
            });
            orderInfo.items = items.join(' | ');
            orderInfo.asins = asins.join(' | ');

            // 4. Link
            const detailsLink = card.querySelector('a[href*="order-details"]');
            if (detailsLink) {
                orderInfo.link = detailsLink.href;
            } else if (orderInfo.orderId !== 'N/A') {
                orderInfo.link = `https://www.amazon.com/gp/your-account/order-details?orderID=${orderInfo.orderId}`;
            }

            // 5. Order Status (Refunded, Not Yet Arrived)
            const refundEl = card.querySelector('h4.od-status-message span.a-text-bold, h4.a-color-base.od-status-message span.a-text-bold');
            if (refundEl && refundEl.textContent.trim().toLowerCase().includes('refund')) {
                orderInfo.orderStatus = 'Refunded';
            }

            if (!orderInfo.orderStatus) {
                // Once delivered, Amazon shows a return window or "Return window closed".
                // If neither exists, the item hasn't arrived yet.
                // Also check for in-transit delivery language as a safeguard.
                const cardText = card.textContent.toLowerCase();
                const isDelivered = cardText.includes('delivered') || cardText.includes('arriving');
                const hasReturnInfo = orderInfo.returnDate !== 'N/A';
                if (!isDelivered && !hasReturnInfo) {
                    orderInfo.orderStatus = 'Not Yet Arrived';
                }
            }

            // 6. Return Eligibility
            const returnSpans = card.querySelectorAll('span.a-size-small');
            returnSpans.forEach(span => {
                const text = span.textContent.trim();
                // Check if eligible
                const eligibleMatch = text.match(/Return\s+(?:or\s+replace\s+items:\s+)?Eligible\s+through\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})/i);
                if (eligibleMatch) {
                    orderInfo.returnEligible = 'Yes';
                    orderInfo.returnDate = eligibleMatch[1].trim();
                }

                // Check if closed
                const closedMatch = text.match(/Return\s+window\s+closed\s+on\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})/i);
                if (closedMatch) {
                    orderInfo.returnEligible = 'No';
                    orderInfo.returnDate = closedMatch[1].trim();
                }
            });

            scrapedData.push(orderInfo);
        } catch (e) {
            console.error("Amazon Scraper: Error parsing an order card", e);
        }
    });

    return scrapedData;
}

// Auto-scrape logic on page load
// Using a short timeout to ensure dynamic content has settled
setTimeout(() => {
    chrome.storage.local.get(['isAutoScraping', 'scrapedData'], (result) => {
        if (result.isAutoScraping) {
            console.log("Amazon Scraper: Auto-scraping this page...");
            const data = scrapeOrders();
            const allData = (result.scrapedData || []).concat(data);

            let foundNextHref = null;
            // Look for links containing "startIndex="
            const nextLinks = Array.from(document.querySelectorAll('a')).filter(a =>
                a.href &&
                a.href.includes('startIndex=') &&
                (a.textContent.includes('Next') || a.innerHTML.includes('→'))
            );

            // Avoid disabled Next buttons
            if (nextLinks.length > 0 && !nextLinks[0].closest('.a-disabled')) {
                foundNextHref = nextLinks[0].href;
            }

            if (foundNextHref) {
                console.log("Amazon Scraper: Found next page, navigating...", foundNextHref);
                chrome.storage.local.set({ scrapedData: allData }, () => {
                    window.location.href = foundNextHref;
                });
            } else {
                console.log("Amazon Scraper: No next page found. Auto-scrape complete.");
                chrome.storage.local.set({ isAutoScraping: false, scrapedData: [] }, () => {
                    // Send data to background script for download
                    chrome.runtime.sendMessage({ action: "download_xlsx", data: allData });
                    alert("Amazon Scraper: Finished gathering all pages! Your XLSX should be downloading now.");
                });
            }
        }
    });
}, 2500); // 2.5 second delay to ensure Amazon's JS finishes rendering the full list
