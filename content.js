// Content script for Teams Transcript Downloader
// SIMPLIFIED v2: Better parsing, single scroll, stoppable

(function () {
    'use strict';

    let isExtracting = false;
    let shouldStop = false;

    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
        if (request.action === 'extractTranscript') {
            shouldStop = false;
            extractWithSingleScroll()
                .then(transcript => sendResponse(transcript))
                .catch(error => {
                    console.error('Extraction error:', error);
                    sendResponse({ error: error.message, entries: [] });
                });
            return true; // Keep channel open for async
        }

        if (request.action === 'stopExtraction') {
            console.log('Stopping extraction...');
            shouldStop = true;
            // Response will be handled by the main promise resolving with partial data
            return true;
        }

        return true;
    });

    function findTranscriptContainer() {
        const selectors = [
            '[data-is-scrollable="true"]',
            '[class*="scrollablePane"]',
            '[class*="transcriptPane"]',
            '.ms-ScrollablePane--contentContainer'
        ];

        for (const selector of selectors) {
            const container = document.querySelector(selector);
            if (container && container.scrollHeight > container.clientHeight) {
                return container;
            }
        }

        const entry = document.querySelector('[class*="itemHeader-"]');
        if (entry) {
            let parent = entry.parentElement;
            for (let i = 0; i < 10 && parent; i++) {
                if (parent.scrollHeight > parent.clientHeight + 100) {
                    return parent;
                }
                parent = parent.parentElement;
            }
        }

        return null;
    }

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    async function extractWithSingleScroll() {
        if (isExtracting) return { error: 'Ya hay una extracción en curso' };
        isExtracting = true;

        const container = findTranscriptContainer();

        if (!container) {
            isExtracting = false;
            console.log('No scrollable container found');
            return extractVisibleOnly();
        }

        console.log('Starting extraction...');

        // Store items by listItem ID
        const capturedItems = new Map();
        let captureOrder = 0;

        // Start at top
        container.scrollTop = 0;
        await sleep(300);

        const scrollStep = 150;
        const scrollDelay = 100;
        let stuckCount = 0;
        let lastScrollTop = -1;

        // Single scroll pass
        while (stuckCount < 3) {
            if (shouldStop) {
                console.log('Extraction stopped by user');
                break;
            }

            // Capture
            captureOrder = captureVisible(capturedItems, captureOrder);

            if (container.scrollTop === lastScrollTop) {
                stuckCount++;
            } else {
                stuckCount = 0;
            }
            lastScrollTop = container.scrollTop;

            if (container.scrollTop + container.clientHeight >= container.scrollHeight - 5) {
                captureOrder = captureVisible(capturedItems, captureOrder);
                break;
            }

            container.scrollTop += scrollStep;
            await sleep(scrollDelay);
        }

        console.log(`Captured ${capturedItems.size} items`);
        isExtracting = false;

        // Even if stopped, return what we found so far
        return buildResult(capturedItems);
    }

    function captureVisible(capturedItems, startOrder) {
        let order = startOrder;
        const listItems = document.querySelectorAll('[id^="listItem-"]');

        listItems.forEach(item => {
            const itemId = item.id;
            if (capturedItems.has(itemId)) return;

            // Get header element and parse it properly
            const headerEl = item.querySelector('[class*="itemHeader-"], [class*="itemHeader_"]');
            const textEl = item.querySelector('[class*="entryText-"], [class*="entryText_"]');

            let speaker = null;
            let timestamp = null;

            if (headerEl) {
                // Get child elements separately for better parsing
                const nameEl = headerEl.querySelector('[class*="speakerName"], [class*="displayName"]');
                const timeEl = headerEl.querySelector('[class*="timestamp"], [class*="time"]');

                if (nameEl) {
                    speaker = nameEl.textContent.trim();
                }
                if (timeEl) {
                    const timeText = timeEl.textContent.trim();
                    const match = timeText.match(/\d{1,2}:\d{2}/);
                    if (match) timestamp = match[0];
                }

                // If no specific elements found, parse the raw text
                if (!speaker || !timestamp) {
                    const parsed = parseHeaderText(headerEl.textContent);
                    if (!speaker) speaker = parsed.speaker;
                    if (!timestamp) timestamp = parsed.timestamp;
                }
            }

            const entryText = textEl ? textEl.textContent.trim() : null;

            if (speaker || entryText) {
                capturedItems.set(itemId, {
                    id: itemId,
                    speaker: speaker,
                    timestamp: timestamp,
                    text: entryText,
                    order: order++
                });
            }
        });

        return order;
    }

    // Parse header text manually
    function parseHeaderText(rawText) {
        const result = { speaker: null, timestamp: null };
        if (!rawText) return result;

        const text = rawText.trim();

        // TIMESTAMP: It's at the very END of the text, format "X:XX" (4-5 chars)
        const lastChars = text.slice(-5);
        const endMatch = lastChars.match(/(\d{1,2}:\d{2})$/);
        if (endMatch) {
            result.timestamp = endMatch[1];
        } else {
            // Fallback: try last 4 characters
            const last4 = text.slice(-4);
            const match4 = last4.match(/(\d:\d{2})$/);
            if (match4) {
                result.timestamp = match4[1];
            }
        }

        // SPEAKER: Get the text that contains "|"
        const pipeIndex = text.indexOf('|');
        if (pipeIndex > 0) {
            const afterPipe = text.substring(pipeIndex + 1);
            const companyMatch = afterPipe.match(/^([^0-9]+)/);
            const company = companyMatch ? companyMatch[1].trim() : '';
            const beforePipe = text.substring(0, pipeIndex).trim();
            result.speaker = beforePipe + ' | ' + company;
        } else {
            const nameMatch = text.match(/^([^0-9]+)/);
            if (nameMatch) {
                result.speaker = nameMatch[1].trim();
            }
        }

        if (result.speaker) {
            result.speaker = result.speaker.replace(/\s*\|\s*$/, '').trim();
            result.speaker = result.speaker.replace(/\s+/g, ' ').trim();
        }

        return result;
    }

    function buildResult(capturedItems) {
        const result = {
            entries: [],
            speakers: [],
            duration: null,
            source: window.location.href,
            title: document.title,
            meetingDate: null
        };

        const dateElement = document.querySelector('[class*="subTitleBar"]');
        if (dateElement) {
            const dateMatch = dateElement.textContent.match(/(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})/i);
            if (dateMatch) {
                result.meetingDate = dateMatch[1];
            }
        }

        const sortedItems = Array.from(capturedItems.values()).sort((a, b) => a.order - b.order);

        const speakersSet = new Set();
        let currentSpeaker = null;
        let currentTimestamp = null;
        let currentTexts = [];

        sortedItems.forEach((item) => {
            if (item.speaker) {
                if (currentTexts.length > 0) {
                    result.entries.push({
                        index: result.entries.length + 1,
                        timestamp: currentTimestamp,
                        speaker: currentSpeaker,
                        text: currentTexts.join(' ')
                    });
                    currentTexts = [];
                }

                currentSpeaker = item.speaker;
                currentTimestamp = item.timestamp;

                speakersSet.add(currentSpeaker);
                if (currentTimestamp) result.duration = currentTimestamp;
            }

            if (item.text && item.text.length > 0) {
                currentTexts.push(item.text);
            }
        });

        if (currentTexts.length > 0) {
            result.entries.push({
                index: result.entries.length + 1,
                timestamp: currentTimestamp,
                speaker: currentSpeaker,
                text: currentTexts.join(' ')
            });
        }

        result.speakers = Array.from(speakersSet);
        return result;
    }

    function extractVisibleOnly() {
        const result = {
            entries: [],
            speakers: [],
            duration: null,
            source: window.location.href,
            title: document.title,
            meetingDate: null
        };

        const listItems = document.querySelectorAll('[id^="listItem-"]');
        const speakersSet = new Set();

        listItems.forEach((item) => {
            const header = item.querySelector('[class*="itemHeader-"]');
            const textEl = item.querySelector('[class*="entryText-"]');

            const parsed = header ? parseHeaderText(header.textContent) : { speaker: null, timestamp: null };
            const text = textEl ? textEl.textContent.trim() : '';

            if (text) {
                result.entries.push({
                    index: result.entries.length + 1,
                    timestamp: parsed.timestamp,
                    speaker: parsed.speaker,
                    text: text
                });
                if (parsed.speaker) speakersSet.add(parsed.speaker);
                if (parsed.timestamp) result.duration = parsed.timestamp;
            }
        });

        result.speakers = Array.from(speakersSet);
        return result;
    }

    console.log('Teams Transcript Downloader v2.1 loaded');
})();
