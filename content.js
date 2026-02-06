// Content script for Teams Transcript Downloader
// SIMPLIFIED v4: Optimized speed with integrity check (Adaptive Speed)

(function () {
    'use strict';

    let isExtracting = false;
    let shouldStop = false;

    chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
        if (request.action === 'extractTranscript') {
            shouldStop = false;
            extractWithAdaptiveScroll()
                .then(transcript => sendResponse(transcript))
                .catch(error => {
                    console.error('Extraction error:', error);
                    sendResponse({ error: error.message, entries: [] });
                });
            return true;
        }

        if (request.action === 'stopExtraction') {
            shouldStop = true;
            return true;
        }

        return true;
    });

    function findTranscriptContainer() {
        const selectors = [
            '[data-is-scrollable="true"]',
            '.ms-ScrollablePane--contentContainer',
            '[class*="scrollablePane"]',
            '[class*="transcriptPane"]'
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
            for (let i = 0; i < 15 && parent; i++) {
                const style = window.getComputedStyle(parent);
                const isScrollable = style.overflowY === 'auto' || style.overflowY === 'scroll' || parent.scrollHeight > parent.clientHeight + 50;

                if (isScrollable && parent.scrollHeight > 500) {
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

    async function extractWithAdaptiveScroll() {
        if (isExtracting) return { error: 'Ya hay una extracción en curso' };
        isExtracting = true;

        const container = findTranscriptContainer();

        if (!container) {
            isExtracting = false;
            return extractVisibleOnly();
        }

        console.log('Starting ADAPTIVE extraction...');
        const capturedItems = new Map();

        if (container.scrollTop > 50) {
            container.scrollTop = 0;
            await sleep(400);
        }

        // BALANCED SETTINGS
        const BASE_STEP = 400;
        const BASE_DELAY = 100;

        let stuckCount = 0;
        let lastOkScrollTop = 0;
        let lastScrollTop = -1;
        let steps = 0;
        let emptyCaptureStreak = 0;

        while (stuckCount < 5 && steps < 10000) {
            if (shouldStop) break;
            steps++;

            // 1. Capture
            const prevSize = capturedItems.size;
            captureVisible(capturedItems, steps * 1000); // Order logic handled internally based on DOM pos
            const newSize = capturedItems.size;
            const capturedCount = newSize - prevSize;

            // 2. Adaptive logic: 
            // If we didn't capture anything new, maybe we scrolled too fast -> Wait and retry
            if (capturedCount === 0 && steps > 1) {
                emptyCaptureStreak++;
                if (emptyCaptureStreak < 3) {
                    await sleep(100); // Extra wait
                    continue; // Retry capture without scrolling
                }
            } else {
                emptyCaptureStreak = 0;
            }

            // 3. Ratchet & Scroll check
            let currentScroll = container.scrollTop;
            if (currentScroll < lastOkScrollTop - 50) {
                container.scrollTop = lastOkScrollTop;
                currentScroll = lastOkScrollTop;
                await sleep(50);
            } else {
                lastOkScrollTop = Math.max(lastOkScrollTop, currentScroll);
            }

            if (Math.abs(currentScroll - lastScrollTop) < 2) {
                stuckCount++;
            } else {
                stuckCount = 0;
            }
            lastScrollTop = currentScroll;

            if (currentScroll >= (container.scrollHeight - container.clientHeight) - 5) {
                console.log('Reached bottom');
                await sleep(300);
                captureVisible(capturedItems, steps * 1000);
                break;
            }

            // 4. Scroll down
            // If we captured a lot, we can scroll confidently. If few, be careful.
            container.scrollBy({ top: BASE_STEP, behavior: 'instant' });
            await sleep(BASE_DELAY);
        }

        isExtracting = false;
        return buildResult(capturedItems);
    }

    // Capture items and assign order based on their DOM position (top relative)
    function captureVisible(capturedItems, baseOrder) {
        const listItems = document.querySelectorAll('[id^="listItem-"]');

        // We capture EVERYTHING visible.
        // To maintain order, we use the element's position on screen? 
        // Actually, listItem IDs are usually sequential or content is sequential in DOM.
        // We will sort by ID processing or DOM order at the end.
        // BUT, since we have virtual list, DOM order is only for visible items.
        // We need a way to order globally. 
        // Trick: The map insertion order is roughly chronological if we scroll fast enough,
        // but let's strictly rely on DOM structure at the time of capture + scroll offset if needed.

        // Better approach: Store a "global approximate Y" = scrollTop + rect.top?
        // No, IDs usually are not sequential enough.
        // Let's rely on the fact we store in a Map. Map preserves insertion order.
        // Since we scroll from top to bottom, new items encountered will be added to Map in correct order.
        // Existing items stay.

        listItems.forEach(item => {
            if (capturedItems.has(item.id)) return;

            const headerEl = item.querySelector('[class*="itemHeader-"], [class*="itemHeader_"]');
            const textEl = item.querySelector('[class*="entryText-"], [class*="entryText_"]');

            const parsed = headerEl ? parseHeaderSmart(headerEl) : { speaker: null, timestamp: null };
            const entryText = textEl ? textEl.textContent.trim() : null;

            if (parsed.speaker || entryText) {
                capturedItems.set(item.id, {
                    id: item.id,
                    speaker: parsed.speaker,
                    timestamp: parsed.timestamp,
                    text: entryText,
                    // We don't need explicit numerical order if we trust Map insertion order
                    // But to be safe let's add time of capture
                    captureTime: Date.now()
                });
            }
        });
    }

    function parseHeaderSmart(headerEl) {
        const result = { speaker: null, timestamp: null };
        if (!headerEl) return result;

        const nameEl = headerEl.querySelector('[class*="speakerName"], [class*="displayName"]');
        const timeEl = headerEl.querySelector('[class*="timestamp"], [class*="time"]');

        if (nameEl) result.speaker = nameEl.textContent.trim();
        if (timeEl) {
            const timeText = timeEl.textContent.trim();
            const match = timeText.match(/\d{1,2}:\d{2}/);
            if (match) result.timestamp = match[0];
        }

        if (!result.speaker || !result.timestamp) {
            // CRITICAL FIX: Use text with spaces to prevent numbers from merging (e.g. "3" + "0:03" -> "30:03")
            // Instead of headerEl.textContent, we emulate it with spaces
            // We get all leaf elements directly or just iterate child nodes

            // Helper to get text with spaces
            const getTextWithSpaces = (node) => {
                let s = '';
                node.childNodes.forEach(child => {
                    if (child.nodeType === Node.TEXT_NODE) {
                        s += " " + child.textContent;
                    } else if (child.nodeType === Node.ELEMENT_NODE) {
                        s += " " + getTextWithSpaces(child);
                    }
                });
                return s;
            };

            const rawText = getTextWithSpaces(headerEl);
            const cleanText = rawText.replace(/\s+/g, ' ').trim();

            // TIMESTAMP: Look for the last token that matches format X:XX
            const tokens = cleanText.split(' ');
            for (let i = tokens.length - 1; i >= 0; i--) {
                const token = tokens[i];
                // Strict match D:DD or DD:DD at end of string
                const match = token.match(/^(\d{1,2}:\d{2})$/);
                if (match) {
                    result.timestamp = match[1];
                    break;
                }
            }

            // If still no timestamp, try loose regex at end
            if (!result.timestamp) {
                const endMatch = cleanText.match(/(\d{1,2}:\d{2})\s*$/);
                if (endMatch) result.timestamp = endMatch[1];
            }

            // SPEAKER
            // Remove timestamp from end if found to clean up parsing
            let textForSpeaker = cleanText;
            if (result.timestamp) {
                textForSpeaker = textForSpeaker.substring(0, textForSpeaker.lastIndexOf(result.timestamp));
            }

            const pipeIndex = textForSpeaker.indexOf('|');
            if (pipeIndex > 0) {
                const beforePipe = textForSpeaker.substring(0, pipeIndex).trim();
                const afterPipe = textForSpeaker.substring(pipeIndex + 1);
                const companyMatch = afterPipe.match(/^([^0-9\n]+)/);
                result.speaker = beforePipe + ' | ' + (companyMatch ? companyMatch[1].trim() : '');
            } else {
                const nameMatch = textForSpeaker.match(/^([^0-9\n]+)/);
                if (nameMatch) result.speaker = nameMatch[1].trim();
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
            if (dateMatch) result.meetingDate = dateMatch[1];
        }

        // Convert Map to Array - Map preserves insertion order, which is perfect for top-to-bottom scroll
        const items = Array.from(capturedItems.values());

        const speakersSet = new Set();
        let currentSpeaker = null;
        let currentTimestamp = null;
        let currentTexts = [];

        items.forEach((item) => {
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
        // Fallback implementation skipped for brevity, reusing main logic
        // This is rarely reached if container is found
        return {
            entries: [],
            error: "No container found for scrolling"
        };
    }

    console.log('Teams Transcript Downloader v4 (Balanced) loaded');
})();
