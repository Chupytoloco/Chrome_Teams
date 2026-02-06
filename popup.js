// Teams Transcript Downloader - Popup Script

document.addEventListener('DOMContentLoaded', () => {
    const extractBtn = document.getElementById('extractBtn');
    const stopBtn = document.getElementById('stopBtn');
    const statusContainer = document.getElementById('statusContainer');
    const statusMessage = document.getElementById('statusMessage');
    const statsContainer = document.getElementById('statsContainer');
    const entryCount = document.getElementById('entryCount');
    const speakerCount = document.getElementById('speakerCount');
    const duration = document.getElementById('duration');

    let processing = false;

    extractBtn.addEventListener('click', async () => {
        if (processing) return;

        startProcessing();

        try {
            const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

            if (!tab) throw new Error('No se pudo acceder a la pestaña activa');
            if (!isValidUrl(tab.url)) throw new Error('Esta extensión solo funciona en páginas de Microsoft Stream/SharePoint');

            // Send extract message
            const response = await chrome.tabs.sendMessage(tab.id, { action: 'extractTranscript' });

            handleResponse(response);

        } catch (error) {
            console.error('Error:', error);
            setStatus('error', error.message || 'Error al extraer la transcripción');
            resetUI();
        }
    });

    stopBtn.addEventListener('click', async () => {
        try {
            const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
            if (tab) {
                // Send stop message
                await chrome.tabs.sendMessage(tab.id, { action: 'stopExtraction' });
                setStatus('loading', 'Deteniendo... procesando datos capturados.');
            }
        } catch (error) {
            console.error('Error stopping:', error);
        }
    });

    function startProcessing() {
        processing = true;
        setStatus('loading', 'Scroll automático en progreso... Capturando todo el contenido.');
        extractBtn.style.display = 'none';
        stopBtn.style.display = 'flex';
        extractBtn.disabled = true;
    }

    function resetUI() {
        processing = false;
        extractBtn.style.display = 'flex';
        stopBtn.style.display = 'none';
        extractBtn.disabled = false;
        extractBtn.innerHTML = `
            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="7,10 12,15 17,10"></polyline>
                <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            Descargar Transcripción
        `;
    }

    function handleResponse(response) {
        if (!response || !response.entries || response.entries.length === 0) {
            if (response && response.error) {
                throw new Error(response.error);
            }
            throw new Error('No se encontraron entradas de transcripción');
        }

        updateStats(response);

        // Use meeting title or default
        const meetingTitle = response.title || 'Transcripcion_Sin_Titulo';
        const filename = generateFilename(meetingTitle);

        const downloadData = {
            metadata: {
                title: response.title,
                exportDate: new Date().toISOString(),
                source: response.source,
                meetingDate: response.meetingDate,
                totalEntries: response.entries.length,
                speakers: response.speakers,
                duration: response.duration
            },
            transcript: response.entries
        };

        downloadJSON(downloadData, filename);

        setStatus('success', `¡Descarga completada! ${response.entries.length} entradas guardadas`);
        resetUI();
    }

    function setStatus(type, message) {
        statusContainer.className = 'status-container ' + type;
        statusMessage.textContent = message;
    }

    function updateStats(response) {
        statsContainer.style.display = 'grid';
        entryCount.textContent = response.entries.length;
        speakerCount.textContent = response.speakers?.length || 0;
        duration.textContent = response.duration || '--:--';
    }

    function isValidUrl(url) {
        if (!url) return false;
        return url.includes('sharepoint.com') ||
            url.includes('microsoft.com') ||
            url.includes('teams.microsoft.com');
    }

    function generateFilename(title) {
        const now = new Date();
        const year = String(now.getFullYear()).slice(-2);
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');

        // Sanitize title: remove extension if any (.mp4), remove invalid chars
        let safeTitle = title.replace(/\.[^/.]+$/, ""); // Remove extension
        safeTitle = safeTitle
            .replace(/Transcripci[oó]n/gi, '') // Remove redundant "transcripcion"
            .replace(/Grabaci[oó]n/gi, '') // Remove redundant "grabacion"
            .trim();

        safeTitle = safeTitle
            .replace(/[^a-zA-Z0-9\u00C0-\u00FF \-_]/g, '') // Keep spaces for now
            .trim()
            .replace(/\s+/g, '_') // Space to underscore
            .toUpperCase();

        if (safeTitle.length > 50) safeTitle = safeTitle.substring(0, 50);
        if (!safeTitle) safeTitle = "REUNION";

        return `${year}${month}${day}_${safeTitle}_Transcripcion.json`;
    }

    function downloadJSON(data, filename) {
        const jsonStr = JSON.stringify(data, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
});
