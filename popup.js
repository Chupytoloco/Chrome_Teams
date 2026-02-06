// Teams Transcript Downloader - Popup Script

document.addEventListener('DOMContentLoaded', () => {
    const extractBtn = document.getElementById('extractBtn');
    const stopBtn = document.getElementById('stopBtn');
    const projectNameInput = document.getElementById('projectName');
    const statusContainer = document.getElementById('statusContainer');
    const statusMessage = document.getElementById('statusMessage');
    const statsContainer = document.getElementById('statsContainer');
    const entryCount = document.getElementById('entryCount');
    const speakerCount = document.getElementById('speakerCount');
    const duration = document.getElementById('duration');

    let processing = false;

    // Load saved project name
    chrome.storage.local.get(['lastProjectName'], (result) => {
        if (result.lastProjectName) {
            projectNameInput.value = result.lastProjectName;
        }
    });

    extractBtn.addEventListener('click', async () => {
        if (processing) return;

        const projectName = projectNameInput.value.trim() || 'PROYECTO';
        chrome.storage.local.set({ lastProjectName: projectName });

        startProcessing();

        try {
            const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

            if (!tab) throw new Error('No se pudo acceder a la pestaña activa');
            if (!isValidUrl(tab.url)) throw new Error('Esta extensión solo funciona en páginas de Microsoft Stream/SharePoint');

            // Send extract message
            const response = await chrome.tabs.sendMessage(tab.id, { action: 'extractTranscript' });

            handleResponse(response, projectName);

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
                // Send stop message - the content script should return what it has collected so far
                const response = await chrome.tabs.sendMessage(tab.id, { action: 'stopExtraction' });
                // We'll handle the partial data in the response of the original request, 
                // but just in case the content script returns directly to this call:
                if (response && response.entries) {
                    // Usually the original promise will resolve, so we might not need to do anything here
                    // But setting status is good
                    setStatus('loading', 'Deteniendo... procesando datos capturados.');
                }
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

    function handleResponse(response, projectName) {
        if (!response || !response.entries || response.entries.length === 0) {
            if (response && response.error) {
                throw new Error(response.error);
            }
            throw new Error('No se encontraron entradas de transcripción');
        }

        updateStats(response);

        const filename = generateFilename(projectName);

        const downloadData = {
            metadata: {
                projectName: projectName,
                exportDate: new Date().toISOString(),
                source: response.source,
                title: response.title,
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

    function generateFilename(projectName) {
        const now = new Date();
        const year = String(now.getFullYear()).slice(-2);
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');

        const sanitizedName = projectName
            .toUpperCase()
            .replace(/[^A-Z0-9_]/g, '_')
            .replace(/_+/g, '_')
            .replace(/^_|_$/g, '');

        return `${year}${month}${day}_${sanitizedName}_Transcripcion.json`;
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
