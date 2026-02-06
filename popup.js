// Teams Transcript Downloader - Popup Script

document.addEventListener('DOMContentLoaded', () => {
    const extractBtn = document.getElementById('extractBtn');
    const projectNameInput = document.getElementById('projectName');
    const statusContainer = document.getElementById('statusContainer');
    const statusMessage = document.getElementById('statusMessage');
    const statsContainer = document.getElementById('statsContainer');
    const entryCount = document.getElementById('entryCount');
    const speakerCount = document.getElementById('speakerCount');
    const duration = document.getElementById('duration');

    // Load saved project name
    chrome.storage.local.get(['lastProjectName'], (result) => {
        if (result.lastProjectName) {
            projectNameInput.value = result.lastProjectName;
        }
    });

    extractBtn.addEventListener('click', async () => {
        const projectName = projectNameInput.value.trim() || 'PROYECTO';

        // Save project name for next time
        chrome.storage.local.set({ lastProjectName: projectName });

        // Update UI to loading state
        setStatus('loading', 'Scroll automático en progreso... Capturando todo el contenido.');
        extractBtn.disabled = true;
        extractBtn.innerHTML = '<div class="loading-spinner"></div> Capturando...';

        try {
            // Get the active tab
            const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

            if (!tab) {
                throw new Error('No se pudo acceder a la pestaña activa');
            }

            // Check if we're on a valid page
            if (!isValidUrl(tab.url)) {
                throw new Error('Esta extensión solo funciona en páginas de Microsoft Stream/SharePoint');
            }

            // Send message to content script
            const response = await chrome.tabs.sendMessage(tab.id, { action: 'extractTranscript' });

            if (!response || !response.entries || response.entries.length === 0) {
                throw new Error('No se encontraron entradas de transcripción en esta página');
            }

            // Update stats
            updateStats(response);

            // Generate filename: AAMMDD_PROYECTO_Transcripcion.json
            const filename = generateFilename(projectName);

            // Prepare download data
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

            // Download the file
            downloadJSON(downloadData, filename);

            setStatus('success', `¡Descarga completada! ${response.entries.length} entradas guardadas`);

        } catch (error) {
            console.error('Error extracting transcript:', error);
            setStatus('error', error.message || 'Error al extraer la transcripción');
        } finally {
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
    });

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

        // Sanitize project name
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
