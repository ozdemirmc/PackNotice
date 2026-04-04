/**
 * Settings management for PackNotice
 * Uses localStorage to persist user configurations.
 */

const SETTINGS_KEY = 'PackNoticeSettings';

const DefaultSettings = {
    zimmetMode: 'BIRIM', // 'BIRIM' or 'PLANNER'
    bay1To: ['tt-ubbsaw-bay1planner@thy.com', 'ttubbsawbay1yoneticipersonel@thy.com'],
    bay2To: ['ttubbsawbay2planner@thy.com', 'ttubbsawbay2yoneticipersonel@thy.com'],
    bay3To: ['ttubbsawbay3planner@thy.com', 'ttubbsawbay3yoneticipersonel@thy.com'],
    cc: ['ttubbsawpakethazirlik@thy.com', 'ttubbsawbakimplanlama@thy.com']
};

function loadSettings() {
    const saved = localStorage.getItem(SETTINGS_KEY);
    if (saved) {
        try {
            const parsed = JSON.parse(saved);
            // Backward compatibility: migrate strings to arrays if needed
            ['bay1To', 'bay2To', 'bay3To', 'cc'].forEach(key => {
                if (typeof parsed[key] === 'string') {
                    parsed[key] = parsed[key].split(',').map(e => e.trim()).filter(e => e);
                }
            });
            return { ...DefaultSettings, ...parsed };
        } catch (e) {
            console.error('Failed to parse settings', e);
        }
    }
    return DefaultSettings;
}

function saveSettings(settings) {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
}

// Global accessor
window.PackSettings = {
    get: loadSettings,
    save: saveSettings,
    reset: () => {
        localStorage.removeItem(SETTINGS_KEY);
        return DefaultSettings;
    }
};
