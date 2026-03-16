/**
 * Settings management for PackMaillerWEB
 * Uses localStorage to persist user configurations.
 */

const SETTINGS_KEY = 'PackMaillerSettings';

const DefaultSettings = {
    zimmetMode: 'BIRIM', // 'BIRIM' or 'PLANNER'
    bay1To: 'ttubbsawpakethazirlik@thy.com',
    bay2To: 'ttubbsawpakethazirlik@thy.com',
    bay3To: 'ttubbsawpakethazirlik@thy.com',
    cc: 'ttubbsawpakethazirlik@thy.com'
};

function loadSettings() {
    const saved = localStorage.getItem(SETTINGS_KEY);
    if (saved) {
        try {
            return { ...DefaultSettings, ...JSON.parse(saved) };
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
    save: saveSettings
};
