/**
 * Main application logic for PackMaillerWEB
 * Ported from VSTO C# implementation.
 */

// Debugging log
console.log("PackMaillerWEB: Script loading version 1.1...");

// Check if settings are available
if (!window.PackSettings) {
    console.error("PackMaillerWEB: settings.js not loaded or PackSettings missing!");
}

let currentSettings = window.PackSettings ? window.PackSettings.get() : { zimmetMode: 'BIRIM' };
let originalFrom = null; // To store the initial sender address
let isProcessed = false; // Flag to skip revert if MAIL HAZIRLA was clicked
const targetAddress = "TTUBBSAWPAKETHAZIRLIK@THY.COM";

// Ultra-defensive Office initialization
function startApp() {
    console.log("PackMaillerWEB: Starting app logic...");
    initApp();
}

// Office context initialization
if (typeof Office !== 'undefined') {
    Office.onReady(function (info) {
        console.log("PackMaillerWEB: Office.js ready check.");
        if (info && info.host) {
            console.log("PackMaillerWEB: Running inside host: " + info.host);
            
            // 1. Store original From address and try to set the target address
            handleInitialFromAddress();

            // 2. Continuous check for From address in Outlook
            checkFromAddress();
            setInterval(checkFromAddress, 5000);

            // 3. Revert on close (best-effort)
            window.addEventListener('beforeunload', revertFromAddress);
        } else {
            console.log("PackMaillerWEB: Running in standalone browser mode.");
        }
        startApp();
    }).catch(function (err) {
        console.error("PackMaillerWEB: Office.onReady failed, starting anyway.", err);
        startApp();
    });
} else {
    console.warn("PackMaillerWEB: Office.js script tag not found or failed to load.");
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', startApp);
    } else {
        startApp();
    }
}

async function handleInitialFromAddress() {
    try {
        const item = Office.context.mailbox.item;
        if (!item || !item.from) return;

        // Get original
        item.from.getAsync(function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                originalFrom = result.value;
                console.log("PackMaillerWEB: Original From saved:", originalFrom.emailAddress);

                // Try to set target
                if (originalFrom.emailAddress.toUpperCase() !== targetAddress) {
                    item.from.setAsync({ emailAddress: targetAddress }, function (setResult) {
                        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.log("PackMaillerWEB: Target From set automatically.");
                        } else {
                            console.warn("PackMaillerWEB: Auto-set From failed:", setResult.error.message);
                        }
                        checkFromAddress(); // Trigger UI update
                    });
                }
            }
        });
    } catch (err) {
        console.warn("PackMaillerWEB: Initial From handling failed", err);
    }
}

function revertFromAddress() {
    try {
        if (!isProcessed && originalFrom && Office.context.mailbox.item && Office.context.mailbox.item.from) {
            // Revert to original if it's different from current
            Office.context.mailbox.item.from.setAsync({ emailAddress: originalFrom.emailAddress });
            console.log("PackMaillerWEB: Reverted From address (cancelled state).");
        }
    } catch (err) {
        // Silently fail as the pane is closing
    }
}

async function checkFromAddress() {
    try {
        if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item) return;

        const item = Office.context.mailbox.item;

        if (item.from && typeof item.from.getAsync === 'function') {
            item.from.getAsync(function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const currentFrom = (result.value.emailAddress || "").toUpperCase();
                    const warningBar = document.getElementById('warningBar');
                    const container = document.getElementById('container');

                    if (warningBar && container) {
                        if (currentFrom !== targetAddress) {
                            warningBar.classList.remove('hidden');
                            container.classList.add('has-warning');
                        } else {
                            warningBar.classList.add('hidden');
                            container.classList.remove('has-warning');
                        }
                    }
                }
            });
        }
    } catch (err) {
        console.warn("PackMaillerWEB: From address check failed", err);
    }
}

function initApp() {
    const btnPrepare = document.getElementById('btnPrepare');
    const btnSettings = document.getElementById('btnSettings');
    const btnCloseSettings = document.getElementById('btnCloseSettings');
    const btnSaveSettings = document.getElementById('btnSaveSettings');
    const settingsOverlay = document.getElementById('settingsOverlay');

    // Email inputs - Add Enter key listener
    ['inBay1', 'inBay2', 'inBay3', 'inCc'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    const keyMap = { 'inBay1': 'bay1To', 'inBay2': 'bay2To', 'inBay3': 'bay3To', 'inCc': 'cc' };
                    addEmailFromUI(keyMap[id], id);
                }
            });
        }
    });

    // Real-time preview updates
    const inputs = ['txtAc', 'txtBakim', 'dateInput'];
    inputs.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', updatePreview);
    });

    document.querySelectorAll('input[type="radio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            if (e.target.name === 'type') handleTypeChange(e.target.value);
            updatePreview();
        });
    });

    document.querySelectorAll('#skillGrid input').forEach(chk => {
        chk.addEventListener('change', updatePreview);
    });

    if (btnPrepare) btnPrepare.onclick = prepareMail;

    if (btnSettings) {
        btnSettings.onclick = () => {
            loadSettingsToUI();
            if (settingsOverlay) settingsOverlay.classList.remove('hidden');
        };
    }

    if (btnCloseSettings) {
        btnCloseSettings.onclick = () => {
            if (settingsOverlay) settingsOverlay.classList.add('hidden');
        };
    }

    if (btnSaveSettings) {
        btnSaveSettings.onclick = () => {
            const newSettings = {
                zimmetMode: document.getElementById('selZimmetMode').value,
                bay1To: currentSettings.bay1To,
                bay2To: currentSettings.bay2To,
                bay3To: currentSettings.bay3To,
                cc: currentSettings.cc
            };
            window.PackSettings.save(newSettings);
            currentSettings = newSettings;
            if (settingsOverlay) settingsOverlay.classList.add('hidden');
            applyZimmetMode();
            updatePreview();
        };
    }

    // Initialize UI
    const today = new Date().toISOString().split('T')[0];
    const dateInput = document.getElementById('dateInput');
    if (dateInput) dateInput.value = today;

    applyZimmetMode();
    updatePreview();
}

function handleTypeChange(type) {
    const isPlanned = type === 'planned';
    const skillChecks = document.querySelectorAll('#skillGrid input');
    if (currentSettings.zimmetMode === 'PLANNER') return;

    skillChecks.forEach(chk => {
        if (isPlanned) {
            chk.checked = true;
            chk.disabled = false;
        } else {
            chk.checked = (chk.id === 'chk30'); // Only mechanics for unplanned
            chk.disabled = true;
        }
    });
}


function applyZimmetMode() {
    const isPlanner = currentSettings.zimmetMode === 'PLANNER';
    const skillGrid = document.getElementById('skillGrid');
    const plannerNote = document.getElementById('plannerNote');

    if (isPlanner) {
        skillGrid.classList.add('hidden');
        plannerNote.classList.remove('hidden');
    } else {
        skillGrid.classList.remove('hidden');
        plannerNote.classList.add('hidden');
        // Reset skills based on current type
        const type = document.querySelector('input[name="type"]:checked').value;
        handleTypeChange(type);
    }
}

function loadSettingsToUI() {
    const s = currentSettings;
    document.getElementById('selZimmetMode').value = s.zimmetMode;
    renderEmailList('bay1To', 'listBay1');
    renderEmailList('bay2To', 'listBay2');
    renderEmailList('bay3To', 'listBay3');
    renderEmailList('cc', 'listCc');
}

function renderEmailList(settingKey, containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;

    container.innerHTML = '';
    const list = currentSettings[settingKey] || [];

    list.forEach((email, index) => {
        const tag = document.createElement('div');
        tag.className = 'email-tag';
        tag.innerHTML = `
            <span>${email}</span>
            <span class="remove-tag" onclick="removeEmail('${settingKey}', ${index}, '${containerId}')">✕</span>
        `;
        container.appendChild(tag);
    });
}

window.addEmailFromUI = function (settingKey, inputId) {
    const input = document.getElementById(inputId);
    const email = input.value.trim().toLowerCase();

    if (!email || !email.includes('@')) {
        alert("Geçersiz e-posta adresi!");
        return;
    }

    if (!currentSettings[settingKey]) currentSettings[settingKey] = [];

    if (currentSettings[settingKey].includes(email)) {
        alert("Bu adres zaten listede!");
        return;
    }

    currentSettings[settingKey].push(email);
    input.value = '';
    renderEmailList(settingKey, inputId.replace('in', 'list'));
};

window.removeEmail = function (settingKey, index, containerId) {
    currentSettings[settingKey].splice(index, 1);
    renderEmailList(settingKey, containerId);
};

function generateHTML() {
    const bay = document.querySelector('input[name="bay"]:checked').value;
    const typeValue = document.querySelector('input[name="type"]:checked').value;
    const ac = document.getElementById('txtAc').value.toUpperCase();
    const bakim = document.getElementById('txtBakim').value.toUpperCase();
    const dateRaw = document.getElementById('dateInput').value;
    const date = dateRaw ? new Date(dateRaw).toLocaleDateString('tr-TR') : '';

    const bakimPlaniText = typeValue === 'planned' ? 'VAR' : typeValue === 'unplanned' ? 'YOK' : 'PERİYODİK BAKIM';
    const isPlanned = typeValue === 'planned';
    const isPlannerMode = currentSettings.zimmetMode === 'PLANNER';

    let html = `
        <div style="font-family: 'Segoe UI', Tahoma, sans-serif; color: #000; line-height: 1.6;">
            <p>Sayın İlgililer,</p>
            <p>Aşağıda bilgileri bulunan bakım paketi hazır olup <strong>BMPM</strong> ofisinden teslim alınabilir.</p>
            
            <div style="background-color: #f4f5f6; border-left: 4px solid #E2001A; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
                <strong>A/C:</strong> ${ac}<br>
                <strong>BAKIM ADI:</strong> ${bakim}<br>
                <strong>BAKIM GİRİŞ TARİHİ:</strong> ${date}<br>
                <strong>BAKIM PLANI:</strong> ${bakimPlaniText}
            </div>
    `;

    if (isPlannerMode) {
        html += `
            <p>Kartların zimmetleneceği <strong>planner</strong> ismini bu e-posta yoluyla bildirmenizi rica ederiz.</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
                <tr>
                    <th style="border: 1px solid #d1d5db; padding: 12px; background-color: #E2001A; color: #fff; text-align: left; width: 1%; white-space: nowrap;">PLANNER</th>
                    <td style="border: 1px solid #d1d5db; padding: 12px;"></td>
                </tr>
            </table>
        `;
    } else if (isPlanned) {
        html += `
            <p>Alt tabloda belirtilen birimlere ait kartların kimlere zimmetleneceğini bu e-posta üzerinden bildirmenizi rica ederiz.</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
                <thead>
                    <tr style="background-color: #E2001A; color: #fff;">
                        <th style="border: 1px solid #d1d5db; padding: 12px; text-align: left;">BİRİM</th>
                        <th style="border: 1px solid #d1d5db; padding: 12px; text-align: left;">ZİMMET ÇIKILACAK İSİM</th>
                    </tr>
                </thead>
                <tbody>
        `;

        const skills = [
            { id: 'chk30', name: 'MEKANİK' },
            { id: 'chk35', name: 'KABİN İÇİ' },
            { id: 'chk40', name: 'AVİYONİK' },
            { id: 'chk51', name: 'YAPISAL' },
            { id: 'chk52', name: 'BOYA' },
            { id: 'chk53', name: 'KOLTUK' },
            { id: 'chk98', name: 'BOROSKOP' },
            { id: 'chk99', name: 'NDT' }
        ];

        skills.forEach(skill => {
            if (document.getElementById(skill.id).checked) {
                html += `<tr><td style="border: 1px solid #d1d5db; padding: 12px; font-weight: bold;">${skill.name}</td><td style="border: 1px solid #d1d5db; padding: 12px;"></td></tr>`;
            }
        });

        html += `</tbody></table>`;
    } else {
        html += `
            <p>Bakım planı bulunmadığından, paket içeriğinin tamamı tek bir isim üzerine zimmetlenecektir. Kartların kime zimmetleneceğini tabloya işleyerek bu e-posta üzerinden bildirmenizi rica ederiz.</p>
            <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
                <thead>
                    <tr style="background-color: #E2001A; color: #fff;">
                        <th style="border: 1px solid #d1d5db; padding: 12px; text-align: left;">BİRİM</th>
                        <th style="border: 1px solid #d1d5db; padding: 12px; text-align: left;">ZİMMET ÇIKILACAK İSİM</th>
                    </tr>
                </thead>
                <tbody>
                    <tr><td style="border: 1px solid #d1d5db; padding: 12px; font-weight: bold;">MEKANİK</td><td style="border: 1px solid #d1d5db; padding: 12px;"></td></tr>
                </tbody>
            </table>
        `;
    }

    html += `
            <div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #d1d5db; font-size: 0.9em; color: #6c757d;">
                <strong>BAKIM PLANLAMA ŞEFLİĞİ (SAW)</strong><br>
                BAKIM HAZIRLIK BİRİMİ
            </div>
        </div>
    `;

    return html;
}

function updatePreview() {
    try {
        const preview = document.getElementById('mailPreview');
        if (preview) {
            preview.innerHTML = generateHTML();
        }
    } catch (err) {
        console.error("PackMaillerWEB: Önizleme güncellenemedi", err);
    }
}

async function prepareMail() {
    isProcessed = true; // Mark as processed so we don't revert on close
    const ac = document.getElementById('txtAc').value.toUpperCase().trim();
    const bakim = document.getElementById('txtBakim').value.toUpperCase().trim();

    if (!ac || !bakim) {
        alert("A/C ve BAKIM ADI alanları boş bırakılamaz!");
        return;
    }

    const bay = document.querySelector('input[name="bay"]:checked').value;
    const subject = `${ac} / ${bakim} BAKIM PAKETİ HK.`;
    const body = generateHTML();

    const toRecipients = bay === 'BAY-1' ? currentSettings.bay1To : bay === 'BAY-2' ? currentSettings.bay2To : currentSettings.bay3To;
    const ccRecipients = currentSettings.cc;

    const toStr = Array.isArray(toRecipients) ? toRecipients.join(', ') : toRecipients;
    const ccStr = Array.isArray(ccRecipients) ? ccRecipients.join(', ') : ccRecipients;

    try {
        // Office.js Calls to set subject, body, to, cc
        // Using Office.context.mailbox.item
        const item = Office.context.mailbox.item;

        if (!item) {
            console.warn("Item not found (might be in browser simulation)");
            return;
        }

        // Auto-set FROM address
        if (item.from && typeof item.from.setAsync === 'function') {
            item.from.setAsync({ emailAddress: targetAddress }, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    alert("UYARI: Gönderen (Kimden) hesabı 'TTUBBSAWPAKETHAZIRLIK@THY.COM' olarak ayarlanamadı. Lütfen yetkilerinizi kontrol edip maili manuel gönderiniz.\n\nHata: " + asyncResult.error.message);
                    console.log("From (Kimden) hesabı otomatik ayarlanamadı: " + asyncResult.error.message);
                    return; // Stop execution if we can't set the from address
                } else {
                    finalizeMailPreparation(item, subject, toRecipients, ccRecipients, body);
                }
            });
        } else {
             // Fallback if from API is totally unavailable, but still prepare the rest
             finalizeMailPreparation(item, subject, toRecipients, ccRecipients, body);
        }

    } catch (error) {
        console.error("Preparation failed", error);
        alert("Mail hazırlanırken beklenmeyen bir hata oluştu.");
    }
}

function finalizeMailPreparation(item, subject, toRecipients, ccRecipients, body) {
    item.subject.setAsync(subject);
    item.to.setAsync(toRecipients); // Office.js accepts arrays
    item.cc.setAsync(ccRecipients);
    item.body.setAsync(body, { coercionType: Office.CoercionType.Html });
    console.log("Mail prepared successfully!");
}
