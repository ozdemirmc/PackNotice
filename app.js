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
            // Continuous check for From address in Outlook
            checkFromAddress();
            setInterval(checkFromAddress, 5000);
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

async function checkFromAddress() {
    try {
        if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item) return;

        const item = Office.context.mailbox.item;
        const targetAddress = "TTUBBSAWPAKETHAZIRLIK@THY.COM";

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
    // UI Elements
    const btnPrepare = document.getElementById('btnPrepare');
    const btnSettings = document.getElementById('btnSettings');
    const btnCloseSettings = document.getElementById('btnCloseSettings');
    const btnSaveSettings = document.getElementById('btnSaveSettings');
    const settingsOverlay = document.getElementById('settingsOverlay');

    const inputs = ['txtAc', 'txtBakim', 'dateInput', 'skillGrid'];
    const radioGroups = ['bay', 'type'];

    // Event listeners for real-time preview
    inputs.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', updatePreview);
    });

    document.querySelectorAll('input[type="radio"]').forEach(radio => {
        radio.addEventListener('change', (e) => {
            if (e.target.name === 'type') {
                handleTypeChange(e.target.value);
            }
            updatePreview();
        });
    });

    document.querySelectorAll('#skillGrid input').forEach(chk => {
        chk.addEventListener('change', updatePreview);
    });

    // Main Actions
    if (btnPrepare) {
        btnPrepare.onclick = function () {
            console.log("PackMaillerWEB: Hazırla butonu tıklandı.");
            prepareMail();
        };
    }

    // Settings Modal Actions
    if (btnSettings) {
        btnSettings.onclick = function () {
            console.log("PackMaillerWEB: Ayarlar butonu tıklandı.");
            loadSettingsToUI();
            if (settingsOverlay) settingsOverlay.classList.remove('hidden');
        };
    }

    if (btnCloseSettings) {
        btnCloseSettings.onclick = function () {
            if (settingsOverlay) settingsOverlay.classList.add('hidden');
        };
    }

    if (btnSaveSettings) {
        btnSaveSettings.onclick = function () {
            console.log("PackMaillerWEB: Ayarlar kaydediliyor.");
            const newSettings = {
                zimmetMode: document.getElementById('selZimmetMode').value,
                bay1To: document.getElementById('txtToBay1').value,
                bay2To: document.getElementById('txtToBay2').value,
                bay3To: document.getElementById('txtToBay3').value,
                cc: document.getElementById('txtCc').value
            };
            window.PackSettings.save(newSettings);
            currentSettings = newSettings;
            if (settingsOverlay) settingsOverlay.classList.add('hidden');
            applyZimmetMode();
            updatePreview();
        };
    }

    // Initialize UI
    try {
        const today = new Date().toISOString().split('T')[0];
        const dateInput = document.getElementById('dateInput');
        if (dateInput) dateInput.value = today;

        applyZimmetMode();
        updatePreview();
        console.log("PackMaillerWEB: UI Hazır.");
    } catch (err) {
        console.error("PackMaillerWEB: UI başlatılırken hata oluştu", err);
    }
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
            // Unplanned or Periodic
            if (chk.id === 'chk30') { // 30-MEKANIK
                chk.checked = true;
                chk.disabled = true;
            } else {
                chk.checked = false;
                chk.disabled = true;
            }
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

        item.subject.setAsync(subject);
        item.to.setAsync(toRecipients); // Office.js accepts arrays
        item.cc.setAsync(ccRecipients);
        item.body.setAsync(body, { coercionType: Office.CoercionType.Html });

        console.log("Mail prepared successfully!");
    } catch (error) {
        console.error("Preparation failed", error);
    }
}
