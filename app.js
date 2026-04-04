/**
 * Main application logic for PackNotice
 * Ported from VSTO C# implementation.
 */

// Debugging log
console.log("PackNotice: Script loading version 1.5...");

// Check if settings are available
if (!window.PackSettings) {
    console.error("PackNotice: settings.js not loaded or PackSettings missing!");
}

let currentSettings = window.PackSettings ? window.PackSettings.get() : { zimmetMode: 'BIRIM' };
const targetAddress = "TTUBBSAWPAKETHAZIRLIK@THY.COM";

// Ultra-defensive Office initialization
function startApp() {
    console.log("PackNotice: Starting app logic...");
    initApp();
}

// Office context initialization
if (typeof Office !== 'undefined') {
    Office.onReady(function (info) {
        console.log("PackNotice: Office.js ready check.");
        if (info && info.host) {
            console.log("PackNotice: Running inside host: " + info.host);

            // Gönderici adresini periyodik olarak kontrol et ve uyarı göster
            setTimeout(checkFromAddress, 500);
            setInterval(checkFromAddress, 5000);
        } else {
            console.log("PackNotice: Running in standalone browser mode.");
        }
        startApp();
    }).catch(function (err) {
        console.error("PackNotice: Office.onReady failed, starting anyway.", err);
        startApp();
    });
} else {
    console.warn("PackNotice: Office.js script tag not found or failed to load.");
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', startApp);
    } else {
        startApp();
    }
}

// handleInitialFromAddress ve revertFromAddress kaldırıldı.
// Office.js API'sinde from.setAsync() metodu bulunmadığından
// gönderici adresi programatik olarak değiştirilemez.
// Bunun yerine sadece uyarı gösteriliyor ve MAİL HAZIRLA engelleniyor.

async function checkFromAddress() {
    try {
        if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item) return;

        const item = Office.context.mailbox.item;

        if (item.from && typeof item.from.getAsync === 'function') {
            item.from.getAsync(function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const currentFrom = (result.value.emailAddress || result.value).toUpperCase();
                    const senderNote = document.getElementById('senderNote');

                    if (senderNote) {
                        if (currentFrom !== targetAddress) {
                            senderNote.style.background = '#fff0f0';
                            senderNote.style.borderColor = 'var(--accent)';
                        } else {
                            senderNote.style.background = '#f0fff4';
                            senderNote.style.borderColor = '#22c55e';
                            senderNote.innerHTML = '✅ GÖNDERİCİ HESABI: <b style="color: #22c55e;">TT-UBB(SAW)-BAKIMHAZIRLIK</b> — DOĞRU';
                        }
                    }
                }
            });
        }
    } catch (err) {
        console.warn("PackNotice: From address check failed", err);
    }
}

function initApp() {
    const btnPrepare = document.getElementById('btnPrepare');
    const btnSaveSettings = document.getElementById('btnSaveSettings');
    const alertOverlay = document.getElementById('alertOverlay');
    const btnAlertOk = document.getElementById('btnAlertOk');
    const sidePanel = document.getElementById('sidePanel');
    const sidePanelBackdrop = document.getElementById('sidePanelBackdrop');
    const btnClosePanel = document.getElementById('btnClosePanel');
    const sidePanelTitle = document.getElementById('sidePanelTitle');
    const panelSettings = document.getElementById('panelSettings');
    const panelAbout = document.getElementById('panelAbout');
    const btnResetSettings = document.getElementById('btnResetSettings');
    const confirmOverlay = document.getElementById('confirmOverlay');
    const btnConfirmYes = document.getElementById('btnConfirmYes');
    const btnConfirmNo = document.getElementById('btnConfirmNo');

    // Yan panel aç/kapat fonksiyonları
    window.openSidePanel = function (view) {
        // İçerik değiştir
        if (view === 'settings') {
            sidePanelTitle.textContent = '⚙️ AYARLAR';
            panelSettings.classList.remove('hidden');
            panelAbout.classList.add('hidden');
            loadSettingsToUI();
        } else if (view === 'about') {
            sidePanelTitle.textContent = 'ℹ️ HAKKINDA';
            panelSettings.classList.add('hidden');
            panelAbout.classList.remove('hidden');
        }
        sidePanel.classList.add('open');
        sidePanelBackdrop.classList.remove('hidden');
    };

    window.closeSidePanel = function () {
        sidePanel.classList.remove('open');
        sidePanelBackdrop.classList.add('hidden');
    };

    // Panel kapatma
    if (btnClosePanel) btnClosePanel.onclick = closeSidePanel;
    if (sidePanelBackdrop) sidePanelBackdrop.onclick = closeSidePanel;

    // Uyarı modali kapatma
    if (btnAlertOk) {
        btnAlertOk.onclick = () => {
            if (alertOverlay) alertOverlay.classList.add('hidden');
        };
    }

    // Onay modali kapatma (İptal)
    if (btnConfirmNo) {
        btnConfirmNo.onclick = () => {
            if (confirmOverlay) confirmOverlay.classList.add('hidden');
        };
    }

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

    // AYARLAR linki
    const linkSettings = document.getElementById('linkSettings');
    if (linkSettings) {
        linkSettings.onclick = (e) => {
            e.preventDefault();
            openSidePanel('settings');
        };
    }

    // HAKKINDA linki
    const linkAbout = document.getElementById('linkAbout');
    if (linkAbout) {
        linkAbout.onclick = (e) => {
            e.preventDefault();
            openSidePanel('about');
        };
    }

    if (btnSaveSettings) {
        btnSaveSettings.onclick = () => {
            // BOŞ LİSTE KONTROLÜ
            if (currentSettings.bay1To.length === 0 || currentSettings.bay2To.length === 0 || currentSettings.bay3To.length === 0) {
                showAlert(
                    "<b>TÜM BAY LİSTELERİNDE EN AZ 1 ALICI OLMALIDIR!</b><br><br>" +
                    "LÜTFEN ALICI EKLEYİN VEYA <b style='color: var(--accent);'>VARSAYILANLARA DÖN</b>'Ü KULLANIN.",
                    "⚠️ KAYIT BAŞARISIZ"
                );
                return;
            }

            const newSettings = {
                zimmetMode: document.getElementById('selZimmetMode').value,
                bay1To: currentSettings.bay1To,
                bay2To: currentSettings.bay2To,
                bay3To: currentSettings.bay3To,
                cc: currentSettings.cc
            };
            window.PackSettings.save(newSettings);
            currentSettings = newSettings;
            closeSidePanel();
            applyZimmetMode();
            updatePreview();
            showAlert("AYARLAR BAŞARIYLA KAYDEDİLDİ.", "✅ BAŞARILI");
        };
    }

    if (btnResetSettings) {
        btnResetSettings.onclick = () => {
            showConfirm(
                "ÖZELLEŞTİRDİĞİNİZ GÖNDERİ ADRESLERİ VARSAYILAN AYARLARA DÖNDÜRÜLECEK.<br><br>DEVAM ETMEK İSTİYOR MUSUNUZ?",
                () => {
                    const defaulted = window.PackSettings.reset();
                    currentSettings = defaulted;
                    loadSettingsToUI();
                    applyZimmetMode();
                    updatePreview();
                    if (confirmOverlay) confirmOverlay.classList.add('hidden');
                    console.log("Settings reset to defaults.");
                }
            );
        };
    }

    // Initialize UI
    const dateInput = document.getElementById('dateInput');
    if (dateInput) {
        // Bugünden önceki tarihleri engelle (Türkiye yerel saati)
        const today = new Date();
        today.setMinutes(today.getMinutes() - today.getTimezoneOffset());
        dateInput.min = today.toISOString().split('T')[0];
    }

    applyZimmetMode();
    updatePreview();
}

function handleTypeChange(type) {
    const isPlanned = type === 'planned';

    // ACID alanını periyodik bakımda gizle
    const acidContainer = document.getElementById('acidContainer');
    if (acidContainer) {
        if (type === 'periodic') {
            acidContainer.classList.add('hidden');
            document.getElementById('txtAcid').value = '';
        } else {
            acidContainer.classList.remove('hidden');
        }
    }
    const skillChecks = document.querySelectorAll('#skillGrid input');
    if (currentSettings.zimmetMode === 'PLANNER') return;

    skillChecks.forEach(chk => {
        if (isPlanned) {
            chk.checked = true;
            chk.disabled = false;
        } else {
            // PLANSIZ ve PERİYODİK: sadece MEKANİK
            chk.checked = (chk.id === 'chk30');
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
    let date = '';
    if (dateRaw) {
        // input type="date" her zaman YYYY-MM-DD döner. 
        // new Date(dateRaw) her zaman güvenilir değildir (saat dilimi kayması gibi).
        // Doğrudan string operasyonu daha güvenlidir.
        const parts = dateRaw.split('-');
        if (parts.length === 3) {
            const year = parts[0];
            const month = parts[1];
            const day = parts[2];

            // Gün ismini hesaplama (Güvenli yöntem)
            const days = ['PAZAR', 'PAZARTESİ', 'SALI', 'ÇARŞAMBA', 'PERŞEMBE', 'CUMA', 'CUMARTESİ'];
            const d = new Date(year, month - 1, day);
            const dayName = days[d.getDay()];

            // Bugün mü kontrolü (saat dilimi güvenli)
            const todayObj = new Date();
            todayObj.setMinutes(todayObj.getMinutes() - todayObj.getTimezoneOffset());
            const todayStr = todayObj.toISOString().split('T')[0];

            if (dateRaw === todayStr) {
                date = `${day}.${month}.${year} ${dayName} <span style="color: #E2001A; font-weight: 700;">(BUGÜN)</span>`;
            } else {
                date = `${day}.${month}.${year} ${dayName}`;
            }
        } else {
            // Fallback: Eğer browser date inputu desteklemiyorsa ve kullanıcı farklı girdiys
            date = dateRaw;
        }
    }

    const bakimPlaniText = typeValue === 'planned' ? 'VAR' : typeValue === 'periodic' ? 'PERİYODİK BAKIM' : 'YOK';
    const isPlanned = typeValue === 'planned';
    const isPeriodic = typeValue === 'periodic';
    const isPlannerMode = currentSettings.zimmetMode === 'PLANNER';

    // ACID: dolu ve periyodik değilse mail gövdesine ekle
    const acidValue = document.getElementById('txtAcid').value.toUpperCase().trim();
    const acidLine = (acidValue && !isPeriodic) ? `<br><strong>ACID:</strong> ${acidValue}` : '';

    let html = `
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #f4f5f7;">
            <tr>
                <td align="center" style="padding: 40px 15px;">
                    <!--[if mso]>
                    <table align="center" border="0" cellspacing="0" cellpadding="0" width="650">
                    <tr>
                    <td align="center" valign="top" width="650">
                    <![endif]-->
                    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width: 650px; border: 1px solid #d1d5db; border-radius: 8px; font-family: 'Segoe UI', Tahoma, sans-serif; color: #000000; line-height: 1.6; text-align: left; background-color: #ffffff; margin: 0 auto;">
                        <tr>
                            <td style="padding: 20px;">
                                <p style="margin-top: 0; font-size: 14px;">Sayın İlgililer,</p>
                                <p style="font-size: 14px;">Aşağıda bilgileri bulunan bakım paketi hazır olup <strong>BMPM</strong> ofisinden teslim alınabilir.</p>
                                
                                <div style="background-color: #f4f5f6; border-left: 4px solid #E2001A; padding: 10px 15px; border-radius: 5px; margin-bottom: 15px; font-size: 13px; line-height: 1.5;">
                                    <strong>A/C:</strong> ${ac}<br>
                                    <strong>BAKIM ADI:</strong> ${bakim}<br>
                                    <strong>BAKIM GİRİŞ TARİHİ:</strong> ${date}<br>
                                    <strong>BAKIM PLANI:</strong> ${bakimPlaniText}
                                    ${acidLine}
                                </div>
    `;

    if (isPlannerMode) {
        html += `
                                <p style="font-size: 14px; margin-bottom: 10px;">Kartların zimmetleneceği <strong>planner sicili ve ismini</strong> bu e-posta yoluyla bildirmenizi rica ederiz.</p>
                                <table style="width: 100%; border-collapse: collapse; margin-top: 5px;">
                                    <tr>
                                        <th style="border: 1px solid #d1d5db; padding: 6px 10px; background-color: #E2001A; color: #ffffff; text-align: left; width: 1%; white-space: nowrap; font-weight: bold; font-size: 13px;">PLANNER</th>
                                        <td style="border: 1px solid #d1d5db; padding: 6px 10px; font-weight: 400 !important; font-style: normal !important; font-family: 'Segoe UI', Tahoma, sans-serif !important;">&nbsp;</td>
                                    </tr>
                                </table>
        `;
    } else if (isPlanned) {
        html += `
                                <p style="font-size: 14px; margin-bottom: 10px;">Alt tabloda belirtilen birimlere ait kartların kimlere zimmetleneceğini bu e-posta üzerinden bildirmenizi rica ederiz.</p>
                                <table style="width: 100%; border-collapse: collapse; margin-top: 5px;">
                                    <thead>
                                        <tr>
                                            <th style="border: 1px solid #d1d5db; padding: 6px 10px; background-color: #E2001A; color: #ffffff; text-align: left; width: 1%; white-space: nowrap; font-weight: bold; font-size: 13px;">BİRİM</th>
                                            <th style="border: 1px solid #d1d5db; padding: 6px 10px; background-color: #E2001A; color: #ffffff; text-align: left; font-weight: bold; font-size: 13px;">ZİMMET ÇIKILACAK SİCİL - İSİM</th>
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
                html += `<tr><td style="border: 1px solid #d1d5db; padding: 6px 10px; font-weight: bold; width: 1%; white-space: nowrap; font-size: 13px;">${skill.name}</td><td style="border: 1px solid #d1d5db; padding: 6px 10px; font-weight: 400 !important; font-style: normal !important; font-family: 'Segoe UI', Tahoma, sans-serif !important; font-size: 13px !important;">&nbsp;</td></tr>`;
            }
        });

        html += `</tbody></table>`;
    } else if (isPeriodic) {
        html += `
                                <p style="font-size: 14px; margin-bottom: 10px;">Periyodik bakım paketi olup, paket içeriğinin tamamı tek bir isim üzerine zimmetlenecektir. Kartların kime zimmetleneceğini tabloya işleyerek bu e-posta üzerinden bildirmenizi rica ederiz.</p>
                                <table style="width: 100%; border-collapse: collapse; margin-top: 5px;">
                                    <thead>
                                        <tr>
                                            <th style="border: 1px solid #d1d5db; padding: 6px 10px; background-color: #E2001A; color: #ffffff; text-align: left; width: 1%; white-space: nowrap; font-weight: bold; font-size: 13px;">BİRİM</th>
                                            <th style="border: 1px solid #d1d5db; padding: 6px 10px; background-color: #E2001A; color: #ffffff; text-align: left; font-weight: bold; font-size: 13px;">ZİMMET ÇIKILACAK SİCİL - İSİM</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr><td style="border: 1px solid #d1d5db; padding: 6px 10px; font-weight: bold; width: 1%; white-space: nowrap; font-size: 13px;">MEKANİK</td><td style="border: 1px solid #d1d5db; padding: 6px 10px; font-weight: 400 !important; font-style: normal !important; font-family: 'Segoe UI', Tahoma, sans-serif !important; font-size: 13px !important;">&nbsp;</td></tr>
                                    </tbody>
                                </table>
        `;
    } else {
        html += `
                                <p style="font-size: 14px; margin-bottom: 10px;">Bakım planı bulunmadığından, paket içeriğinin tamamı tek bir isim üzerine zimmetlenecektir. Kartların kime zimmetleneceğini tabloya işleyerek bu e-posta üzerinden bildirmenizi rica ederiz.</p>
                                <table style="width: 100%; border-collapse: collapse; margin-top: 5px;">
                                    <thead>
                                        <tr>
                                            <th style="border: 1px solid #d1d5db; padding: 6px 10px; background-color: #E2001A; color: #ffffff; text-align: left; width: 1%; white-space: nowrap; font-weight: bold; font-size: 13px;">BİRİM</th>
                                            <th style="border: 1px solid #d1d5db; padding: 6px 10px; background-color: #E2001A; color: #ffffff; text-align: left; font-weight: bold; font-size: 13px;">ZİMMET ÇIKILACAK SİCİL - İSİM</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        <tr><td style="border: 1px solid #d1d5db; padding: 6px 10px; font-weight: bold; width: 1%; white-space: nowrap; font-size: 13px;">MEKANİK</td><td style="border: 1px solid #d1d5db; padding: 6px 10px; font-weight: 400 !important; font-style: normal !important; font-family: 'Segoe UI', Tahoma, sans-serif !important; font-size: 13px !important;">&nbsp;</td></tr>
                                    </tbody>
                                </table>
        `;
    }

    html += `
                                <div style="margin-top: 20px; padding-top: 10px; border-top: 1px solid #d1d5db; font-size: 12px; color: #6c757d; line-height: 1.4;">
                                    <strong>BAKIM MÜHENDİSLİK VE PLANLAMA MÜDÜRLÜĞÜ</strong><br>
                                    BAKIM PLANLAMA ŞEFLİĞİ (SAW)
                                </div>
                            </td>
                        </tr>
                    </table>
                    <!--[if mso]>
                    </td>
                    </tr>
                    </table>
                    <![endif]-->
                </td>
            </tr>
        </table>
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
        console.error("PackNotice: Önizleme güncellenemedi", err);
    }
}

// Sayfa içi uyarı modali (alert() Office taskpane'de çalışmadığı için)
function showAlert(message, title) {
    const overlay = document.getElementById('alertOverlay');
    const msgEl = document.getElementById('alertMessage');
    const titleEl = document.getElementById('alertTitle');
    if (overlay && msgEl) {
        msgEl.innerHTML = message;
        if (titleEl) titleEl.textContent = title || '⚠️ UYARI';
        overlay.classList.remove('hidden');
    } else {
        // Fallback: konsola yaz
        console.error("PackNotice UYARI:", message);
    }
}

// Onay modali göster
function showConfirm(message, onYes) {
    const overlay = document.getElementById('confirmOverlay');
    const msgEl = document.getElementById('confirmMessage');
    const btnYes = document.getElementById('btnConfirmYes');

    if (overlay && msgEl && btnYes) {
        msgEl.innerHTML = message;
        btnYes.onclick = onYes;
        overlay.classList.remove('hidden');
    }
}

async function prepareMail() {
    const ac = document.getElementById('txtAc').value.toUpperCase().trim();
    const bakim = document.getElementById('txtBakim').value.toUpperCase().trim();
    const dateRaw = document.getElementById('dateInput').value;

    if (!ac || !bakim || !dateRaw) {
        showAlert("<b style='color: var(--accent);'>A/C</b>, <b style='color: var(--accent);'>BAKIM ADI</b> VE <b style='color: var(--accent);'>TARİH</b> ALANLARI BOŞ BIRAKILAMAZ!");
        return;
    }

    // A/C format kontrolü: tire (-) içermeli (örn. TC-ABC)
    if (!ac.includes('-')) {
        showAlert(
            "<b>UÇAK TESCİL KODU (A/C) HATALI GÖRÜNÜYOR.</b><br><br>" +
            "GİRDİĞİNİZ DEĞER: <b style='color: var(--accent);'>" + ac + "</b><br><br>" +
            "TESCİL KODUNUN TİRE (-) İÇERMESİ GEREKİR.<br>" +
            "ÖRNEK: <b>TC-ABC</b>",
            "⚠️ FORMAT UYARISI"
        );
        return;
    }

    // Office konteksti yoksa engelle
    if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
        console.warn("PackNotice: Office context not available.");
        showAlert("OUTLOOK BAĞLANTISI KURULAMADI.<br>EKLENTİYİ OUTLOOK İÇİNDEN AÇTIĞINIZDAN EMİN OLUN.");
        return;
    }

    const item = Office.context.mailbox.item;

    // from.getAsync yoksa da engelle
    if (!item.from || typeof item.from.getAsync !== 'function') {
        console.warn("PackNotice: item.from.getAsync not available.");
        showAlert(
            "<b>GÖNDERİCİ ADRESİ KONTROL EDİLEMİYOR.</b><br><br>" +
            "LÜTFEN MAİLİN <b>'KİMDEN' (FROM)</b> ALANINDAN<br>" +
            "<b style='color: var(--accent);'>TT-UBB(SAW)-BAKIMHAZIRLIK</b><br>" +
            "HESABININ SEÇİLİ OLDUĞUNDAN EMİN OLUN."
        );
        return;
    }

    // Gönderici adresi kontrolü
    item.from.getAsync(function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const currentFrom = (result.value.emailAddress || "").toUpperCase();
            console.log("PackNotice: prepareMail check - current From:", currentFrom);

            if (currentFrom !== targetAddress) {
                // Gönderici adresi yanlış - mail hazırlamayı engelle
                showAlert(
                    "<b>GÖNDERİCİ ADRESİ YANLIŞ!</b><br><br>" +
                    "MAİL HAZIRLANAMAZ.<br><br>" +
                    "LÜTFEN MAİLİN <b>'KİMDEN' (FROM)</b> ALANINDAN<br>" +
                    "<b style='color: var(--accent);'>TT-UBB(SAW)-BAKIMHAZIRLIK</b><br>" +
                    "HESABINI SEÇİN VE TEKRAR DENEYİN."
                );
                console.warn("PackNotice: Preparation blocked - wrong sender: " + currentFrom);
                return;
            }

            // Gönderici doğru - mail hazırla
            executeMailPreparation();
        } else {
            // from.getAsync başarısız olursa da uyarı ver
            showAlert(
                "<b>GÖNDERİCİ ADRESİ DOĞRULANAMADI!</b><br><br>" +
                "LÜTFEN MAİLİN <b>'KİMDEN' (FROM)</b> ALANINDAN<br>" +
                "<b style='color: var(--accent);'>TT-UBB(SAW)-BAKIMHAZIRLIK</b><br>" +
                "HESABININ SEÇİLİ OLDUĞUNDAN EMİN OLUN."
            );
            console.error("PackNotice: from.getAsync failed:", result.error ? result.error.message : "unknown");
        }
    });
}

function executeMailPreparation() {
    const ac = document.getElementById('txtAc').value.toUpperCase().trim();
    const bakim = document.getElementById('txtBakim').value.toUpperCase().trim();
    const bay = document.querySelector('input[name="bay"]:checked').value;
    const subject = `${ac} / ${bakim} BAKIM PAKETİ HK.`;
    const body = generateHTML();

    const toRecipients = bay === 'BAY-1' ? currentSettings.bay1To : bay === 'BAY-2' ? currentSettings.bay2To : currentSettings.bay3To;
    const ccRecipients = currentSettings.cc;

    const item = Office.context.mailbox.item;
    finalizeMailPreparation(item, subject, toRecipients, ccRecipients, body);
}

function finalizeMailPreparation(item, subject, toRecipients, ccRecipients, body) {
    item.subject.setAsync(subject);
    item.to.setAsync(toRecipients); // Office.js accepts arrays
    item.cc.setAsync(ccRecipients);
    item.body.setAsync(body, { coercionType: Office.CoercionType.Html });
    console.log("Mail prepared successfully!");
}
