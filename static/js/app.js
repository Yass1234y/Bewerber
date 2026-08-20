
// ==========================================
// STATE MANAGEMENT
// ==========================================
const App = {
    state: {
        loggedIn: false,
        isAdmin: false,
        currentPage: 'anschreiben',
        files: {
            excel: null,
            cv: null,
            template: null,
            other: null,
            letter: null,
            extra: null,
            zip: null,
            direct_excel: null,
            direct_letter: null,
            direct_pdf: null,
        },
        generating: false,
        sending: false,
        generated: false,
        genPollInterval: null,
        sendPollInterval: null,
        countdownInterval: null,
    },

    // ==========================================
    // INIT
    // ==========================================
    init() {
        this.bindEvents();
        this.checkInitialState();
    },

    checkInitialState() {
        fetch('/api/state')
            .then(r => r.json())
            .then(data => {
                if (data.logged_in) {
                    this.state.loggedIn = true;
                    this.state.isAdmin = data.is_admin;
                    this.showMainApp();
                    if (data.generating) {
                        this.startGeneratePolling();
                    }
                    if (data.sending || data.waiting_scheduled) {
                        this.startSendPolling();
                    }
                    if (data.generated) {
                        this.updateSendPage();
                    }
                }
            })
            .catch(() => {});
    },

    // ==========================================
    // EVENT BINDING
    // ==========================================
    bindEvents() {
        // Login
        document.getElementById('login-btn').addEventListener('click', () => this.login());
        document.getElementById('access-code').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') this.login();
        });
        
                // Mode toggle
        document.getElementById('mode-toggle').addEventListener('change', (e) => this.toggleMode(e.target.checked));

        // Zip upload
        this.bindFileUpload('zip-input', 'zip');

        // Import
        document.getElementById('import-btn').addEventListener('click', () => this.startImport());
        
        // Export generated files
        document.getElementById('export-generated-btn').addEventListener('click', () => this.exportGenerated());

        // Navigation
        document.querySelectorAll('.nav-btn').forEach(btn => {
            btn.addEventListener('click', () => this.navigate(btn.dataset.page));
        });

        // File uploads
        this.bindFileUpload('excel-input', 'excel');
        this.bindFileUpload('cv-input', 'cv');
        this.bindFileUpload('template-input', 'template');
        this.bindFileUpload('other-input', 'other');
        this.bindFileUpload('letter-input', 'letter');
        this.bindFileUpload('extra-input', 'extra');

        // Generate
        document.getElementById('generate-btn').addEventListener('click', () => this.startGenerate());

        // Send
        document.getElementById('send-btn').addEventListener('click', () => this.startSend());
        
                // Send mode toggle (with/without generate)
        document.getElementById('send-mode-toggle').addEventListener('change', (e) => this.toggleSendMode(e.target.checked));

        // Direct file uploads
        this.bindFileUpload('direct-excel-input', 'direct_excel');
        this.bindFileUpload('direct-letter-input', 'direct_letter');
        this.bindFileUpload('direct-pdf-input', 'direct_pdf');

        // Direct send
        document.getElementById('send-direct-btn').addEventListener('click', () => this.startDirectSend());

        // Send mode
        document.querySelectorAll('input[name="send-mode"]').forEach(radio => {
            radio.addEventListener('change', (e) => this.toggleSchedule(e.target.value));
        });

        // Schedule inputs
        ['schedule-date', 'schedule-hour', 'schedule-minute'].forEach(id => {
            document.getElementById(id).addEventListener('change', () => this.updateScheduleInfo());
        });

        // Resume
        document.getElementById('resume-btn').addEventListener('click', () => this.resumeSend());

        // New application
        document.getElementById('new-application-btn').addEventListener('click', () => this.resetAll());

        // Dashboard filters
        document.getElementById('filter-session').addEventListener('change', () => this.loadDashboard());
        document.getElementById('filter-time').addEventListener('change', () => this.loadDashboard());
    },

    bindFileUpload(inputId, type) {
        const input = document.getElementById(inputId);
        if (!input) return;
        input.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                this.state.files[type] = file;
                const nameEl = document.getElementById(`${type}-name`);
                if (nameEl) nameEl.textContent = file.name;
                input.closest('.upload-card').classList.add('has-file');
                this.updateButtons();
            }
        });
    },

    updateButtons() {
        // Generate button
        const genBtn = document.getElementById('generate-btn');
        const genReady = this.state.files.excel && this.state.files.cv &&
                         this.state.files.template && this.state.files.other;
        genBtn.disabled = !genReady;

        // Send button (with generate)
        const sendBtn = document.getElementById('send-btn');
        const letterReady = this.state.files.letter;
        const scheduleMode = document.querySelector('input[name="send-mode"]:checked').value === 'schedule';
        let scheduleValid = true;
        if (scheduleMode) {
            scheduleValid = this.validateSchedule();
        }
        sendBtn.disabled = !letterReady || !scheduleValid;

        // Direct send button (without generate)
        const directBtn = document.getElementById('send-direct-btn');
        if (directBtn) {
            const directReady = this.state.files.direct_excel && this.state.files.direct_letter && this.state.files.direct_pdf;
            directBtn.disabled = !directReady || !scheduleValid;
        }

        // Import button
        const importBtn = document.getElementById('import-btn');
        if (importBtn) {
            importBtn.disabled = !this.state.files.zip;
        }
    },

    // ==========================================
    // LOGIN
    // ==========================================
    async login() {
        const code = document.getElementById('access-code').value.trim();
        const errorEl = document.getElementById('login-error');
        errorEl.textContent = '';

        if (!code) {
            errorEl.textContent = 'Bitte Code eingeben';
            return;
        }

        try {
            const resp = await fetch('/api/login', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ code }),
            });
            const data = await resp.json();

            if (data.success) {
                this.state.loggedIn = true;
                this.state.isAdmin = data.is_admin;
                this.showMainApp();
            } else {
                errorEl.textContent = data.error || 'Falscher Code';
            }
        } catch (e) {
            errorEl.textContent = 'Verbindungsfehler';
        }
    },

    showMainApp() {
        document.getElementById('login-screen').classList.add('hidden');
        document.getElementById('main-app').classList.remove('hidden');

        if (this.state.isAdmin) {
            document.querySelectorAll('.admin-only').forEach(el => el.classList.remove('hidden'));
        }

        this.navigate('anschreiben');
        this.checkGeneratedStatus();
    },

    // ==========================================
    // NAVIGATION
    // ==========================================
    navigate(page) {
        this.state.currentPage = page;

        // Update nav buttons
        document.querySelectorAll('.nav-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.page === page);
        });

        // Show page
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        document.getElementById(`page-${page}`).classList.add('active');

        // Page-specific actions
        if (page === 'send-email') {
            this.updateSendPage();
        } else if (page === 'dashboard') {
            this.loadDashboard();
        }
    },

    // ==========================================
    // GENERATE
    // ==========================================
    async startGenerate() {
        const bewerbungsname = document.getElementById('bewerbungsname').value.trim() || 'Bewerbung';
        const position = parseInt(document.getElementById('anschreiben-pos').value) || 2;

        const formData = new FormData();
        formData.append('excel', this.state.files.excel);
        formData.append('cv', this.state.files.cv);
        formData.append('template', this.state.files.template);
        formData.append('other', this.state.files.other);
        formData.append('position', position);
        formData.append('bewerbungsname', bewerbungsname);

        document.getElementById('generate-btn').disabled = true;
        document.getElementById('gen-form').classList.add('hidden');
        document.querySelector('.config-section').classList.add('hidden');
        document.getElementById('gen-progress').classList.remove('hidden');

        try {
            const resp = await fetch('/api/generate', { method: 'POST', body: formData });
            const data = await resp.json();

            if (data.success) {
                this.startGeneratePolling();
            } else {
                this.toast(data.error || 'Fehler beim Starten', 'error');
                document.getElementById('generate-btn').disabled = false;
                document.getElementById('gen-form').classList.remove('hidden');
                document.querySelector('.config-section').classList.remove('hidden');
                document.getElementById('gen-progress').classList.add('hidden');
            }
        } catch (e) {
            this.toast('Verbindungsfehler', 'error');
            document.getElementById('generate-btn').disabled = false;
            document.getElementById('gen-form').classList.remove('hidden');
            document.querySelector('.config-section').classList.remove('hidden');
            document.getElementById('gen-progress').classList.add('hidden');
        }
    },

    startGeneratePolling() {
        if (this.state.genPollInterval) clearInterval(this.state.genPollInterval);

        this.state.genPollInterval = setInterval(async () => {
            try {
                const resp = await fetch('/api/generate/status');
                const data = await resp.json();

                // Update progress bar
                const pct = Math.round(data.progress * 100);
                document.getElementById('gen-progress-bar').style.width = pct + '%';
                document.getElementById('gen-progress-text').textContent =
                    `${data.progress > 0 ? Math.round(data.progress * data.total) : 0} / ${data.total}`;

                // Update log
                this.updateLog('gen-log', data.log);

                if (!data.generating) {
                    clearInterval(this.state.genPollInterval);
                    this.state.genPollInterval = null;

                    if (data.generated) {
                        document.getElementById('gen-progress').classList.add('hidden');
                        document.getElementById('gen-success').classList.remove('hidden');
                        this.state.generated = true;
                        this.toast('Generierung abgeschlossen!', 'success');
                    }
                }
            } catch (e) {}
        }, 1000);
    },
    
        // ==========================================
    // MODE TOGGLE (Generate / Import)
    // ==========================================
    toggleMode(isImport) {
        const generateSection = document.getElementById('generate-mode-section');
        const importSection = document.getElementById('import-mode-section');
        const labelGenerate = document.getElementById('label-generate');
        const labelImport = document.getElementById('label-import');

        if (isImport) {
            generateSection.classList.add('hidden');
            importSection.classList.remove('hidden');
            labelGenerate.classList.remove('active');
            labelImport.classList.add('active');
        } else {
            generateSection.classList.remove('hidden');
            importSection.classList.add('hidden');
            labelGenerate.classList.add('active');
            labelImport.classList.remove('active');
        }

        this.updateButtons();
    },

    // ==========================================
    // IMPORT
    // ==========================================
    async startImport() {
        if (!this.state.files.zip) return;

        const formData = new FormData();
        formData.append('zip', this.state.files.zip);

        document.getElementById('import-btn').disabled = true;
        document.getElementById('loading-overlay').classList.remove('hidden');

        try {
            const resp = await fetch('/api/generate/import', {
                method: 'POST',
                body: formData
            });
            const data = await resp.json();

            document.getElementById('loading-overlay').classList.add('hidden');

            if (data.success) {
                this.state.generated = true;
                document.getElementById('import-mode-section').classList.add('hidden');
                document.getElementById('gen-success').classList.remove('hidden');
                document.getElementById('gen-success-title').textContent = 'Import erfolgreich!';
                this.toast(`${data.total_companies} Unternehmen importiert!`, 'success');
            } else {
                this.toast(data.error || 'Import fehlgeschlagen', 'error');
                document.getElementById('import-btn').disabled = false;
            }
        } catch (e) {
            document.getElementById('loading-overlay').classList.add('hidden');
            this.toast('Verbindungsfehler', 'error');
            document.getElementById('import-btn').disabled = false;
        }
    },
    
    // ==========================================
    // EXPORT GENERATED
    // ==========================================
    exportGenerated() {
        this.toast('ZIP wird erstellt...', 'info');
        window.location.href = '/api/generate/export';
    },

    // ==========================================
    // SEND EMAIL
    // ==========================================
    async checkGeneratedStatus() {
        try {
            const resp = await fetch('/api/state');
            const data = await resp.json();
            this.state.generated = data.generated;
            this.state.isAdmin = data.is_admin;
        } catch (e) {}
    },

        updateSendPage() {
        fetch('/api/send/status')
            .then(r => r.json())
            .then(data => {
                // If direct_mode from backend, switch toggle
                if (data.direct_mode) {
                    document.getElementById('send-mode-toggle').checked = true;
                    document.getElementById('send-label-with').classList.remove('active');
                    document.getElementById('send-label-without').classList.add('active');
                    document.getElementById('send-btn').classList.add('hidden');
                    document.getElementById('send-direct-btn').classList.remove('hidden');
                }

                // 1. Send done → show success
                if (data.send_done) {
                    document.getElementById('send-form').classList.add('hidden');
                    document.getElementById('send-direct-form').classList.add('hidden');
                    document.getElementById('send-config-section').classList.add('hidden');
                    document.getElementById('send-success').classList.remove('hidden');
                    return;
                }

                // 2. Interrupted → show interrupted card
                if (data.interrupted_at) {
                    document.getElementById('send-form').classList.add('hidden');
                    document.getElementById('send-direct-form').classList.add('hidden');
                    document.getElementById('send-config-section').classList.add('hidden');
                    document.getElementById('send-interrupted').classList.remove('hidden');
                    document.getElementById('resume-num').value = data.interrupted_at;
                    document.getElementById('resume-num').max = data.total;
                    return;
                }

                // 3. Sending or waiting → show countdown/progress
                if (data.sending || data.waiting_scheduled) {
                    document.getElementById('send-form').classList.add('hidden');
                    document.getElementById('send-direct-form').classList.add('hidden');
                    document.getElementById('send-config-section').classList.add('hidden');
                    if (data.waiting_scheduled) {
                        document.getElementById('schedule-countdown').classList.remove('hidden');
                        this.startCountdown(data.scheduled_dt);
                    } else {
                        document.getElementById('send-progress').classList.remove('hidden');
                        this.startSendPolling();
                    }
                    return;
                }

                // 4. Not sending → show appropriate form
                const isDirect = document.getElementById('send-mode-toggle').checked;

                if (isDirect) {
                    // Without Generate mode → always show direct form
                    document.getElementById('send-warning').classList.add('hidden');
                    document.getElementById('send-form').classList.add('hidden');
                    document.getElementById('send-direct-form').classList.remove('hidden');
                    document.getElementById('send-config-section').classList.remove('hidden');
                    document.getElementById('send-info').textContent = 'Direct Send Mode — Keine Generierung erforderlich';
                } else {
                    // With Generate mode → check if generated
                    fetch('/api/state')
                        .then(r => r.json())
                        .then(stateData => {
                            if (stateData.generated) {
                                document.getElementById('send-warning').classList.add('hidden');
                                document.getElementById('send-form').classList.remove('hidden');
                                document.getElementById('send-direct-form').classList.add('hidden');
                                document.getElementById('send-config-section').classList.remove('hidden');
                                document.getElementById('send-info').textContent =
                                    `Folder: ${stateData.bewerbungsname} — Unternehmen: ${stateData.total_companies}`;
                            } else {
                                document.getElementById('send-warning').classList.remove('hidden');
                                document.getElementById('send-form').classList.add('hidden');
                                document.getElementById('send-direct-form').classList.add('hidden');
                                document.getElementById('send-config-section').classList.add('hidden');
                            }
                        });
                }
            });
    },

    toggleSchedule(mode) {
        const schedConfig = document.getElementById('schedule-config');
        if (mode === 'schedule') {
            schedConfig.classList.remove('hidden');
            // Set default date to today
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('schedule-date').value = today;
            this.updateScheduleInfo();
        } else {
            schedConfig.classList.add('hidden');
            document.getElementById('schedule-info').classList.add('hidden');
            document.getElementById('schedule-error').classList.add('hidden');
        }
        this.updateButtons();
    },

    validateSchedule() {
        const date = document.getElementById('schedule-date').value;
        const hour = parseInt(document.getElementById('schedule-hour').value);
        const minute = parseInt(document.getElementById('schedule-minute').value);

        if (!date) return false;

        const scheduled = new Date(date);
        scheduled.setHours(hour, minute, 0, 0);
        const now = new Date();

        return scheduled > now;
    },

    updateScheduleInfo() {
        if (!this.validateSchedule()) {
            document.getElementById('schedule-info').classList.add('hidden');
            document.getElementById('schedule-error').classList.remove('hidden');
            document.getElementById('schedule-error').textContent =
                'Die gewählte Zeit liegt in der Vergangenheit!';
            this.updateButtons();
            return;
        }

        document.getElementById('schedule-error').classList.add('hidden');
        const date = document.getElementById('schedule-date').value;
        const hour = parseInt(document.getElementById('schedule-hour').value);
        const minute = parseInt(document.getElementById('schedule-minute').value);

        const scheduled = new Date(date);
        scheduled.setHours(hour, minute, 0, 0);
        const now = new Date();
        const diff = scheduled - now;
        const h = Math.floor(diff / 3600000);
        const m = Math.floor((diff % 3600000) / 60000);

        const infoEl = document.getElementById('schedule-info');
        infoEl.classList.remove('hidden');
        infoEl.textContent = `Senden um: ${scheduled.toLocaleString('de-DE')} — in ${h}h ${m}m`;

        this.updateButtons();
    },

    getScheduledDT() {
        const date = document.getElementById('schedule-date').value;
        const hour = parseInt(document.getElementById('schedule-hour').value);
        const minute = parseInt(document.getElementById('schedule-minute').value);
        const scheduled = new Date(date);
        scheduled.setHours(hour, minute, 0, 0);
        return scheduled.toISOString();
    },

    async startSend() {
        const delay = parseInt(document.getElementById('delay-input').value) || 10;
        const startNum = parseInt(document.getElementById('start-input').value) || 1;
        const sendMode = document.querySelector('input[name="send-mode"]:checked').value;
        const scheduledDT = sendMode === 'schedule' ? this.getScheduledDT() : '';

        const formData = new FormData();
        formData.append('letter', this.state.files.letter);
        if (this.state.files.extra) {
            formData.append('extra', this.state.files.extra);
        }
        formData.append('delay', delay);
        formData.append('start', startNum);
        formData.append('scheduled_dt', scheduledDT);

        document.getElementById('send-btn').disabled = true;

        try {
            const resp = await fetch('/api/send', { method: 'POST', body: formData });
            const data = await resp.json();

            if (data.success) {
                document.getElementById('send-form').classList.add('hidden');

                if (sendMode === 'schedule') {
                    document.getElementById('schedule-countdown').classList.remove('hidden');
                    this.startCountdown(scheduledDT);
                }

                this.startSendPolling();
            } else {
                this.toast(data.error || 'Fehler', 'error');
                document.getElementById('send-btn').disabled = false;
            }
        } catch (e) {
            this.toast('Verbindungsfehler', 'error');
            document.getElementById('send-btn').disabled = false;
        }
    },
    
        // ==========================================
    // TOGGLE SEND MODE (With/Without Generate)
    // ==========================================
    toggleSendMode(isDirect) {
        const labelWith = document.getElementById('send-label-with');
        const labelWithout = document.getElementById('send-label-without');
        const sendBtn = document.getElementById('send-btn');
        const directBtn = document.getElementById('send-direct-btn');

        if (isDirect) {
            labelWith.classList.remove('active');
            labelWithout.classList.add('active');
            sendBtn.classList.add('hidden');
            directBtn.classList.remove('hidden');
        } else {
            labelWith.classList.add('active');
            labelWithout.classList.remove('active');
            sendBtn.classList.remove('hidden');
            directBtn.classList.add('hidden');
        }

        this.updateSendPage();
    },

    // ==========================================
    // DIRECT SEND (Without Generate)
    // ==========================================
    async startDirectSend() {
        const delay = parseInt(document.getElementById('delay-input').value) || 10;
        const startNum = parseInt(document.getElementById('start-input').value) || 1;
        const sendMode = document.querySelector('input[name="send-mode"]:checked').value;
        const scheduledDT = sendMode === 'schedule' ? this.getScheduledDT() : '';

        const formData = new FormData();
        formData.append('excel', this.state.files.direct_excel);
        formData.append('letter', this.state.files.direct_letter);
        formData.append('pdf', this.state.files.direct_pdf);
        formData.append('delay', delay);
        formData.append('start', startNum);
        formData.append('scheduled_dt', scheduledDT);

        document.getElementById('send-direct-btn').disabled = true;

        try {
            const resp = await fetch('/api/send/direct', { method: 'POST', body: formData });
            const data = await resp.json();

            if (data.success) {
                document.getElementById('send-direct-form').classList.add('hidden');
                document.getElementById('send-config-section').classList.add('hidden');

                if (sendMode === 'schedule') {
                    document.getElementById('schedule-countdown').classList.remove('hidden');
                    this.startCountdown(scheduledDT);
                }

                this.startSendPolling();
            } else {
                this.toast(data.error || 'Fehler', 'error');
                document.getElementById('send-direct-btn').disabled = false;
            }
        } catch (e) {
            this.toast('Verbindungsfehler', 'error');
            document.getElementById('send-direct-btn').disabled = false;
        }
    },

    startCountdown(scheduledDT) {
        if (this.state.countdownInterval) clearInterval(this.state.countdownInterval);

        const update = () => {
            const target = new Date(scheduledDT);
            const now = new Date();
            const remaining = target - now;

            if (remaining <= 0) {
                clearInterval(this.state.countdownInterval);
                this.state.countdownInterval = null;
                document.getElementById('countdown-timer').textContent = '00:00:00';
                document.getElementById('countdown-text').textContent = 'Zeit erreicht! Senden beginnt...';
                document.getElementById('schedule-countdown').classList.add('hidden');
                document.getElementById('send-progress').classList.remove('hidden');
                return;
            }

            const h = Math.floor(remaining / 3600000);
            const m = Math.floor((remaining % 3600000) / 60000);
            const s = Math.floor((remaining % 60000) / 1000);

            document.getElementById('countdown-timer').textContent =
                `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
            document.getElementById('countdown-text').textContent =
                `Senden um: ${target.toLocaleString('de-DE')}`;
        };

        update();
        this.state.countdownInterval = setInterval(update, 1000);
    },

    startSendPolling() {
        if (this.state.sendPollInterval) clearInterval(this.state.sendPollInterval);

        this.state.sendPollInterval = setInterval(async () => {
            try {
                const resp = await fetch('/api/send/status');
                const data = await resp.json();

                // If waiting for schedule
                if (data.waiting_scheduled) {
                    if (!document.getElementById('schedule-countdown').classList.contains('hidden') === false) {
                        document.getElementById('schedule-countdown').classList.remove('hidden');
                        this.startCountdown(data.scheduled_dt);
                    }
                    return;
                }

                // If countdown is done and sending started
                if (data.sending) {
                    document.getElementById('schedule-countdown').classList.add('hidden');
                    document.getElementById('send-progress').classList.remove('hidden');

                    const pct = Math.round(data.progress * 100);
                    document.getElementById('send-progress-bar').style.width = pct + '%';
                    document.getElementById('send-progress-text').textContent =
                        `${data.progress > 0 ? Math.round(data.progress * data.total) : 0} / ${data.total}`;

                    this.updateLog('send-log', data.log);
                }

                // Check for interruption
                if (data.interrupted_at) {
                    clearInterval(this.state.sendPollInterval);
                    this.state.sendPollInterval = null;
                    document.getElementById('send-progress').classList.add('hidden');
                    document.getElementById('send-interrupted').classList.remove('hidden');
                    document.getElementById('resume-num').value = data.interrupted_at;
                    document.getElementById('resume-num').max = data.total;
                    this.toast('Verbindung unterbrochen!', 'error');
                    return;
                }

                // Check for completion
                if (data.send_done) {
                    clearInterval(this.state.sendPollInterval);
                    this.state.sendPollInterval = null;
                    document.getElementById('send-progress').classList.add('hidden');
                    document.getElementById('send-success').classList.remove('hidden');
                    this.toast('Alle E-Mails gesendet!', 'success');
                }
            } catch (e) {}
        }, 1000);
    },

    async resumeSend() {
        const resumeFrom = parseInt(document.getElementById('resume-num').value) || 1;

        try {
            const resp = await fetch('/api/send/resume', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ resume_from: resumeFrom }),
            });
            const data = await resp.json();

            if (data.success) {
                document.getElementById('send-interrupted').classList.add('hidden');
                document.getElementById('send-progress').classList.remove('hidden');
                this.startSendPolling();
            } else {
                this.toast(data.error || 'Fehler', 'error');
            }
        } catch (e) {
            this.toast('Verbindungsfehler', 'error');
        }
    },

        async resetAll() {
        try {
            await fetch('/api/reset', { method: 'POST' });

            // Reset UI
            document.getElementById('send-success').classList.add('hidden');
            document.getElementById('send-interrupted').classList.add('hidden');
            document.getElementById('send-progress').classList.add('hidden');
            document.getElementById('gen-success').classList.add('hidden');
            document.getElementById('gen-progress').classList.add('hidden');
            document.getElementById('import-mode-section').classList.add('hidden');
            document.getElementById('generate-mode-section').classList.remove('hidden');

            // Reset mode toggle
            document.getElementById('mode-toggle').checked = false;
            this.toggleMode(false);

            // Reset files
            this.state.files = { excel: null, cv: null, template: null, other: null, letter: null, extra: null, zip: null };
            document.querySelectorAll('.upload-filename').forEach(el => el.textContent = '');
            document.querySelectorAll('.upload-card').forEach(el => el.classList.remove('has-file'));
            document.querySelectorAll('input[type="file"]').forEach(el => el.value = '');

            this.state.generated = false;
            this.updateButtons();
            this.navigate('anschreiben');
            this.toast('Zurückgesetzt', 'info');
        } catch (e) {
            this.toast('Fehler beim Zurücksetzen', 'error');
        }
    },

    // ==========================================
    // DASHBOARD
    // ==========================================
    async loadDashboard() {
        const session = document.getElementById('filter-session').value || 'Alle';
        const hours = document.getElementById('filter-time').value || 0;

        try {
            const resp = await fetch(`/api/dashboard?session=${encodeURIComponent(session)}&hours=${hours}`);
            const data = await resp.json();

            // Update stats
            document.getElementById('stat-generated').textContent = data.stats.generated;
            document.getElementById('stat-sent').textContent = data.stats.sent;
            document.getElementById('stat-skipped').textContent = data.stats.skipped;
            document.getElementById('stat-errors').textContent = data.stats.errors;

            // Update session filter
            const sessionSelect = document.getElementById('filter-session');
            const currentSession = sessionSelect.value;
            sessionSelect.innerHTML = '<option value="Alle">Alle</option>';
            data.sessions.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.textContent = s;
                sessionSelect.appendChild(opt);
            });
            sessionSelect.value = currentSession;

            // Update table
            const tbody = document.getElementById('dashboard-table-body');
            if (data.companies.length === 0) {
                tbody.innerHTML = '<tr><td colspan="6" class="no-data">Keine Daten vorhanden</td></tr>';
            } else {
                tbody.innerHTML = data.companies.map(c => {
                    const genColor = c.generiert === 'Ja' ? '#4ade80' : c.generiert === 'Nein' ? '#ff6b6b' : '#888';
                    const sendColor = c.gesendet === 'Ja' ? '#4ade80' : '#ff6b6b';
                    const errColor = c.fehler === 'Ja' ? '#f97316' : '#555';
                    return `<tr>
                        <td>${c.session || ''}</td>
                        <td style="color:${genColor};font-weight:600;">${c.generiert}</td>
                        <td style="color:${sendColor};font-weight:600;">${c.gesendet}</td>
                        <td style="color:#aaa;">${c.email_firma || ''}</td>
                        <td style="color:${errColor};font-weight:600;">${c.fehler}</td>
                        <td style="color:#555;font-size:12px;">${c.zeit || ''}</td>
                    </tr>`;
                }).join('');
            }

            // Load Excel files
            this.loadExcelFiles(session);
        } catch (e) {
            this.toast('Fehler beim Laden des Dashboards', 'error');
        }
    },

    async loadExcelFiles(session) {
        try {
            const resp = await fetch(`/api/dashboard/excels?session=${encodeURIComponent(session)}`);
            const data = await resp.json();

            const list = document.getElementById('excel-list');
            if (data.files.length === 0) {
                list.innerHTML = '<div class="no-data">Keine Excel-Dateien gespeichert.</div>';
                return;
            }

            list.innerHTML = data.files.map(f => `
                <div class="excel-item">
                    <div class="excel-info">
                        <span class="excel-filename">📄 ${f.filename}</span>
                        <span class="excel-meta">${f.session_name || ''} - ${f.uploaded_at || ''}</span>
                    </div>
                    <a href="/api/dashboard/excel/download/${f.id}" class="download-btn">Herunterladen</a>
                </div>
            `).join('');
        } catch (e) {}
    },

    // ==========================================
    // UTILITIES
    // ==========================================
    updateLog(elementId, logEntries) {
        const el = document.getElementById(elementId);
        if (!el) return;

        el.innerHTML = logEntries.map(entry => {
            const colorClass = entry.status || 'ok';
            const text = entry.firma || entry.email || '';
            return `<div class="log-entry ${colorClass}">
                <span class="log-num">${entry.num}</span>
                <span>${text}</span>
            </div>`;
        }).join('');

        el.scrollTop = el.scrollHeight;
    },

    toast(message, type = 'info') {
        const container = document.getElementById('toast-container');
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.textContent = message;
        container.appendChild(toast);

        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transform = 'translateX(100px)';
            toast.style.transition = 'all 0.3s ease';
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    },
};

// ==========================================
// START
// ==========================================
document.addEventListener('DOMContentLoaded', () => App.init());