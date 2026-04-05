const tabOrder = ["find", "sort", "rename", "convert", "media", "logs", "config", "settings", "about"];
const themeStorageKey = "argus-theme-mode";
const uiStateStorageKey = "argus-ui-state-v4";

const defaultUiState = {
    workTab: "find",
    findMode: "Copy",
    findDuplicateHandling: "Rename",
    findIncludeSubfolders: true,
    findExtensions: "",
    includeRename: true,
    metadataRenameMode: "TitleThenContent",
    includeConvert: false,
    convertSourceFormat: "Any",
    convertTargetFormat: "Keep",
    convertQuality: "Balanced",
    mediaMusicEnabled: true,
    mediaImagesEnabled: true,
    mediaVideoEnabled: true
};

const state = {
    serverState: null,
    selectedHistoryRunId: null,
    selectedLogPath: null,
    activeTab: "find",
    activeWorkTab: "find",
    currentAbortController: null,
    persistTimer: null,
    suggestedDestinationPath: null,
    suggestedManifestPath: null,
    hasHydrated: false,
    themeMode: localStorage.getItem(themeStorageKey) || "system",
    ui: loadUiState()
};

const colorScheme = window.matchMedia("(prefers-color-scheme: dark)");

const elements = {
    tabButtons: [...document.querySelectorAll(".tab-button[data-tab]")],
    tabPanels: [...document.querySelectorAll(".tab-panel")],
    sourcePath: document.getElementById("sourcePath"),
    destinationPath: document.getElementById("destinationPath"),
    manifestPath: document.getElementById("manifestPath"),
    rulesPath: document.getElementById("rulesPath"),
    browseSourceButton: document.getElementById("browseSourceButton"),
    browseDestinationButton: document.getElementById("browseDestinationButton"),
    browseManifestButton: document.getElementById("browseManifestButton"),
    browseRulesButton: document.getElementById("browseRulesButton"),
    findModeSelect: document.getElementById("findModeSelect"),
    findDuplicateHandlingSelect: document.getElementById("findDuplicateHandlingSelect"),
    findExtensions: document.getElementById("findExtensions"),
    findIncludeSubfolders: document.getElementById("findIncludeSubfolders"),
    sortModeSelect: document.getElementById("sortModeSelect"),
    sortDuplicateHandlingSelect: document.getElementById("sortDuplicateHandlingSelect"),
    sortOrganizationModeSelect: document.getElementById("sortOrganizationModeSelect"),
    sortIncludeSubfolders: document.getElementById("sortIncludeSubfolders"),
    sortUseExtensionSubfolders: document.getElementById("sortUseExtensionSubfolders"),
    includeRename: document.getElementById("includeRename"),
    metadataRenameModeSelect: document.getElementById("metadataRenameModeSelect"),
    renamePattern: document.getElementById("renamePattern"),
    includeConvert: document.getElementById("includeConvert"),
    convertSourceFormatSelect: document.getElementById("convertSourceFormatSelect"),
    convertTargetFormatSelect: document.getElementById("convertTargetFormatSelect"),
    convertQualitySelect: document.getElementById("convertQualitySelect"),
    mediaMusicEnabled: document.getElementById("mediaMusicEnabled"),
    mediaImagesEnabled: document.getElementById("mediaImagesEnabled"),
    mediaVideoEnabled: document.getElementById("mediaVideoEnabled"),
    themeSelect: document.getElementById("themeSelect"),
    writeTagManifest: document.getElementById("writeTagManifest"),
    cleanEmptyFolders: document.getElementById("cleanEmptyFolders"),
    openDestinationWhenDone: document.getElementById("openDestinationWhenDone"),
    exportReport: document.getElementById("exportReport"),
    reportFormatSelect: document.getElementById("reportFormatSelect"),
    copySummaryToClipboard: document.getElementById("copySummaryToClipboard"),
    previewButton: document.getElementById("previewButton"),
    runButton: document.getElementById("runButton"),
    stopButton: document.getElementById("stopButton"),
    undoLastButton: document.getElementById("undoLastButton"),
    openOutputButton: document.getElementById("openOutputButton"),
    openLogsButton: document.getElementById("openLogsButton"),
    copyConsoleButton: document.getElementById("copyConsoleButton"),
    progressTrack: document.getElementById("progressTrack"),
    progressFill: document.getElementById("progressFill"),
    progressLabel: document.getElementById("progressLabel"),
    resultLabel: document.getElementById("resultLabel"),
    activeTaskLabel: document.getElementById("activeTaskLabel"),
    outputConsole: document.getElementById("outputConsole"),
    historyTableBody: document.querySelector("#historyTable tbody"),
    historyDetail: document.getElementById("historyDetail"),
    refreshHistoryButton: document.getElementById("refreshHistoryButton"),
    undoSelectedButton: document.getElementById("undoSelectedButton"),
    openHistoryLogButton: document.getElementById("openHistoryLogButton"),
    openOperationLogButton: document.getElementById("openOperationLogButton"),
    logsTableBody: document.querySelector("#logsTable tbody"),
    logDetail: document.getElementById("logDetail"),
    refreshLogsButton: document.getElementById("refreshLogsButton"),
    openSelectedLogButton: document.getElementById("openSelectedLogButton"),
    configRulesPath: document.getElementById("configRulesPath"),
    configSettingsPath: document.getElementById("configSettingsPath"),
    configHistoryPath: document.getElementById("configHistoryPath"),
    configLogsPath: document.getElementById("configLogsPath"),
    configReportsPath: document.getElementById("configReportsPath"),
    configBrowseRulesButton: document.getElementById("configBrowseRulesButton"),
    configOpenRulesButton: document.getElementById("configOpenRulesButton"),
    configUseDefaultRulesButton: document.getElementById("configUseDefaultRulesButton"),
    openSettingsPathButton: document.getElementById("openSettingsPathButton"),
    openHistoryPathButton: document.getElementById("openHistoryPathButton"),
    openLogsPathButton: document.getElementById("openLogsPathButton"),
    openReportsPathButton: document.getElementById("openReportsPathButton"),
    refreshRulesPreviewButton: document.getElementById("refreshRulesPreviewButton"),
    rulesPreview: document.getElementById("rulesPreview"),
    aboutVersion: document.getElementById("aboutVersion"),
    aboutHostUrl: document.getElementById("aboutHostUrl"),
    aboutAppDirectory: document.getElementById("aboutAppDirectory"),
    aboutDataDirectory: document.getElementById("aboutDataDirectory"),
    capShellMode: document.getElementById("capShellMode"),
    capImagePipeline: document.getElementById("capImagePipeline"),
    capAudioVideoPipeline: document.getElementById("capAudioVideoPipeline"),
    capDocumentPipeline: document.getElementById("capDocumentPipeline"),
    presetButtons: [...document.querySelectorAll("[data-preset]")],
    sortPresetButtons: [...document.querySelectorAll("[data-sort-mode]")]
};

document.addEventListener("DOMContentLoaded", init);
colorScheme.addEventListener("change", () => {
    if (state.themeMode === "system") {
        applyTheme("system");
    }
});

async function init() {
    bindEvents();
    applyTheme(state.themeMode);
    applyLocalUiState();
    updateTaskIndicator();
    updateControlAvailability();
    startHeartbeat();
    await refreshState(true);
    elements.outputConsole.textContent = "";
}

function bindEvents() {
    elements.tabButtons.forEach(button => button.addEventListener("click", () => setActiveTab(button.dataset.tab, true)));
    elements.themeSelect.addEventListener("change", () => {
        applyTheme(elements.themeSelect.value || "system");
        saveUiState();
    });

    window.addEventListener("keydown", event => {
        if (event.ctrlKey && event.shiftKey && event.key === "Enter") {
            event.preventDefault();
            void runOperation("/api/operations/run", "Running");
            return;
        }

        if (event.ctrlKey && !event.shiftKey && event.key === "Enter") {
            event.preventDefault();
            void runOperation("/api/operations/preview", "Previewing");
        }
    });

    elements.browseSourceButton.addEventListener("click", async () => {
        const path = await browsePath("/api/dialog/source", elements.sourcePath.value);
        if (path) {
            elements.sourcePath.value = path;
            handleSourcePathChanged();
            queuePersistSettings();
        }
    });

    elements.browseDestinationButton.addEventListener("click", async () => {
        const path = await browsePath("/api/dialog/destination", elements.destinationPath.value);
        if (path) {
            elements.destinationPath.value = path;
            handleDestinationPathChanged();
            queuePersistSettings();
        }
    });

    elements.browseManifestButton.addEventListener("click", async () => {
        const path = await browsePath("/api/dialog/manifest", elements.manifestPath.value);
        if (path) {
            elements.manifestPath.value = path;
            queuePersistSettings();
            updateControlAvailability();
        }
    });

    const browseRules = async () => {
        const path = await browsePath("/api/dialog/rules", elements.rulesPath.value);
        if (path) {
            elements.rulesPath.value = path;
            queuePersistSettings();
            await refreshRulesPreview();
        }
    };

    elements.browseRulesButton.addEventListener("click", browseRules);
    elements.configBrowseRulesButton.addEventListener("click", browseRules);
    elements.configOpenRulesButton.addEventListener("click", () => openPath(elements.rulesPath.value));
    elements.configUseDefaultRulesButton.addEventListener("click", async () => {
        if (!state.serverState) {
            return;
        }

        elements.rulesPath.value = state.serverState.paths.defaultRulesPath;
        queuePersistSettings();
        await refreshRulesPreview();
    });

    elements.openSettingsPathButton.addEventListener("click", () => openPath(state.serverState?.paths.settingsPath));
    elements.openHistoryPathButton.addEventListener("click", () => openPath(state.serverState?.paths.historyPath));
    elements.openLogsPathButton.addEventListener("click", () => openPath(state.serverState?.paths.logsDirectory));
    elements.openReportsPathButton.addEventListener("click", () => openPath(state.serverState?.paths.reportsDirectory));
    elements.refreshRulesPreviewButton.addEventListener("click", refreshRulesPreview);

    [
        elements.sourcePath,
        elements.destinationPath,
        elements.manifestPath,
        elements.rulesPath,
        elements.findModeSelect,
        elements.findDuplicateHandlingSelect,
        elements.findExtensions,
        elements.findIncludeSubfolders,
        elements.sortModeSelect,
        elements.sortDuplicateHandlingSelect,
        elements.sortOrganizationModeSelect,
        elements.sortIncludeSubfolders,
        elements.sortUseExtensionSubfolders,
        elements.includeRename,
        elements.metadataRenameModeSelect,
        elements.includeConvert,
        elements.convertSourceFormatSelect,
        elements.convertTargetFormatSelect,
        elements.convertQualitySelect,
        elements.mediaMusicEnabled,
        elements.mediaImagesEnabled,
        elements.mediaVideoEnabled,
        elements.writeTagManifest,
        elements.cleanEmptyFolders,
        elements.openDestinationWhenDone,
        elements.exportReport,
        elements.reportFormatSelect,
        elements.copySummaryToClipboard
    ].forEach(control => {
        const eventName = control.tagName === "INPUT" && control.type === "text" ? "input" : "change";
        control.addEventListener(eventName, async () => {
            if (control === elements.sourcePath) {
                handleSourcePathChanged();
            }
            if (control === elements.destinationPath) {
                handleDestinationPathChanged();
            }
            if (control === elements.rulesPath) {
                await refreshRulesPreview();
            }

            updateControlAvailability();
            saveUiState();
            queuePersistSettings();
        });
    });

    elements.presetButtons.forEach(button => button.addEventListener("click", () => {
        elements.findExtensions.value = button.dataset.preset || "";
        saveUiState();
    }));

    elements.sortPresetButtons.forEach(button => button.addEventListener("click", () => {
        setSelectValue(elements.sortOrganizationModeSelect, button.dataset.sortMode || "ByType");
        updateControlAvailability();
        queuePersistSettings();
    }));

    elements.previewButton.addEventListener("click", () => runOperation("/api/operations/preview", "Previewing"));
    elements.runButton.addEventListener("click", () => runOperation("/api/operations/run", "Running"));
    elements.undoLastButton.addEventListener("click", () => runUndo(null));
    elements.undoSelectedButton.addEventListener("click", () => runUndo(state.selectedHistoryRunId));
    elements.stopButton.addEventListener("click", stopCurrentOperation);
    elements.openOutputButton.addEventListener("click", () => openPath(elements.destinationPath.value));
    elements.openLogsButton.addEventListener("click", () => openPath(state.serverState?.paths.logsDirectory));
    elements.copyConsoleButton.addEventListener("click", copyConsoleOutput);
    elements.refreshHistoryButton.addEventListener("click", () => refreshState(false));
    elements.refreshLogsButton.addEventListener("click", () => refreshState(false));
    elements.openHistoryLogButton.addEventListener("click", () => {
        const entry = getSelectedHistoryEntry();
        if (entry) {
            openPath(entry.logFilePath);
        }
    });
    elements.openOperationLogButton.addEventListener("click", () => {
        const entry = getSelectedHistoryEntry();
        if (entry) {
            openPath(entry.operationLogPath);
        }
    });
    elements.openSelectedLogButton.addEventListener("click", () => {
        if (state.selectedLogPath) {
            openPath(state.selectedLogPath);
        }
    });
}

async function refreshState(hydrateForm) {
    const response = await fetch("/api/state");
    const payload = await parseResponse(response);
    state.serverState = payload;
    applyServerState(payload, hydrateForm || !state.hasHydrated);
}

function applyServerState(payload, hydrateForm) {
    if (hydrateForm) {
        hydrateFormFromSettings(payload.settings);
        applyLocalUiState();
        setActiveTab(tabOrder[payload.settings.selectedTabIndex] || "find", false);
        state.hasHydrated = true;
    }

    renderCapabilities(payload.capabilities);
    renderHistory(payload.history);
    renderLogs(payload.logs);
    renderConfig(payload);
    renderAbout(payload);
    elements.rulesPreview.textContent = payload.rulesPreview || "";
    updateControlAvailability();
}

function hydrateFormFromSettings(settings) {
    elements.sourcePath.value = settings.lastSourcePath || "";
    elements.destinationPath.value = settings.lastDestinationPath || "";
    elements.manifestPath.value = settings.lastTagManifestPath || "";
    elements.rulesPath.value = settings.lastRulesPath || "";
    setSelectValue(elements.sortModeSelect, settings.mode);
    setSelectValue(elements.sortDuplicateHandlingSelect, settings.duplicateHandling);
    setSelectValue(elements.sortOrganizationModeSelect, settings.organizationMode);
    setSelectValue(elements.metadataRenameModeSelect, settings.pdfRenameMode);
    elements.sortIncludeSubfolders.checked = !settings.topLevelOnly;
    elements.sortUseExtensionSubfolders.checked = settings.useExtensionSubfolders;
    elements.writeTagManifest.checked = settings.writeTagManifest;
    elements.cleanEmptyFolders.checked = settings.cleanEmptyFolders;
    elements.openDestinationWhenDone.checked = settings.openDestinationWhenDone;
    elements.exportReport.checked = settings.exportReport;
    setSelectValue(elements.reportFormatSelect, settings.reportExportFormat);
    elements.copySummaryToClipboard.checked = settings.copySummaryToClipboard;
    elements.includeConvert.checked = settings.includeConvert;
    setSelectValue(elements.convertSourceFormatSelect, settings.convertSourceFormat);
    setSelectValue(elements.convertTargetFormatSelect, settings.convertTargetFormat);
    setSelectValue(elements.convertQualitySelect, settings.convertQuality);
    elements.mediaMusicEnabled.checked = settings.mediaMusicUseMetadata;
    elements.mediaImagesEnabled.checked = settings.mediaImagesUseMetadata;
    elements.mediaVideoEnabled.checked = settings.mediaVideoUseMetadata;
    syncPathSuggestions();
}

function applyLocalUiState() {
    setSelectValue(elements.findModeSelect, state.ui.findMode);
    setSelectValue(elements.findDuplicateHandlingSelect, state.ui.findDuplicateHandling);
    elements.findIncludeSubfolders.checked = state.ui.findIncludeSubfolders;
    elements.findExtensions.value = state.ui.findExtensions;
    elements.includeRename.checked = state.ui.includeRename;
    setSelectValue(elements.metadataRenameModeSelect, state.ui.metadataRenameMode);
    elements.includeConvert.checked = state.ui.includeConvert;
    setSelectValue(elements.convertSourceFormatSelect, state.ui.convertSourceFormat);
    setSelectValue(elements.convertTargetFormatSelect, state.ui.convertTargetFormat);
    setSelectValue(elements.convertQualitySelect, state.ui.convertQuality);
    elements.mediaMusicEnabled.checked = state.ui.mediaMusicEnabled;
    elements.mediaImagesEnabled.checked = state.ui.mediaImagesEnabled;
    elements.mediaVideoEnabled.checked = state.ui.mediaVideoEnabled;
    state.activeWorkTab = state.ui.workTab === "sort" ? "sort" : "find";
}

function saveUiState() {
    state.ui = {
        workTab: state.activeWorkTab,
        findMode: elements.findModeSelect.value,
        findDuplicateHandling: elements.findDuplicateHandlingSelect.value,
        findIncludeSubfolders: elements.findIncludeSubfolders.checked,
        findExtensions: elements.findExtensions.value,
        includeRename: elements.includeRename.checked,
        metadataRenameMode: elements.metadataRenameModeSelect.value,
        includeConvert: elements.includeConvert.checked,
        convertSourceFormat: elements.convertSourceFormatSelect.value,
        convertTargetFormat: elements.convertTargetFormatSelect.value,
        convertQuality: elements.convertQualitySelect.value,
        mediaMusicEnabled: elements.mediaMusicEnabled.checked,
        mediaImagesEnabled: elements.mediaImagesEnabled.checked,
        mediaVideoEnabled: elements.mediaVideoEnabled.checked
    };

    localStorage.setItem(uiStateStorageKey, JSON.stringify(state.ui));
    updateTaskIndicator();
}

function renderCapabilities(capabilities) {
    if (!capabilities) {
        return;
    }

    if (elements.capShellMode) elements.capShellMode.textContent = capabilities.shellMode;
    if (elements.capImagePipeline) elements.capImagePipeline.textContent = capabilities.imagePipeline;
    if (elements.capAudioVideoPipeline) elements.capAudioVideoPipeline.textContent = capabilities.audioVideoPipeline;
    if (elements.capDocumentPipeline) elements.capDocumentPipeline.textContent = capabilities.documentPipeline;
}

function renderHistory(history) {
    elements.historyTableBody.innerHTML = "";
    const selectedRunId = state.selectedHistoryRunId && history.some(entry => entry.runId === state.selectedHistoryRunId)
        ? state.selectedHistoryRunId
        : history[0]?.runId || null;

    state.selectedHistoryRunId = selectedRunId;

    for (const entry of history) {
        const row = document.createElement("tr");
        row.dataset.runId = entry.runId;
        if (entry.runId === selectedRunId) {
            row.classList.add("is-selected");
        }

        row.innerHTML = `
            <td>${formatDate(entry.timestamp)}</td>
            <td>${escapeHtml(entry.runKind)}</td>
            <td>${escapeHtml(entry.mode)}</td>
            <td>${escapeHtml(entry.organizationMode)}</td>
            <td>${entry.executedCount}</td>
        `;

        row.addEventListener("click", () => {
            state.selectedHistoryRunId = entry.runId;
            renderHistory(history);
        });

        elements.historyTableBody.appendChild(row);
    }

    const current = getSelectedHistoryEntry();
    elements.historyDetail.textContent = current ? formatHistoryDetails(current) : "";
    elements.undoSelectedButton.disabled = !current || !current.canUndo || current.isUndone || Boolean(state.currentAbortController);
    elements.openHistoryLogButton.disabled = !current;
    elements.openOperationLogButton.disabled = !current;
}

function renderLogs(logs) {
    elements.logsTableBody.innerHTML = "";
    const selectedLogPath = state.selectedLogPath && logs.some(entry => entry.path === state.selectedLogPath)
        ? state.selectedLogPath
        : logs[0]?.path || null;

    state.selectedLogPath = selectedLogPath;

    for (const log of logs) {
        const row = document.createElement("tr");
        row.dataset.path = log.path;
        if (log.path === selectedLogPath) {
            row.classList.add("is-selected");
        }

        row.innerHTML = `
            <td>${formatDate(log.lastWriteTime)}</td>
            <td>${escapeHtml(log.fileName)}</td>
        `;

        row.addEventListener("click", async () => {
            state.selectedLogPath = log.path;
            renderLogs(logs);
            await loadLogContent(log.path);
        });

        elements.logsTableBody.appendChild(row);
    }

    if (selectedLogPath) {
        void loadLogContent(selectedLogPath);
    } else {
        elements.logDetail.textContent = "";
    }

    elements.openSelectedLogButton.disabled = !selectedLogPath;
}

async function loadLogContent(path) {
    try {
        const response = await fetch(`/api/file-content?path=${encodeURIComponent(path)}`);
        const payload = await parseResponse(response);
        elements.logDetail.textContent = payload.contents;
    } catch (error) {
        elements.logDetail.textContent = error.message;
    }
}

function renderConfig(payload) {
    elements.configRulesPath.textContent = elements.rulesPath.value || payload.paths.defaultRulesPath;
    elements.configSettingsPath.textContent = payload.paths.settingsPath;
    elements.configHistoryPath.textContent = payload.paths.historyPath;
    elements.configLogsPath.textContent = payload.paths.logsDirectory;
    elements.configReportsPath.textContent = payload.paths.reportsDirectory;
}

function renderAbout(payload) {
    elements.aboutVersion.textContent = payload.meta.version;
    elements.aboutHostUrl.textContent = payload.meta.hostUrl;
    elements.aboutAppDirectory.textContent = payload.paths.appDirectory;
    elements.aboutDataDirectory.textContent = payload.paths.dataDirectory;
}

function setActiveTab(tabId, persist) {
    state.activeTab = tabId;
    if (tabId === "find" || tabId === "sort") {
        state.activeWorkTab = tabId;
        saveUiState();
    } else {
        updateTaskIndicator();
    }

    elements.tabButtons.forEach(button => button.classList.toggle("is-active", button.dataset.tab === tabId));
    elements.tabPanels.forEach(panel => panel.classList.toggle("is-active", panel.dataset.panel === tabId));

    if (persist) {
        queuePersistSettings();
    }
}

function updateTaskIndicator() {
    elements.activeTaskLabel.textContent = state.activeWorkTab === "sort" ? "Sort" : "Find";
}

function updateControlAvailability() {
    const manifestEnabled = elements.writeTagManifest.checked;
    const extensionFoldersSupported = ["ByType", "CategoryYearMonth"].includes(elements.sortOrganizationModeSelect.value);
    elements.manifestPath.disabled = !manifestEnabled;
    elements.browseManifestButton.disabled = !manifestEnabled;
    elements.reportFormatSelect.disabled = !elements.exportReport.checked;
    elements.sortUseExtensionSubfolders.disabled = !extensionFoldersSupported;
    if (!extensionFoldersSupported) {
        elements.sortUseExtensionSubfolders.checked = false;
    }

    const convertEnabled = elements.includeConvert.checked;
    elements.convertSourceFormatSelect.disabled = !convertEnabled;
    elements.convertTargetFormatSelect.disabled = !convertEnabled;
    elements.convertQualitySelect.disabled = !convertEnabled;
    elements.stopButton.disabled = !state.currentAbortController;
}

function queuePersistSettings() {
    if (!state.hasHydrated) {
        return;
    }

    if (state.persistTimer) {
        clearTimeout(state.persistTimer);
    }

    state.persistTimer = setTimeout(() => {
        void persistSettings();
    }, 200);
}

async function persistSettings() {
    state.persistTimer = null;
    try {
        const response = await fetch("/api/settings", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(getPersistPayload())
        });

        const payload = await parseResponse(response);
        state.serverState = payload;
        renderCapabilities(payload.capabilities);
        renderConfig(payload);
        renderAbout(payload);
    } catch {
        // Background saves fail quietly.
    }
}

async function browsePath(endpoint, initialPath) {
    const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ initialPath })
    });
    const payload = await parseResponse(response);
    return payload.path;
}

async function runOperation(endpoint, verb) {
    if (state.currentAbortController) {
        return;
    }

    const validationError = validateOperationRequest();
    if (validationError) {
        elements.outputConsole.textContent = validationError;
        elements.resultLabel.textContent = "Blocked";
        return;
    }

    const controller = new AbortController();
    const payload = getOperationPayload();
    state.currentAbortController = controller;
    updateControlAvailability();
    setBusyState(`${verb} ${state.activeWorkTab}...`);

    try {
        const response = await fetch(endpoint, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
            signal: controller.signal
        });

        const resultPayload = await parseResponse(response);
        state.serverState = resultPayload.state;
        applyServerState(resultPayload.state, false);
        elements.outputConsole.textContent = resultPayload.result.detailedReport;
        elements.resultLabel.textContent = formatOperationResult(resultPayload.result);
        setProgressDone(calculateCompletion(resultPayload.result));

        if (payload.copySummaryToClipboard && navigator.clipboard?.writeText) {
            await navigator.clipboard.writeText(resultPayload.result.summaryText);
        }

        if (resultPayload.exportedReportPath) {
            elements.outputConsole.textContent += `\n\nReport exported: ${resultPayload.exportedReportPath}`;
        }
    } catch (error) {
        if (error.name === "AbortError") {
            elements.outputConsole.textContent += "\n\nOperation canceled locally.";
            elements.resultLabel.textContent = "Canceled";
        } else {
            elements.outputConsole.textContent = error.message;
            elements.resultLabel.textContent = "Error";
        }
        setProgressDone(0);
    } finally {
        state.currentAbortController = null;
        updateControlAvailability();
        elements.progressTrack.classList.remove("is-busy");
    }
}

async function runUndo(runId) {
    if (state.currentAbortController) {
        return;
    }

    const controller = new AbortController();
    state.currentAbortController = controller;
    updateControlAvailability();
    setBusyState(runId ? "Undoing selected run..." : "Undoing last run...");

    try {
        const response = await fetch("/api/operations/undo", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ runId }),
            signal: controller.signal
        });

        const payload = await parseResponse(response);
        state.serverState = payload.state;
        applyServerState(payload.state, false);
        elements.outputConsole.textContent = payload.result.detailedReport;
        elements.resultLabel.textContent = payload.result.summaryText;
        setProgressDone(100);
    } catch (error) {
        if (error.name === "AbortError") {
            elements.outputConsole.textContent += "\n\nUndo canceled locally.";
            elements.resultLabel.textContent = "Undo canceled";
        } else {
            elements.outputConsole.textContent = error.message;
            elements.resultLabel.textContent = "Undo error";
        }
        setProgressDone(0);
    } finally {
        state.currentAbortController = null;
        updateControlAvailability();
        elements.progressTrack.classList.remove("is-busy");
    }
}

function stopCurrentOperation() {
    if (state.currentAbortController) {
        state.currentAbortController.abort();
        elements.progressLabel.textContent = "Canceling";
    }
}

async function refreshRulesPreview() {
    try {
        const response = await fetch(`/api/rules-content?path=${encodeURIComponent(elements.rulesPath.value || "")}`);
        const payload = await parseResponse(response);
        elements.rulesPreview.textContent = payload.contents;
        elements.configRulesPath.textContent = payload.path;
    } catch (error) {
        elements.rulesPreview.textContent = error.message;
    }
}

async function openPath(path) {
    if (!path) {
        return;
    }

    try {
        const response = await fetch("/api/open-path", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ path })
        });
        await parseResponse(response);
    } catch (error) {
        elements.outputConsole.textContent = error.message;
    }
}

async function copyConsoleOutput() {
    if (!navigator.clipboard?.writeText) {
        return;
    }

    try {
        await navigator.clipboard.writeText(elements.outputConsole.textContent);
    } catch {
        // Ignore clipboard failures.
    }
}

function getPersistPayload() {
    return buildPayload();
}

function getOperationPayload() {
    return buildPayload();
}

function buildPayload() {
    const useFindProfile = state.activeWorkTab !== "sort";
    return {
        sourcePath: normalizeInput(elements.sourcePath.value),
        destinationPath: normalizeInput(elements.destinationPath.value),
        tagManifestPath: normalizeInput(elements.manifestPath.value),
        rulesPath: normalizeInput(elements.rulesPath.value),
        mode: useFindProfile ? elements.findModeSelect.value : elements.sortModeSelect.value,
        duplicateHandling: useFindProfile ? elements.findDuplicateHandlingSelect.value : elements.sortDuplicateHandlingSelect.value,
        organizationMode: useFindProfile ? "ByExtension" : elements.sortOrganizationModeSelect.value,
        pdfRenameMode: elements.includeRename.checked ? elements.metadataRenameModeSelect.value : "Disabled",
        topLevelOnly: useFindProfile ? !elements.findIncludeSubfolders.checked : !elements.sortIncludeSubfolders.checked,
        useExtensionSubfolders: useFindProfile ? false : elements.sortUseExtensionSubfolders.checked,
        writeTagManifest: elements.writeTagManifest.checked,
        cleanEmptyFolders: elements.cleanEmptyFolders.checked,
        openDestinationWhenDone: elements.openDestinationWhenDone.checked,
        exportReport: elements.exportReport.checked,
        reportExportFormat: elements.reportFormatSelect.value,
        copySummaryToClipboard: elements.copySummaryToClipboard.checked,
        includeConvert: elements.includeConvert.checked,
        convertSourceFormat: elements.convertSourceFormatSelect.value,
        convertTargetFormat: elements.convertTargetFormatSelect.value,
        convertQuality: elements.convertQualitySelect.value,
        mediaMusicUseMetadata: elements.mediaMusicEnabled.checked,
        mediaImagesUseMetadata: elements.mediaImagesEnabled.checked,
        mediaVideoUseMetadata: elements.mediaVideoEnabled.checked,
        selectedTabIndex: Math.max(0, tabOrder.indexOf(state.activeTab)),
        includeExtensions: useFindProfile ? normalizeInput(elements.findExtensions.value) : null
    };
}

function validateOperationRequest() {
    const source = normalizeInput(elements.sourcePath.value);
    const destination = normalizeInput(elements.destinationPath.value);

    if (!source) {
        return "Choose a source folder before running.";
    }

    if (destination && sameWindowsPath(source, destination)) {
        return "Source and destination cannot be the same folder.";
    }

    if (state.activeWorkTab === "find" && parseExtensionList(elements.findExtensions.value).length === 0) {
        return "Enter at least one file type on the Find tab before running.";
    }

    return null;
}

function setBusyState(label) {
    elements.progressLabel.textContent = label;
    elements.progressTrack.classList.add("is-busy");
    elements.progressFill.style.width = "38%";
    elements.previewButton.disabled = true;
    elements.runButton.disabled = true;
    elements.undoLastButton.disabled = true;
    elements.undoSelectedButton.disabled = true;
}

function setProgressDone(percent) {
    elements.progressTrack.classList.remove("is-busy");
    elements.progressFill.style.width = `${Math.max(0, Math.min(100, percent))}%`;
    elements.progressLabel.textContent = percent === 0 ? "Idle" : "Completed";
    elements.previewButton.disabled = false;
    elements.runButton.disabled = false;
    elements.undoLastButton.disabled = false;
}

function handleSourcePathChanged() {
    const previousSuggestedDestination = state.suggestedDestinationPath;
    const previousSuggestedManifest = state.suggestedManifestPath;
    const source = normalizeWindowsPath(elements.sourcePath.value);

    state.suggestedDestinationPath = source ? joinWindowsPath(source, "_organized") : null;

    const shouldUpdateDestination = !elements.destinationPath.value.trim() || sameWindowsPath(elements.destinationPath.value, previousSuggestedDestination);
    if (shouldUpdateDestination && state.suggestedDestinationPath) {
        elements.destinationPath.value = state.suggestedDestinationPath;
    }

    const destination = normalizeWindowsPath(elements.destinationPath.value) || state.suggestedDestinationPath;
    state.suggestedManifestPath = destination ? joinWindowsPath(destination, "file-tags.json") : null;

    const shouldUpdateManifest = !elements.manifestPath.value.trim() || sameWindowsPath(elements.manifestPath.value, previousSuggestedManifest);
    if (shouldUpdateManifest && state.suggestedManifestPath) {
        elements.manifestPath.value = state.suggestedManifestPath;
    }
}

function handleDestinationPathChanged() {
    const previousSuggestedManifest = state.suggestedManifestPath;
    const destination = normalizeWindowsPath(elements.destinationPath.value) || state.suggestedDestinationPath;
    state.suggestedManifestPath = destination ? joinWindowsPath(destination, "file-tags.json") : null;

    const shouldUpdateManifest = !elements.manifestPath.value.trim() || sameWindowsPath(elements.manifestPath.value, previousSuggestedManifest);
    if (shouldUpdateManifest && state.suggestedManifestPath) {
        elements.manifestPath.value = state.suggestedManifestPath;
    }
}

function syncPathSuggestions() {
    const source = normalizeWindowsPath(elements.sourcePath.value);
    state.suggestedDestinationPath = source ? joinWindowsPath(source, "_organized") : null;
    const destination = normalizeWindowsPath(elements.destinationPath.value) || state.suggestedDestinationPath;
    state.suggestedManifestPath = destination ? joinWindowsPath(destination, "file-tags.json") : null;
}

function getSelectedHistoryEntry() {
    return state.serverState?.history?.find(entry => entry.runId === state.selectedHistoryRunId) || null;
}

function formatHistoryDetails(entry) {
    return [
        `Run ID: ${entry.runId}`,
        `Type: ${entry.runKind}`,
        `Time: ${formatDate(entry.timestamp)}`,
        `Source: ${entry.sourceRoot}`,
        `Destination: ${entry.destinationRoot}`,
        `Mode: ${entry.mode}`,
        `Organize By: ${entry.organizationMode}`,
        `Duplicate Handling: ${entry.duplicateHandling}`,
        `Metadata Rename Mode: ${entry.pdfRenameMode}`,
        `Preview Run: ${entry.whatIf}`,
        `Planned: ${entry.plannedCount}`,
        `Executed: ${entry.executedCount}`,
        `Skipped: ${entry.skippedCount}`,
        `Conversions: ${entry.conversionCount || 0}`,
        `Rules: ${entry.rulesPath}`,
        entry.tagManifestPath ? `Tag Manifest: ${entry.tagManifestPath}` : null,
        `Log: ${entry.logFilePath}`,
        `Operation Log: ${entry.operationLogPath}`,
        `Undo Available: ${entry.canUndo && !entry.isUndone}`,
        `Undone: ${entry.isUndone}`,
        entry.undoneByRunId ? `Undone By: ${entry.undoneByRunId}` : null,
        "",
        entry.summaryText || ""
    ].filter(Boolean).join("\n");
}

function formatOperationResult(result) {
    const changed = result.operations.filter(item => item.status === "Executed").length;
    const skipped = result.operations.filter(item => item.status === "Skipped").length;
    const failed = result.operations.filter(item => item.status === "Failed").length;
    const conversions = result.conversions?.filter(item => item.status === "Executed").length || 0;
    return result.whatIf
        ? `Preview â€¢ ${changed} planned â€¢ ${skipped} skipped â€¢ ${failed} failed â€¢ ${conversions} conversions staged`
        : `Run â€¢ ${changed} changed â€¢ ${skipped} skipped â€¢ ${failed} failed â€¢ ${conversions} conversions`;
}

function calculateCompletion(result) {
    const planned = result.operations.length;
    if (planned === 0) {
        return 100;
    }

    const completed = result.operations.filter(item => ["Executed", "Skipped", "Failed"].includes(item.status)).length;
    return Math.round((completed / planned) * 100);
}

function applyTheme(mode) {
    state.themeMode = mode;
    localStorage.setItem(themeStorageKey, mode);
    const effectiveTheme = mode === "system" ? (colorScheme.matches ? "dark" : "light") : mode;
    document.body.dataset.themeMode = mode;
    document.body.dataset.effectiveTheme = effectiveTheme;
    elements.themeSelect.value = mode;
}

function loadUiState() {
    try {
        return { ...defaultUiState, ...JSON.parse(localStorage.getItem(uiStateStorageKey) || "{}") };
    } catch {
        return { ...defaultUiState };
    }
}

function parseExtensionList(value) {
    return (value || "")
        .split(/[\s,;]+/)
        .map(part => part.trim())
        .filter(Boolean)
        .map(part => part.startsWith(".") ? part.toLowerCase() : `.${part.toLowerCase()}`)
        .filter((part, index, array) => array.indexOf(part) === index);
}

function setSelectValue(select, value) {
    if (!select || value == null) {
        return;
    }

    const stringValue = String(value);
    if ([...select.options].some(option => option.value === stringValue)) {
        select.value = stringValue;
    }
}

function normalizeInput(value) {
    const trimmed = value?.trim();
    return trimmed ? trimmed : null;
}

function normalizeWindowsPath(path) {
    const trimmed = (path || "").trim();
    if (!trimmed) {
        return null;
    }

    return trimmed.replace(/[\\/]+$/, "");
}

function sameWindowsPath(left, right) {
    const normalizedLeft = normalizeWindowsPath(left);
    const normalizedRight = normalizeWindowsPath(right);
    return Boolean(normalizedLeft && normalizedRight && normalizedLeft.toLowerCase() === normalizedRight.toLowerCase());
}

function joinWindowsPath(root, leaf) {
    const normalizedRoot = normalizeWindowsPath(root) || "";
    return `${normalizedRoot}\\${leaf}`;
}

function startHeartbeat() {
    void fetch("/api/heartbeat", { method: "POST", keepalive: true });
    window.setInterval(() => {
        void fetch("/api/heartbeat", { method: "POST", keepalive: true });
    }, 5000);
}

function formatDate(value) {
    return new Date(value).toLocaleString();
}

async function parseResponse(response) {
    const text = await response.text();
    const payload = text ? JSON.parse(text) : {};
    if (!response.ok) {
        throw new Error(payload.message || `Request failed with status ${response.status}.`);
    }
    return payload;
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;");
}



