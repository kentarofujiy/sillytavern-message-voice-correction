import { eventSource, event_types, saveSettingsDebounced } from '../../../../script.js';
import { extension_settings, getContext, renderExtensionTemplateAsync } from '../../../extensions.js';
import { Popup, POPUP_RESULT, POPUP_TYPE } from '../../../popup.js';

const MODULE_NAME = 'message-voice-correction';
const TEMPLATE_MODULE = 'third-party/message-voice-correction';
const SETTINGS_KEY = 'messageVoiceCorrection';
const SETTINGS_VERSION = 1;

const DEFAULT_SETTINGS = {
    settingsVersion: SETTINGS_VERSION,
    enabled: true,
    responseLength: 220,
    showPromptPreview: false,
    participantInstructions: {},
};

const runtimeState = {
    initialized: false,
    selectedMessageId: null,
    pendingMessages: new Set(),
};

function getSettings() {
    if (!extension_settings[SETTINGS_KEY] || typeof extension_settings[SETTINGS_KEY] !== 'object') {
        extension_settings[SETTINGS_KEY] = structuredClone(DEFAULT_SETTINGS);
    }

    const settings = extension_settings[SETTINGS_KEY];
    Object.entries(DEFAULT_SETTINGS).forEach(([key, value]) => {
        if (settings[key] === undefined) {
            settings[key] = value;
        }
    });

    settings.settingsVersion = SETTINGS_VERSION;
    settings.enabled = Boolean(settings.enabled);
    settings.responseLength = clampNumber(settings.responseLength, 80, 6000, DEFAULT_SETTINGS.responseLength);
    settings.showPromptPreview = Boolean(settings.showPromptPreview);
    settings.participantInstructions = normalizeInstructionMap(settings.participantInstructions);

    return settings;
}

function normalizeInstructionMap(value) {
    const map = {};

    if (!value || typeof value !== 'object') {
        return map;
    }

    for (const [key, entry] of Object.entries(value)) {
        if (!key) {
            continue;
        }

        map[key] = String(entry ?? '').trim();
    }

    return map;
}

function clampNumber(value, min, max, fallback) {
    const number = Number(value);
    if (!Number.isFinite(number)) {
        return fallback;
    }

    return Math.min(max, Math.max(min, Math.round(number)));
}

function showToast(level, message) {
    if (globalThis.toastr?.[level]) {
        globalThis.toastr[level](message, 'Message Voice Correction');
    }
}

function setStatus(text) {
    const element = document.getElementById('message_voice_correction_status');
    if (element) {
        element.textContent = text;
    }
}

function isCharacterMessage(message) {
    return Boolean(message && !message.is_user && !message.is_system && typeof message.mes === 'string');
}

function getMessageElement(messageId) {
    return document.querySelector(`#chat .mes[mesid="${Number(messageId)}"]`);
}

function getCurrentGroup(context) {
    if (!context.groupId) {
        return null;
    }

    return context.groups.find((group) => String(group.id) === String(context.groupId)) ?? null;
}

function getParticipants(context) {
    const group = getCurrentGroup(context);

    if (!group) {
        const character = context.characters[context.characterId];
        if (!character) {
            return [];
        }

        return [{
            key: `solo:${context.chatId || character.avatar || character.name}`,
            name: character.name,
            characterId: context.characterId,
            avatar: character.avatar,
        }];
    }

    return group.members.map((avatar, index) => {
        const characterId = context.characters.findIndex((item) => item.avatar === avatar);
        const character = characterId >= 0 ? context.characters[characterId] : null;
        const fallbackName = character?.name || `Member ${index + 1}`;

        return {
            key: `group:${group.id}:${avatar || fallbackName}`,
            name: fallbackName,
            characterId: characterId >= 0 ? characterId : null,
            avatar,
        };
    });
}

function getParticipantForMessage(context, message) {
    const participants = getParticipants(context);

    const byName = participants.find((participant) => participant.name === message.name);
    if (byName) {
        return byName;
    }

    if (!context.groupId && participants.length > 0) {
        return participants[0];
    }

    return {
        key: `group:${context.groupId || 'unknown'}:${message.name || 'unknown'}`,
        name: String(message.name || 'Unknown'),
        characterId: null,
        avatar: null,
    };
}

function summarizeCharacterField(value, maxLength = 800) {
    return String(value ?? '').replace(/\s+/g, ' ').trim().slice(0, maxLength);
}

function buildPrompt({ messageText, participantInstruction, additionalInstruction, characterFields, speakerName }) {
    return [
        'You are correcting one SillyTavern roleplay message.',
        'Return only the corrected final message text.',
        'Do not include analysis, labels, markdown, or explanations.',
        'Preserve intent while improving coherence, reducing repetition, and keeping the character voice consistent.',
        'Do not add information unsupported by the provided context.',
        '',
        `Speaker: ${speakerName}`,
        characterFields.description && `Character description: ${characterFields.description}`,
        characterFields.personality && `Character personality: ${characterFields.personality}`,
        characterFields.scenario && `Scenario: ${characterFields.scenario}`,
        characterFields.mesExamples && `Dialogue examples: ${characterFields.mesExamples}`,
        participantInstruction && `Per-participant instruction: ${participantInstruction}`,
        additionalInstruction && `Additional user instruction: ${additionalInstruction}`,
        '',
        'Message to correct:',
        `<message>${messageText}</message>`,
        '',
        'Corrected message:',
    ].filter(Boolean).join('\n');
}

function cleanupOutput(text, fallback) {
    const normalized = String(text ?? '').trim();
    if (!normalized) {
        return fallback;
    }

    return normalized
        .replace(/^```(?:[\w-]+)?\n?/g, '')
        .replace(/\n?```$/g, '')
        .trim() || fallback;
}

function clearSelectedMessageHighlight() {
    document.querySelectorAll('#chat .mes.message-voice-correction-selected').forEach((element) => {
        element.classList.remove('message-voice-correction-selected');
    });
}

function updateActionVisibility() {
    document.querySelectorAll('#chat .mes .message-voice-correction-action').forEach((element) => {
        const action = element;
        if (!(action instanceof HTMLElement)) {
            return;
        }

        const messageId = Number(action.dataset.messageId);
        const shouldShow = runtimeState.selectedMessageId !== null && Number.isFinite(messageId) && messageId === runtimeState.selectedMessageId;
        action.classList.toggle('message-voice-correction-hidden', !shouldShow);
    });

    clearSelectedMessageHighlight();

    if (runtimeState.selectedMessageId !== null) {
        getMessageElement(runtimeState.selectedMessageId)?.classList.add('message-voice-correction-selected');
    }
}

function upsertMessageAction(messageId) {
    const context = getContext();
    const settings = getSettings();
    const message = context.chat[Number(messageId)];
    const messageElement = getMessageElement(messageId);

    if (!(messageElement instanceof HTMLElement)) {
        return;
    }

    messageElement.querySelector('.message-voice-correction-action')?.remove();

    if (!settings.enabled || !isCharacterMessage(message)) {
        return;
    }

    const buttons = messageElement.querySelector('.extraMesButtons') || messageElement.querySelector('.mes_buttons');
    if (!(buttons instanceof HTMLElement)) {
        return;
    }

    const action = document.createElement('div');
    action.className = 'mes_button message-voice-correction-action message-voice-correction-hidden fa-solid fa-wand-magic-sparkles';
    action.dataset.messageId = String(messageId);
    action.title = 'Correct Message';
    buttons.prepend(action);
}

function refreshMessageActions() {
    const context = getContext();
    context.chat.forEach((_, messageId) => upsertMessageAction(messageId));
    updateActionVisibility();
}

function resetSelection() {
    runtimeState.selectedMessageId = null;
    updateActionVisibility();
}

function renderParticipants() {
    const container = document.getElementById('message_voice_correction_participants');
    if (!(container instanceof HTMLElement)) {
        return;
    }

    const context = getContext();
    const settings = getSettings();
    const participants = getParticipants(context);
    container.innerHTML = '';

    if (participants.length === 0) {
        const placeholder = document.createElement('small');
        placeholder.textContent = 'No active participants found.';
        container.appendChild(placeholder);
        return;
    }

    participants.forEach((participant) => {
        const field = document.createElement('label');
        field.className = 'message-voice-correction-participant';

        const title = document.createElement('span');
        title.textContent = participant.name;

        const area = document.createElement('textarea');
        area.className = 'text_pole';
        area.rows = 4;
        area.placeholder = 'Add extra voice/content instructions for this participant.';
        area.value = settings.participantInstructions[participant.key] || '';
        area.addEventListener('input', () => {
            settings.participantInstructions[participant.key] = area.value.trim();
            saveSettingsDebounced();
        });

        field.append(title, area);
        container.appendChild(field);
    });
}

function bindSettings() {
    const settings = getSettings();
    const enabled = document.getElementById('message_voice_correction_enabled');
    const responseLength = document.getElementById('message_voice_correction_response_length');
    const responseLengthValue = document.getElementById('message_voice_correction_response_length_value');
    const previewToggle = document.getElementById('message_voice_correction_show_prompt_preview');

    if (enabled instanceof HTMLInputElement) {
        enabled.checked = settings.enabled;
        enabled.addEventListener('input', () => {
            settings.enabled = enabled.checked;
            saveSettingsDebounced();
            refreshMessageActions();
        });
    }

    if (responseLength instanceof HTMLInputElement && responseLengthValue instanceof HTMLElement) {
        responseLength.value = String(settings.responseLength);
        responseLengthValue.textContent = `${settings.responseLength} tokens`;
        responseLength.addEventListener('input', () => {
            settings.responseLength = clampNumber(responseLength.value, 80, 6000, DEFAULT_SETTINGS.responseLength);
            responseLengthValue.textContent = `${settings.responseLength} tokens`;
            saveSettingsDebounced();
        });
    }

    if (previewToggle instanceof HTMLInputElement) {
        previewToggle.checked = settings.showPromptPreview;
        previewToggle.addEventListener('input', () => {
            settings.showPromptPreview = previewToggle.checked;
            saveSettingsDebounced();
        });
    }
}

async function renderSettings() {
    if (!document.getElementById('message_voice_correction_container')) {
        const html = await renderExtensionTemplateAsync(TEMPLATE_MODULE, 'settings');
        const container = document.getElementById('extensions_settings2') || document.getElementById('extensions_settings');

        if (!container) {
            return;
        }

        container.insertAdjacentHTML('beforeend', html);
        bindSettings();
    }

    renderParticipants();
    setStatus('Idle');
}

function handleMessageRendered(messageId) {
    upsertMessageAction(messageId);
    updateActionVisibility();
}

function isInteractiveTarget(target) {
    if (!(target instanceof HTMLElement)) {
        return true;
    }

    return Boolean(target.closest('button, a, input, textarea, select, .mes_buttons, .mes_edit_buttons, .extraMesButtons, .swipe_left, .swipe_right'));
}

async function confirmCorrection({ speakerName, promptPreview }) {
    const settings = getSettings();
    const wrapper = document.createElement('div');
    wrapper.className = 'message-voice-correction-dialog';

    const description = document.createElement('div');
    description.textContent = `Submit this ${speakerName} message for voice and content correction?`;
    wrapper.appendChild(description);

    if (settings.showPromptPreview) {
        const preview = document.createElement('pre');
        preview.className = 'message-voice-correction-dialog-preview';
        preview.textContent = promptPreview;
        wrapper.appendChild(preview);
    }

    const popup = new Popup(wrapper, POPUP_TYPE.CONFIRM, '', {
        wide: true,
        okButton: 'Correct Message',
        cancelButton: 'Cancel',
        customInputs: [{
            id: 'message_voice_correction_additional',
            label: 'Additional instructions (optional)',
            type: 'textarea',
            rows: 5,
            defaultState: '',
        }],
    });

    const result = await popup.show();
    if (result !== POPUP_RESULT.AFFIRMATIVE) {
        return null;
    }

    return String(popup.inputResults?.get('message_voice_correction_additional') ?? '').trim();
}

async function correctMessage(messageId) {
    const id = Number(messageId);
    const context = getContext();
    const settings = getSettings();
    const message = context.chat[id];

    if (!isCharacterMessage(message)) {
        showToast('warning', 'Only character messages can be corrected.');
        return;
    }

    if (runtimeState.pendingMessages.has(id)) {
        return;
    }

    const participant = getParticipantForMessage(context, message);
    const participantInstruction = settings.participantInstructions[participant.key] || '';

    const characterFields = participant.characterId !== null
        ? context.getCharacterCardFields({ chid: participant.characterId })
        : { description: '', personality: '', scenario: '', mesExamples: '' };

    const initialPrompt = buildPrompt({
        messageText: message.mes,
        participantInstruction,
        additionalInstruction: '',
        characterFields: {
            description: summarizeCharacterField(characterFields.description),
            personality: summarizeCharacterField(characterFields.personality),
            scenario: summarizeCharacterField(characterFields.scenario),
            mesExamples: summarizeCharacterField(characterFields.mesExamples, 1200),
        },
        speakerName: participant.name,
    });

    const additionalInstruction = await confirmCorrection({
        speakerName: participant.name,
        promptPreview: initialPrompt,
    });

    if (additionalInstruction === null) {
        return;
    }

    const finalPrompt = buildPrompt({
        messageText: message.mes,
        participantInstruction,
        additionalInstruction,
        characterFields: {
            description: summarizeCharacterField(characterFields.description),
            personality: summarizeCharacterField(characterFields.personality),
            scenario: summarizeCharacterField(characterFields.scenario),
            mesExamples: summarizeCharacterField(characterFields.mesExamples, 1200),
        },
        speakerName: participant.name,
    });

    runtimeState.pendingMessages.add(id);
    setStatus(`Correcting message for ${participant.name}`);

    try {
        const corrected = await context.generateQuietPrompt({
            quietPrompt: finalPrompt,
            responseLength: settings.responseLength,
            skipWIAN: true,
            removeReasoning: true,
        });

        const cleaned = cleanupOutput(corrected, message.mes);
        if (!cleaned || cleaned === message.mes) {
            showToast('info', 'No changes were suggested for this message.');
            setStatus('No changes applied');
            return;
        }

        message.mes = cleaned;
        context.updateMessageBlock(id, message);
        await context.saveChat();
        showToast('success', 'Message corrected.');
        setStatus(`Corrected message for ${participant.name}`);
    } catch (error) {
        console.error('[Message Voice Correction] Correction failed', error);
        showToast('error', error instanceof Error ? error.message : 'Correction failed.');
        setStatus('Correction failed');
    } finally {
        runtimeState.pendingMessages.delete(id);
    }
}

function handleDocumentClick(event) {
    const target = event.target instanceof HTMLElement ? event.target : null;

    const action = target?.closest('.message-voice-correction-action');
    if (action instanceof HTMLElement) {
        event.preventDefault();
        event.stopPropagation();
        void correctMessage(action.dataset.messageId);
        return;
    }

    if (!target || isInteractiveTarget(target) || window.getSelection()?.toString()) {
        return;
    }

    const messageElement = target.closest('#chat .mes[mesid]');
    if (!(messageElement instanceof HTMLElement)) {
        return;
    }

    const messageId = Number(messageElement.getAttribute('mesid'));
    const context = getContext();
    const message = context.chat[messageId];

    if (!isCharacterMessage(message)) {
        resetSelection();
        return;
    }

    runtimeState.selectedMessageId = messageId;
    updateActionVisibility();
}

function handleChatContextChanged() {
    resetSelection();
    renderParticipants();
    refreshMessageActions();
}

async function initialize() {
    if (runtimeState.initialized) {
        await renderSettings();
        return;
    }

    runtimeState.initialized = true;
    getSettings();
    document.addEventListener('click', handleDocumentClick);

    eventSource.on(event_types.CHARACTER_MESSAGE_RENDERED, handleMessageRendered);
    eventSource.on(event_types.CHAT_CHANGED, handleChatContextChanged);
    eventSource.on(event_types.CHAT_LOADED, handleChatContextChanged);
    eventSource.on(event_types.MORE_MESSAGES_LOADED, refreshMessageActions);

    await renderSettings();
    refreshMessageActions();
}

eventSource.on(event_types.APP_READY, () => void initialize());
eventSource.on(event_types.EXTENSION_SETTINGS_LOADED, () => void renderSettings());
void initialize();
