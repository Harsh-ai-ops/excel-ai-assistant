import * as React from 'react';
import { useState, useEffect } from 'react';
import {
    Modal,
    Stack,
    TextField,
    PrimaryButton,
    DefaultButton,
    Dropdown,
    IDropdownOption,
    Label,
    Link,
    MessageBar,
    MessageBarType,
    IconButton,
    ComboBox,
    IComboBoxOption,
    IComboBox,
    Spinner,
    SpinnerSize,
} from '@fluentui/react';
import { StorageService, LLMProvider } from '../services/storageService';
import { OPENROUTER_FREE_MODELS, LLMService } from '../services/llmService';

interface SettingsPanelProps {
    isOpen: boolean;
    onClose: () => void;
}

const providerOptions: IDropdownOption[] = [
    { key: 'openrouter', text: 'OpenRouter (FREE models)' },
    { key: 'gemini', text: 'Google Gemini (FREE)' },
    { key: 'huggingface', text: 'HuggingFace (FREE)' },
];

const SettingsPanel: React.FC<SettingsPanelProps> = ({ isOpen, onClose }) => {
    const [apiKey, setApiKey] = useState('');
    const [provider, setProvider] = useState<LLMProvider>('openrouter');
    const [model, setModel] = useState(OPENROUTER_FREE_MODELS[0].id);
    const [saved, setSaved] = useState(false);
    const [availableModels, setAvailableModels] = useState<{ id: string; name: string }[]>(OPENROUTER_FREE_MODELS);
    const [isLoadingModels, setIsLoadingModels] = useState(false);

    // Load settings on mount
    useEffect(() => {
        if (isOpen) {
            const settings = StorageService.getSettings();
            setApiKey(settings.apiKey || '');
            setProvider(settings.provider);
            setModel(settings.model);
            setSaved(false);

            // Initial fetch if we have models
            if (settings.provider === 'openrouter') {
                fetchModels();
            }
        }
    }, [isOpen]);

    const fetchModels = async () => {
        setIsLoadingModels(true);
        try {
            const models = await LLMService.getOpenRouterModels();
            setAvailableModels(models);
        } catch (error) {
            console.error('Error fetching models:', error);
        } finally {
            setIsLoadingModels(false);
        }
    };

    const handleSave = () => {
        if (apiKey.trim()) {
            StorageService.saveApiKey(apiKey.trim());
        }
        StorageService.saveProvider(provider);
        StorageService.saveModel(model);
        setSaved(true);
        setTimeout(() => {
            onClose();
        }, 1000);
    };

    const handleClearKey = () => {
        StorageService.clearApiKey();
        setApiKey('');
        setSaved(false);
    };

    const getProviderInfo = () => {
        switch (provider) {
            case 'openrouter':
                return {
                    placeholder: 'sk-or-v1-...',
                    link: 'https://openrouter.ai/keys',
                    linkText: 'Get FREE API key from OpenRouter →',
                    description: '✅ Completely FREE forever (no credit card needed)\n✅ Use models like: Llama 3.1, Mistral 7B, Gemma\n✅ Rate limited but no billing',
                };
            case 'gemini':
                return {
                    placeholder: 'AIza...',
                    link: 'https://makersuite.google.com/app/apikey',
                    linkText: 'Get FREE API key from Google →',
                    description: '✅ 60 requests/minute FREE\n✅ Very capable model\n✅ No credit card required',
                };
            case 'huggingface':
                return {
                    placeholder: 'hf_...',
                    link: 'https://huggingface.co/settings/tokens',
                    linkText: 'Get FREE token from HuggingFace →',
                    description: '✅ Completely free inference API\n✅ Many models available\n✅ No credit card required',
                };
        }
    };

    const providerInfo = getProviderInfo();

    const modelOptions: IComboBoxOption[] = availableModels.map((m) => ({
        key: m.id,
        text: m.name,
    }));

    return (
        <Modal
            isOpen={isOpen}
            onDismiss={onClose}
            isBlocking={false}
            containerClassName="settings-modal"
        >
            <Stack className="settings-panel" tokens={{ childrenGap: 15 }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <h2 style={{ margin: 0 }}>⚙️ Settings</h2>
                    <IconButton
                        iconProps={{ iconName: 'Cancel' }}
                        onClick={onClose}
                        ariaLabel="Close"
                    />
                </Stack>

                {saved && (
                    <MessageBar messageBarType={MessageBarType.success}>
                        ✅ Settings saved successfully!
                    </MessageBar>
                )}

                <Dropdown
                    label="LLM Provider (all FREE)"
                    selectedKey={provider}
                    options={providerOptions}
                    onChange={(_, option) => setProvider(option?.key as LLMProvider)}
                />

                {provider === 'openrouter' && (
                    <Stack tokens={{ childrenGap: 5 }}>
                        <div style={{ display: 'flex', alignItems: 'flex-end', gap: 10 }}>
                            <ComboBox
                                label="Model"
                                selectedKey={model}
                                options={modelOptions}
                                onChange={(_: React.FormEvent<IComboBox>, option?: IComboBoxOption, index?: number, value?: string) => {
                                    if (option) {
                                        setModel(option.key as string);
                                    } else if (value) {
                                        setModel(value);
                                    }
                                }}
                                allowFreeform
                                autoComplete="on"
                                styles={{ root: { flex: 1 } }}
                                placeholder="Select or type a model ID..."
                            />
                            <IconButton
                                iconProps={{ iconName: 'Refresh' }}
                                title="Refresh Models"
                                onClick={fetchModels}
                                disabled={isLoadingModels}
                            />
                        </div>
                        {isLoadingModels && <Spinner size={SpinnerSize.small} label="Loading models..." />}
                    </Stack>
                )}

                <Stack tokens={{ childrenGap: 8 }}>
                    <Label>API Key</Label>
                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <TextField
                            type="password"
                            value={apiKey}
                            onChange={(_, val) => setApiKey(val || '')}
                            placeholder={providerInfo.placeholder}
                            styles={{ root: { flex: 1 } }}
                            canRevealPassword
                        />
                        {apiKey && (
                            <IconButton
                                iconProps={{ iconName: 'Delete' }}
                                title="Clear API key"
                                onClick={handleClearKey}
                            />
                        )}
                    </Stack>
                </Stack>

                <Link href={providerInfo.link} target="_blank">
                    {providerInfo.linkText}
                </Link>

                <MessageBar messageBarType={MessageBarType.info}>
                    <pre style={{ margin: 0, whiteSpace: 'pre-wrap', fontSize: 12 }}>
                        {providerInfo.description}
                    </pre>
                </MessageBar>

                <Stack
                    horizontal
                    horizontalAlign="end"
                    tokens={{ childrenGap: 10 }}
                    style={{ marginTop: 10 }}
                >
                    <DefaultButton text="Cancel" onClick={onClose} />
                    <PrimaryButton text="Save" onClick={handleSave} />
                </Stack>

                <div className="settings-footer">
                    <p style={{ fontSize: 11, color: '#666', margin: 0 }}>
                        💡 Your API key is stored locally in your browser and never sent to our servers.
                    </p>
                </div>
            </Stack>
        </Modal>
    );
};

export default SettingsPanel;
