/**
 * Storage Service
 * Handles local storage for API keys and settings
 */

const STORAGE_KEYS = {
    API_KEY: 'excel_ai_api_key',
    PROVIDER: 'excel_ai_provider',
    MODEL: 'excel_ai_model',
    MESSAGES: 'excel_ai_messages',
};

export type LLMProvider = 'openrouter' | 'gemini' | 'huggingface';

export interface Settings {
    apiKey: string | null;
    provider: LLMProvider;
    model: string;
}

export class StorageService {
    /**
     * Save API key (encoded in localStorage)
     */
    static saveApiKey(key: string): void {
        try {
            // Simple base64 encoding (not true encryption, but obscures the key)
            const encoded = btoa(key);
            localStorage.setItem(STORAGE_KEYS.API_KEY, encoded);
        } catch (error) {
            console.error('Failed to save API key:', error);
        }
    }

    /**
     * Get API key
     */
    static getApiKey(): string | null {
        try {
            const encoded = localStorage.getItem(STORAGE_KEYS.API_KEY);
            if (!encoded) return null;
            return atob(encoded);
        } catch (error) {
            console.error('Failed to get API key:', error);
            return null;
        }
    }

    /**
     * Clear API key
     */
    static clearApiKey(): void {
        localStorage.removeItem(STORAGE_KEYS.API_KEY);
    }

    /**
     * Check if API key exists
     */
    static hasApiKey(): boolean {
        return !!localStorage.getItem(STORAGE_KEYS.API_KEY);
    }

    /**
     * Save provider preference
     */
    static saveProvider(provider: LLMProvider): void {
        localStorage.setItem(STORAGE_KEYS.PROVIDER, provider);
    }

    /**
     * Get provider preference
     */
    static getProvider(): LLMProvider {
        return (localStorage.getItem(STORAGE_KEYS.PROVIDER) as LLMProvider) || 'openrouter';
    }

    /**
     * Save model preference
     */
    static saveModel(model: string): void {
        localStorage.setItem(STORAGE_KEYS.MODEL, model);
    }

    /**
     * Get model preference
     */
    static getModel(): string {
        return localStorage.getItem(STORAGE_KEYS.MODEL) || 'meta-llama/llama-3.1-8b-instruct:free';
    }

    /**
     * Get all settings
     */
    static getSettings(): Settings {
        return {
            apiKey: this.getApiKey(),
            provider: this.getProvider(),
            model: this.getModel(),
        };
    }

    /**
     * Save messages to localStorage
     */
    static saveMessages(messages: any[]): void {
        try {
            // Keep only last 50 messages to avoid storage limits
            const trimmed = messages.slice(-50);
            localStorage.setItem(STORAGE_KEYS.MESSAGES, JSON.stringify(trimmed));
        } catch (error) {
            console.error('Failed to save messages:', error);
        }
    }

    /**
     * Get messages from localStorage
     */
    static getMessages(): any[] {
        try {
            const stored = localStorage.getItem(STORAGE_KEYS.MESSAGES);
            return stored ? JSON.parse(stored) : [];
        } catch (error) {
            console.error('Failed to get messages:', error);
            return [];
        }
    }

    /**
     * Clear all messages
     */
    static clearMessages(): void {
        localStorage.removeItem(STORAGE_KEYS.MESSAGES);
    }
}

export default StorageService;
