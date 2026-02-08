import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import {
    Stack,
    TextField,
    PrimaryButton,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType,
    IconButton,
} from '@fluentui/react';
import { LLMService, ChatMessage } from '../services/llmService';
import { ExcelService } from '../services/excelService';
import { StorageService } from '../services/storageService';

interface Message {
    id: string;
    role: 'user' | 'assistant';
    content: string;
    timestamp: Date;
}

interface ChatInterfaceProps {
    onOpenSettings: () => void;
}

const ChatInterface: React.FC<ChatInterfaceProps> = ({ onOpenSettings }) => {
    const [messages, setMessages] = useState<Message[]>([]);
    const [input, setInput] = useState('');
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const messagesEndRef = useRef<HTMLDivElement>(null);

    // Load messages on mount
    useEffect(() => {
        const saved = StorageService.getMessages();
        if (saved.length > 0) {
            setMessages(
                saved.map((m: any) => ({
                    ...m,
                    timestamp: new Date(m.timestamp),
                }))
            );
        }
    }, []);

    // Save messages when they change
    useEffect(() => {
        if (messages.length > 0) {
            StorageService.saveMessages(messages);
        }
    }, [messages]);

    // Auto-scroll to bottom
    useEffect(() => {
        messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
    }, [messages]);

    const sendMessage = async () => {
        if (!input.trim() || isLoading) return;

        // Check for API key
        if (!StorageService.hasApiKey()) {
            setError('Please configure your API key in Settings first.');
            return;
        }

        const userMessage: Message = {
            id: Date.now().toString(),
            role: 'user',
            content: input.trim(),
            timestamp: new Date(),
        };

        setMessages((prev) => [...prev, userMessage]);
        setInput('');
        setError(null);
        setIsLoading(true);

        try {
            // Get Excel context
            const excelContext = await ExcelService.buildContextForLLM();

            // Build message history for API
            const chatHistory: ChatMessage[] = messages.slice(-10).map((m) => ({
                role: m.role,
                content: m.content,
            }));
            chatHistory.push({ role: 'user', content: userMessage.content });

            // Call LLM
            const response = await LLMService.chat(chatHistory, excelContext);

            const assistantMessage: Message = {
                id: (Date.now() + 1).toString(),
                role: 'assistant',
                content: response,
                timestamp: new Date(),
            };

            setMessages((prev) => [...prev, assistantMessage]);
        } catch (err: any) {
            setError(err.message || 'Failed to get response');
        } finally {
            setIsLoading(false);
        }
    };

    const handleKeyPress = (e: React.KeyboardEvent) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            sendMessage();
        }
    };

    const clearChat = () => {
        setMessages([]);
        StorageService.clearMessages();
    };

    const formatMessage = (content: string) => {
        // Basic formatting: convert code blocks and line breaks
        return content.split('\n').map((line, i) => (
            <React.Fragment key={i}>
                {line}
                <br />
            </React.Fragment>
        ));
    };

    return (
        <Stack className="chat-interface" tokens={{ childrenGap: 10 }}>
            {/* Header */}
            <Stack
                horizontal
                horizontalAlign="space-between"
                verticalAlign="center"
                className="chat-header"
            >
                <span className="chat-title">💬 Excel AI Assistant</span>
                <Stack horizontal tokens={{ childrenGap: 5 }}>
                    <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Clear chat"
                        onClick={clearChat}
                    />
                    <IconButton
                        iconProps={{ iconName: 'Settings' }}
                        title="Settings"
                        onClick={onOpenSettings}
                    />
                </Stack>
            </Stack>

            {/* Error bar */}
            {error && (
                <MessageBar
                    messageBarType={MessageBarType.error}
                    onDismiss={() => setError(null)}
                    dismissButtonAriaLabel="Close"
                >
                    {error}
                </MessageBar>
            )}

            {/* API Key warning */}
            {!StorageService.hasApiKey() && (
                <MessageBar messageBarType={MessageBarType.warning}>
                    ⚠️ Please configure your API key in{' '}
                    <a onClick={onOpenSettings} style={{ cursor: 'pointer', color: '#005a9e' }}>
                        Settings
                    </a>{' '}
                    to start chatting.
                </MessageBar>
            )}

            {/* Messages */}
            <div className="messages-container">
                {messages.length === 0 ? (
                    <div className="empty-state">
                        <h3>👋 Welcome to Excel AI Assistant!</h3>
                        <p>I can help you with:</p>
                        <ul>
                            <li>📊 Analyzing your data</li>
                            <li>📝 Writing Excel formulas</li>
                            <li>📈 Creating insights from your spreadsheet</li>
                            <li>🔍 Finding patterns and trends</li>
                        </ul>
                        <p>Just type your question below!</p>
                    </div>
                ) : (
                    messages.map((msg) => (
                        <div
                            key={msg.id}
                            className={`message ${msg.role === 'user' ? 'user-message' : 'assistant-message'}`}
                        >
                            <div className="message-header">
                                <span className="message-role">
                                    {msg.role === 'user' ? '👤 You' : '🤖 Assistant'}
                                </span>
                                <span className="message-time">
                                    {msg.timestamp.toLocaleTimeString([], {
                                        hour: '2-digit',
                                        minute: '2-digit',
                                    })}
                                </span>
                            </div>
                            <div className="message-content">{formatMessage(msg.content)}</div>
                        </div>
                    ))
                )}

                {isLoading && (
                    <div className="message assistant-message">
                        <div className="message-header">
                            <span className="message-role">🤖 Assistant</span>
                        </div>
                        <div className="message-content">
                            <Spinner size={SpinnerSize.small} label="Thinking..." />
                        </div>
                    </div>
                )}

                <div ref={messagesEndRef} />
            </div>

            {/* Input */}
            <Stack horizontal tokens={{ childrenGap: 8 }} className="input-container">
                <TextField
                    placeholder="Ask about your Excel data..."
                    value={input}
                    onChange={(_, newValue) => setInput(newValue || '')}
                    onKeyPress={handleKeyPress}
                    disabled={isLoading}
                    multiline
                    autoAdjustHeight
                    className="chat-input"
                    styles={{ root: { flex: 1 } }}
                />
                <PrimaryButton
                    iconProps={{ iconName: 'Send' }}
                    onClick={sendMessage}
                    disabled={isLoading || !input.trim()}
                    styles={{ root: { height: 'auto', minHeight: 32 } }}
                />
            </Stack>
        </Stack>
    );
};

export default ChatInterface;
