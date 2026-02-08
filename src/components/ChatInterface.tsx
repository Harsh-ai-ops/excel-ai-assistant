import * as React from 'react';
/* global Excel */
import { useState, useRef, useEffect } from 'react';
import {
    Stack,
    TextField,
    PrimaryButton,
    DefaultButton,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType,
    IconButton,
} from '@fluentui/react';
import { LLMService, ChatMessage } from '../services/llmService';
import { ExcelService } from '../services/excelService';
import { ContextBuilder } from '../services/contextBuilder';
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

// Helper to extract JSON from response
const extractExcelActions = (content: string): any | null => {
    const match = content.match(/```excel-json\s*([\s\S]*?)```/);
    if (match && match[1]) {
        try {
            return JSON.parse(match[1]);
        } catch (e) {
            console.error('Failed to parse Excel actions', e);
        }
    }
    return null;
};

const ChatInterface: React.FC<ChatInterfaceProps> = ({ onOpenSettings }) => {
    const [messages, setMessages] = useState<Message[]>([]);
    const [input, setInput] = useState('');
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [pendingActions, setPendingActions] = useState<any[]>([]);
    const [executingActions, setExecutingActions] = useState(false);
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
    }, [messages, pendingActions, isLoading]);

    const handleApplyActions = async () => {
        if (pendingActions.length === 0) return;

        setExecutingActions(true);
        try {
            await ExcelService.executeOperations(pendingActions);

            // Add success message
            const successMsg: Message = {
                id: Date.now().toString(),
                role: 'assistant',
                content: '✅ Changes applied successfully!',
                timestamp: new Date()
            };
            setMessages(prev => [...prev, successMsg]);
            setPendingActions([]);
        } catch (error: any) {
            setError('Failed to apply changes: ' + error.message);
        } finally {
            setExecutingActions(false);
        }
    };

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
        setPendingActions([]); // Clear previous pending actions

        try {
            // Get holistic workbook context
            const workbookMetadata = await ContextBuilder.getWorkbookMetadata();

            // Get active sheet data (existing logic)
            let excelContext = "";
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getUsedRange();
                sheet.load("name");
                range.load("values, formulas, address");
                await context.sync();

                excelContext = ContextBuilder.buildContextString(workbookMetadata,
                    `Sheet: ${sheet.name}\nRange: ${range.address}\nData: ${JSON.stringify(range.values)}\nFormulas: ${JSON.stringify(range.formulas)}`);
            });

            // Call LLM
            const response = await LLMService.chat(
                [...messages, userMessage],
                excelContext
            );
            let assistantContent = response.text || '';
            const toolCalls = response.toolCalls;

            if (toolCalls && toolCalls.length > 0) {
                // Flatten operations from all tool calls
                const allOperations = toolCalls.flatMap(tc => {
                    // OpenRouter/Gemini might return args as object or string
                    // We already parsed it in service if possible, but double check
                    if (tc.name === 'execute_excel_operations') {
                        return tc.arguments.operations || [];
                    }
                    return [];
                });

                if (allOperations.length > 0) {
                    setPendingActions(allOperations);
                    if (!assistantContent) {
                        assistantContent = "I have prepared the changes for you. Please review and apply them.";
                    }
                }
            }

            // Older JSON fallback (just in case model refuses to use tools)
            const actionsData = extractExcelActions(assistantContent);
            if (actionsData && actionsData.operations) {
                setPendingActions(actionsData.operations);
                assistantContent = assistantContent.replace(/```excel-json[\s\S]*?```/, '').trim() || "I have prepared the changes for you.";
            }

            const assistantMessage: Message = {
                id: (Date.now() + 1).toString(),
                role: 'assistant',
                content: assistantContent,
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
        setPendingActions([]);
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

    const handleCitationClick = async (address: string) => {
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange(address);
                range.select();
                await context.sync();
            });
        } catch (error) {
            console.error("Failed to navigate to citation", error);
        }
    };

    // Helper to render message with clickable citations
    const renderMessageContent = (content: string) => {
        // Find patterns like [Cell: A1] or [Cell: Sheet1!B2]
        const parts = (content || '').split(/(\[Cell: [^\]]+\])/);
        return parts.map((part, i) => {
            const match = part.match(/\[Cell: ([^\]]+)\]/);
            if (match) {
                const address = match[1];
                return (
                    <span
                        key={i}
                        onClick={() => handleCitationClick(address)}
                        style={{ color: '#0078d4', textDecoration: 'underline', cursor: 'pointer', fontWeight: 'bold' }}
                    >
                        {part}
                    </span>
                );
            }
            return <span key={i}>{part}</span>;
        });
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

            {/* Pending Actions Banner */}
            {pendingActions.length > 0 && (
                <MessageBar
                    messageBarType={MessageBarType.success}
                    actions={
                        <div>
                            <PrimaryButton
                                text={executingActions ? "Applying..." : "Apply Changes"}
                                onClick={handleApplyActions}
                                disabled={executingActions}
                            />
                            <DefaultButton
                                text="Discard"
                                onClick={() => setPendingActions([])}
                                disabled={executingActions}
                                styles={{ root: { marginLeft: 10 } }}
                            />
                        </div>
                    }
                >
                    <b>AI suggested {pendingActions.length} changes.</b> Click Apply to update your sheet.
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
                            <div
                                className={`message-bubble ${msg.role}`}
                                style={{
                                    whiteSpace: 'pre-wrap',
                                    wordBreak: 'break-word'
                                }}
                            >
                                {renderMessageContent(msg.content)}
                            </div>
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
