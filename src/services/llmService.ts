/**
 * LLM Service
 * Handles API calls to OpenRouter, Google Gemini, and HuggingFace
 */

import { StorageService, LLMProvider } from './storageService';

export interface ChatMessage {
    role: 'system' | 'user' | 'assistant';
    content: string;
}

// Free models available on OpenRouter
export const OPENROUTER_FREE_MODELS = [
    { id: 'meta-llama/llama-3.1-8b-instruct:free', name: 'Llama 3.1 8B (Free)' },
    { id: 'mistralai/mistral-7b-instruct:free', name: 'Mistral 7B (Free)' },
    { id: 'google/gemma-7b-it:free', name: 'Gemma 7B (Free)' },
    { id: 'microsoft/phi-3-mini-128k-instruct:free', name: 'Phi-3 Mini (Free)' },
];

/**
 * Tool Definition for Excel Actions
 */
const EXCEL_TOOLS = [
    {
        function: {
            name: "execute_excel_operations",
            description: "Execute batch operations on the Excel sheet (edit cells, write formulas, format, create tables/charts, pivot tables, manage sheets)",
            parameters: {
                type: "object",
                properties: {
                    operations: {
                        type: "array",
                        items: {
                            type: "object",
                            properties: {
                                action: {
                                    type: "string",
                                    enum: [
                                        "setCellValue", "setFormula", "format", "createTable", "createChart",
                                        "createWorksheet", "activateWorksheet", "deleteWorksheet",
                                        "sortRange", "filterRange", "autoFit", "createPivotTable"
                                    ],
                                    description: "The action to perform"
                                },
                                address: {
                                    type: "string",
                                    description: "Target cell/range address (e.g. 'A1', 'Sheet1!B2:C5')"
                                },
                                value: {
                                    type: "string",
                                    description: "Value to write (for setCellValue)"
                                },
                                formula: {
                                    type: "string",
                                    description: "Formula to write (for setFormula), must start with ="
                                },
                                format: {
                                    type: "object",
                                    description: "Format options",
                                    properties: {
                                        bold: { type: "boolean" },
                                        fill: { type: "string" },
                                        color: { type: "string" }
                                    }
                                },
                                name: {
                                    type: "string",
                                    description: "Name for table or chart"
                                },
                                title: {
                                    type: "string",
                                    description: "Title for chart"
                                },
                                // Pivot Table & Advanced Props
                                sourceSheet: { type: "string" },
                                sourceAddress: { type: "string" },
                                destinationSheet: { type: "string" },
                                destinationAddress: { type: "string" },
                                rows: { type: "array", items: { type: "string" } },
                                columns: { type: "array", items: { type: "string" } },
                                values: {
                                    type: "array",
                                    items: {
                                        type: "object",
                                        properties: {
                                            field: { type: "string" },
                                            function: { type: "string", enum: ["Sum", "Count", "Average", "Min", "Max"] }
                                        },
                                        required: ["field"]
                                    }
                                },
                                key: { type: "integer", description: "Sort column index (0-based)" },
                                ascending: { type: "boolean" }
                            },
                            required: ["action"]
                        }
                    }
                },
                required: ["operations"]
            }
        }
    }
];

// System prompt for Excel expertise
const EXCEL_SYSTEM_PROMPT = `You are the World's Best Excel AI Assistant. You are an elite analyst, data scientist, and executive dashboard designer embedded in a Microsoft Excel add-in.

GOAL: Provide the most professional, well-formatted, and insightful Excel solutions possible. You don't just "do" Excel; you create Executive-Grade spreadsheets.

CAPABILITIES:
- Advanced Data Analysis: Identify trends, correlations, and outliers.
- Executive Formatting: Use professional color palettes, proper alignment, and data validation.
- Complex Modeling: Build scalable models with clear assumptions and summary sections.
- Tool Usage: You have direct access to tools that manipulate cells, sheets, pivot tables, and charts.

CRITICAL EXPERT GUIDELINES:

1. CONTEXTUAL INTELLIGENCE: 
   - Before acting, understand the user's industry (Finance, Healthcare, Sales, etc.).
   - If they ask for a "Sales Table", don't just list columns. Include Month-over-Month growth, Total summaries, and percentage of totals.
   - Use Data Bars or Conditional Formatting suggestions in your text to guide them.

2. EXECUTIVE DESIGN PRINCIPLES:
   - HEADERS: Use bold text, specific fill colors (Slate, Deep Blue, or Dark Grey), and white font for headers.
   - ALTERNATING ROWS: Suggest or implement banded rows for readability.
   - NUMBER TYPES: Always suggest/apply proper formatting (Currency ($), Percentages (%), or 1,000 separators).
   - ALIGNMENT: Right-align numbers, left-align text. Centre-align headers.

3. DATA INTERNALS & BEST PRACTICES:
   - Use exact Excel formula syntax.
   - Reference cells explicitly (e.g., "See result in B15").
   - Use Excel Tables (ListObject) for structured data to enable easy filtering.
   - Create Pivot Tables for complex aggregations.

4. BEYOND THE BASICS:
   - When creating charts, pick the BEST type for the data (Line for time series, Pie for parts-of-a-whole, Bar for comparisons).
   - Add a "Summary" or "Dashboard" sheet if the data is large.
   - Use helper columns to make complex calculations readable.

5. ERROR PROTECTION:
   - Warn if an operation might overwrite existing data.
   - Check for #DIV/0! or #N/A errors in your formula suggestions and use IFERROR.
`;

export interface LLMResponse {
    text: string;
    toolCalls?: any[];
}

export class LLMService {
    /**
     * Call the appropriate LLM based on provider settings
     */
    static async chat(
        messages: ChatMessage[],
        excelContext?: string
    ): Promise<LLMResponse> {
        const settings = StorageService.getSettings();

        if (!settings.apiKey) {
            throw new Error('Please configure your API key in Settings');
        }

        const fullMessages: ChatMessage[] = [
            {
                role: 'system',
                content: EXCEL_SYSTEM_PROMPT + (excelContext ? `\n\n${excelContext}` : ''),
            },
            ...messages,
        ];

        switch (settings.provider) {
            case 'openrouter':
                return await this.callOpenRouter(fullMessages, settings.apiKey, settings.model);
            case 'gemini':
                return await this.callGemini(messages, settings.apiKey, excelContext); // Gemini handles context differently
            case 'huggingface':
                // HuggingFace implementation typically doesn't support tools easily in this setup
                // Fallback to text
                const text = await this.callHuggingFace(messages, settings.apiKey, excelContext);
                return { text };
            default:
                throw new Error(`Unknown provider: ${settings.provider}`);
        }
    }

    /**
     * Call OpenRouter API with Tools
     */
    private static async callOpenRouter(
        messages: ChatMessage[],
        apiKey: string,
        model: string
    ): Promise<LLMResponse> {
        const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${apiKey}`,
                'Content-Type': 'application/json',
                'HTTP-Referer': window.location.href,
                'X-Title': 'Excel AI Assistant',
            },
            body: JSON.stringify({
                model: model,
                messages: messages,
                tools: EXCEL_TOOLS.map(t => ({ type: "function", function: t.function })),
                tool_choice: "auto",
                temperature: 0.7,
                max_tokens: 2000,
            }),
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error?.message || `OpenRouter API error: ${response.status}`);
        }

        const data = await response.json();
        const choice = data.choices[0];
        const message = choice.message;

        return {
            text: message.content || '',
            toolCalls: message.tool_calls ? message.tool_calls.map((tc: any) => ({
                name: tc.function.name,
                arguments: JSON.parse(tc.function.arguments)
            })) : undefined
        };
    }

    /**
     * Call Google Gemini API with Tools
     */
    private static async callGemini(
        messages: ChatMessage[],
        apiKey: string,
        excelContext?: string
    ): Promise<LLMResponse> {
        let prompt = EXCEL_SYSTEM_PROMPT;

        if (excelContext) {
            prompt += `\n\n${excelContext}`;
        }
        prompt += '\n\nConversation:\n';

        messages.forEach((msg) => {
            const role = msg.role === 'user' ? 'User' : 'Assistant';
            prompt += `${role}: ${msg.content}\n`;
        });

        prompt += 'Assistant:';

        const tools = [{
            function_declarations: [EXCEL_TOOLS[0].function]
        }];

        const response = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${apiKey}`,
            {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    contents: [
                        {
                            parts: [{ text: prompt }],
                        },
                    ],
                    tools: tools,
                    generationConfig: {
                        temperature: 0.7,
                        maxOutputTokens: 2000,
                    },
                }),
            }
        );

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error?.message || `Gemini API error: ${response.status}`);
        }

        const data = await response.json();
        const candidate = data.candidates[0];
        const parts = candidate.content.parts;

        let text = '';
        let toolCalls = [];

        for (const part of parts) {
            if (part.text) text += part.text;
            if (part.functionCall) {
                // Gemini returns args as detailed object, or potentially just simpler object?
                // V1Beta usually returns 'args' as object directly
                toolCalls.push({
                    name: part.functionCall.name,
                    arguments: part.functionCall.args
                });
            }
        }

        return { text, toolCalls: toolCalls.length > 0 ? toolCalls : undefined };
    }

    /**
     * Call HuggingFace Inference API
     */
    private static async callHuggingFace(
        messages: ChatMessage[],
        apiKey: string,
        excelContext?: string
    ): Promise<string> {
        // Build prompt from messages
        let prompt = EXCEL_SYSTEM_PROMPT;
        if (excelContext) {
            prompt += `\n\n${excelContext}`;
        }
        prompt += '\n\nConversation:\n';

        messages.forEach((msg) => {
            const role = msg.role === 'user' ? 'User' : 'Assistant';
            prompt += `${role}: ${msg.content}\n`;
        });

        prompt += 'Assistant:';

        const response = await fetch(
            'https://api-inference.huggingface.co/models/meta-llama/Llama-3.2-1B-Instruct',
            {
                method: 'POST',
                headers: {
                    Authorization: `Bearer ${apiKey}`,
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    inputs: prompt,
                    parameters: {
                        max_new_tokens: 1000,
                        temperature: 0.7,
                    },
                }),
            }
        );

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error || `HuggingFace API error: ${response.status}`);
        }

        const data = await response.json();

        // HuggingFace returns generated text differently
        if (Array.isArray(data)) {
            return data[0].generated_text.replace(prompt, '').trim();
        }
        return data.generated_text?.replace(prompt, '').trim() || 'No response generated';
    }

    /**
     * Get list of available models from OpenRouter
     */
    static async getOpenRouterModels(): Promise<{ id: string; name: string }[]> {
        try {
            const response = await fetch('https://openrouter.ai/api/v1/models');
            if (!response.ok) {
                return OPENROUTER_FREE_MODELS;
            }

            const data = await response.json();
            const models = data.data.map((m: any) => ({
                id: m.id,
                name: m.name || m.id,
            }));

            // Sort: Free models first, then alphabetical
            return models.sort((a: any, b: any) => {
                const aFree = a.id.includes(':free');
                const bFree = b.id.includes(':free');
                if (aFree && !bFree) return -1;
                if (!aFree && bFree) return 1;
                return a.name.localeCompare(b.name);
            });
        } catch (e) {
            console.error('Failed to fetch models', e);
            return OPENROUTER_FREE_MODELS;
        }
    }

    /**
     * Test if the API key is valid
     */
    static async testConnection(): Promise<boolean> {
        try {
            const response = await this.chat([{ role: 'user', content: 'Say "Hello"' }]);
            return !!response.text;
        } catch {
            return false;
        }
    }
}

export default LLMService;
