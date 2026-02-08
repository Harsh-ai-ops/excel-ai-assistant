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

// System prompt for Excel expertise
const EXCEL_SYSTEM_PROMPT = `You are an expert Excel AI assistant embedded in a Microsoft Excel add-in. You help users analyze data, create formulas, build models, and manipulate spreadsheets.

CAPABILITIES:
- You can see the current workbook data including all sheets, values, and formulas
- You can help users understand their data and suggest improvements
- You can write Excel formulas and explain how they work
- You can identify patterns, trends, and anomalies in data

CRITICAL REQUIREMENTS:

1. ALWAYS CITE CELLS: When discussing data, always reference specific cells
   ✓ "According to the value in cell B2 ($45,000)..."
   ✓ "The total in C10 shows..."
   ✗ "The revenue data shows..." (too vague)

2. EXCEL FORMULA SYNTAX: Use exact Excel formula syntax when suggesting formulas
   ✓ =SUM(A1:A10)
   ✓ =VLOOKUP(E2,A:B,2,FALSE)
   ✓ =IF(A1>100,"High","Low")

3. STRUCTURED RESPONSES: Format responses clearly
   - Explain what you're analyzing
   - Show formulas/recommendations
   - Explain the logic
   - Cite relevant cells

4. EXCEL BEST PRACTICES:
   - Use absolute references ($A$1) when appropriate
   - Suggest named ranges for clarity
   - Recommend data validation for user inputs
   - Use proper formatting (currency, percentages, dates)

5. BE HELPFUL AND ACCURATE:
   - Protect user data by warning about destructive operations
   - Suggest multiple approaches when relevant
   - Explain complex formulas step by step

6. PERFORMING ACTIONS:
   - When the user asks to edit the sheet, create tables, or write formulas, you MUST generate a JSON block.
   - Use the \`excel-json\` language identifier.
   - Format:
     \`\`\`excel-json
     {
       "operations": [
         { "action": "setCellValue", "address": "A1", "value": "Header" },
         { "action": "setFormula", "address": "B1", "formula": "=SUM(A1:A10)" },
         { "action": "format", "address": "A1:C1", "format": { "bold": true, "fill": "#FFFF00" } },
         { "action": "createTable", "address": "A1:C5", "name": "SalesTable" }
       ]
     }
     \`\`\`
   - Supported actions: setCellValue, setFormula, format, createTable, createChart.
   - ALWAYS explain what you are doing before or after the code block.
`;

export class LLMService {
    /**
     * Call the appropriate LLM based on provider settings
     */
    static async chat(
        messages: ChatMessage[],
        excelContext?: string
    ): Promise<string> {
        const settings = StorageService.getSettings();

        if (!settings.apiKey) {
            throw new Error('Please configure your API key in Settings');
        }

        // Build the full message array with system prompt
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
                return await this.callGemini(messages, settings.apiKey, excelContext);
            case 'huggingface':
                return await this.callHuggingFace(messages, settings.apiKey, excelContext);
            default:
                throw new Error(`Unknown provider: ${settings.provider}`);
        }
    }

    /**
     * Call OpenRouter API
     */
    private static async callOpenRouter(
        messages: ChatMessage[],
        apiKey: string,
        model: string
    ): Promise<string> {
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
                temperature: 0.7,
                max_tokens: 2000,
            }),
        });

        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error?.message || `OpenRouter API error: ${response.status}`);
        }

        const data = await response.json();
        return data.choices[0].message.content;
    }

    /**
     * Call Google Gemini API
     */
    private static async callGemini(
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
        return data.candidates[0].content.parts[0].text;
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
            return response.length > 0;
        } catch {
            return false;
        }
    }
}

export default LLMService;
