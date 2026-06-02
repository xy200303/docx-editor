/**
 * Chat API route — Vercel AI SDK + OpenAI + the docx agent toolkit.
 *
 * Streams tokens token-by-token to the client via AI SDK's
 * `toUIMessageStreamResponse()`. The client uses `useChat` from
 * `@ai-sdk/react` and runs tool calls through `useDocxAgentTools` —
 * zero stream-parsing code. Pure BYOA — no library changes needed.
 */

import { NextRequest } from 'next/server';
import {
  streamText,
  type UIMessage,
  convertToModelMessages,
  stepCountIs,
  generateImage,
  jsonSchema,
  tool,
} from 'ai';
import { createOpenAI } from '@ai-sdk/openai';
import { type AgentContextSnapshot } from '@eigenpal/docx-editor-agents/server';
import { getAiSdkTools } from '@eigenpal/docx-editor-agents/ai-sdk/server';

type GenerateImageInput = {
  prompt: string;
  alt?: string;
  aspectRatio?: '1:1' | '4:3' | '3:4' | '16:9' | '9:16';
  displayWidth?: number;
  displayHeight?: number;
};

type GenerateImageOutput = {
  src: string;
  alt: string;
  width: number;
  height: number;
};

type ParsedImageSize = {
  value: `${number}x${number}`;
  width: number;
  height: number;
};

type ReasoningEffort = 'none' | 'minimal' | 'low' | 'medium' | 'high' | 'xhigh';

function env(name: string): string | undefined {
  const value = process.env[name]?.trim();
  return value ? value : undefined;
}

function envBool(name: string, fallback: boolean): boolean {
  const value = env(name)?.toLowerCase();
  if (!value) return fallback;
  if (['1', 'true', 'yes', 'on'].includes(value)) return true;
  if (['0', 'false', 'no', 'off'].includes(value)) return false;
  return fallback;
}

function envPositiveInt(name: string): number | undefined {
  const value = env(name);
  if (!value) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? Math.floor(parsed) : undefined;
}

function envReasoningEffort(): ReasoningEffort | undefined {
  const value = env('OPENAI_REASONING_EFFORT');
  if (!value) return undefined;
  const allowed: ReasoningEffort[] = ['none', 'minimal', 'low', 'medium', 'high', 'xhigh'];
  return allowed.includes(value as ReasoningEffort) ? (value as ReasoningEffort) : undefined;
}

function parseImageSize(value: string | undefined): ParsedImageSize | undefined {
  const match = value?.match(/^(\d+)x(\d+)$/);
  if (!match) return undefined;

  const width = Number(match[1]);
  const height = Number(match[2]);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return undefined;
  }

  return {
    value: `${width}x${height}` as `${number}x${number}`,
    width,
    height,
  };
}

function getImageDisplayRatio(
  aspectRatio: GenerateImageInput['aspectRatio'] | undefined,
  imageSize: ParsedImageSize | undefined
): [number, number] {
  if (aspectRatio) return aspectRatio.split(':').map(Number) as [number, number];
  if (imageSize) return [imageSize.width, imageSize.height];
  return [16, 9];
}

async function withTimeout<T>(
  timeoutMs: number | undefined,
  parentSignal: AbortSignal | undefined,
  run: (signal: AbortSignal | undefined) => Promise<T>
): Promise<T> {
  if (!timeoutMs) return run(parentSignal);

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  let onParentAbort: (() => void) | undefined;

  if (parentSignal) {
    if (parentSignal.aborted) {
      controller.abort();
    } else {
      onParentAbort = () => controller.abort();
      parentSignal.addEventListener('abort', onParentAbort, { once: true });
    }
  }

  try {
    return await run(controller.signal);
  } finally {
    clearTimeout(timer);
    if (parentSignal && onParentAbort) {
      parentSignal.removeEventListener('abort', onParentAbort);
    }
  }
}

const requestTimeoutMs = envPositiveInt('OPENAI_REQUEST_TIMEOUT_MS');
const maxOutputTokens = envPositiveInt('OPENAI_MAX_OUTPUT_TOKENS');
const chatModel = env('OPENAI_CHAT_MODEL') || env('OPENAI_MODEL') || 'gpt-5.4-mini';
const imageModel = env('OPENAI_IMAGE_MODEL') || 'gpt-image-1';
const imageSize = parseImageSize(env('OPENAI_IMAGE_SIZE'));
const imageGenerationEnabled = envBool('OPENAI_AUTO_IMAGE_GENERATION', true);

const chatOpenAI = createOpenAI({
  apiKey: env('OPENAI_API_KEY'),
  baseURL: env('OPENAI_BASE_URL') || env('OPENAI_API_BASE_URL'),
});

const imageOpenAI = createOpenAI({
  apiKey: env('OPENAI_IMAGE_API_KEY') || env('OPENAI_API_KEY'),
  baseURL: env('OPENAI_IMAGE_BASE_URL') || env('OPENAI_BASE_URL') || env('OPENAI_API_BASE_URL'),
});

function chatProviderOptions() {
  if (!envBool('OPENAI_THINKING_ENABLED', false)) return undefined;

  return {
    openai: {
      forceReasoning: true,
      ...(envReasoningEffort() ? { reasoningEffort: envReasoningEffort() } : {}),
    },
  };
}

const generateImageTool = tool<GenerateImageInput, GenerateImageOutput>({
  description:
    'Generate an image for the document and return a base64 data URL. After this tool returns, call insert_image with the returned src/alt/width/height.',
  inputSchema: jsonSchema<GenerateImageInput>({
    type: 'object',
    properties: {
      prompt: {
        type: 'string',
        description:
          'Image prompt. Include the document context, visual style, and what the image should communicate.',
      },
      alt: {
        type: 'string',
        description: 'Short accessible description to use as the inserted image alt text.',
      },
      aspectRatio: {
        type: 'string',
        enum: ['1:1', '4:3', '3:4', '16:9', '9:16'],
        description:
          'Desired image aspect ratio. Defaults to OPENAI_IMAGE_SIZE when set, otherwise 16:9.',
      },
      displayWidth: {
        type: 'number',
        description: 'Rendered width in the document, in pixels. Defaults to 520.',
      },
      displayHeight: {
        type: 'number',
        description:
          'Rendered height in the document, in pixels. Defaults from aspect ratio or OPENAI_IMAGE_SIZE.',
      },
    },
    required: ['prompt'],
  }),
  execute: async (input, { abortSignal }) => {
    if (!input.prompt || typeof input.prompt !== 'string') {
      throw new Error('generate_image requires a prompt.');
    }

    const shapeOptions =
      imageSize && !input.aspectRatio
        ? { size: imageSize.value }
        : { aspectRatio: input.aspectRatio ?? '16:9' };

    const result = await withTimeout(requestTimeoutMs, abortSignal, (signal) =>
      generateImage({
        model: imageOpenAI.image(imageModel),
        prompt: input.prompt,
        ...shapeOptions,
        providerOptions: {
          openai: {
            quality: env('OPENAI_IMAGE_QUALITY') || 'auto',
            outputFormat: 'png',
          },
        },
        abortSignal: signal,
      })
    );

    const [ratioW, ratioH] = getImageDisplayRatio(input.aspectRatio, imageSize);
    const width = input.displayWidth && input.displayWidth > 0 ? input.displayWidth : 520;
    const height =
      input.displayHeight && input.displayHeight > 0
        ? input.displayHeight
        : Math.round(width * (ratioH / ratioW));

    return {
      src: `data:${result.image.mediaType};base64,${result.image.base64}`,
      alt: input.alt || input.prompt.slice(0, 120),
      width,
      height,
    };
  },
});

// Built-in docx tools have no `execute` → AI SDK forwards each call to the client's
// `useChat({ onToolCall })`, where it runs against the live editor via
// `useDocxAgentTools().executeToolCall`.
const aiSdkTools = {
  ...getAiSdkTools(),
  ...(imageGenerationEnabled ? { generate_image: generateImageTool } : {}),
};

const SYSTEM_PROMPT = `You are a WPS-style real-time Word editing assistant embedded inside the user's DOCX editor. Your job is to help the user rewrite, format, structure, and enrich the document while the editor stays live.

Core behavior:
 - Prefer visible, useful edits over long explanations. After tool calls, summarize what changed in one short paragraph.
 - Use read_document, read_selection, find_text, read_page, read_pages, or read_content_controls before targeted edits so you can anchor precisely.
 - For normal writing, insertion, rewriting, replacement, deletion, and fill-in requests, edit directly with insert_text, replace_text, set_content_control, or set_content_control_value. Do not create comments or tracked changes for ordinary edits.
 - Use suggest_change only when the user explicitly asks for suggestions, reviewable tracked changes, revision marks, accept/reject workflow, or "do not apply it yet".
 - Use add_comment only when the user asks for comments, critique, notes, or review feedback instead of applying the edit.
 - Prefer read_content_controls plus set_content_control for template text fields, or set_content_control_value for checkbox/dropdown/date controls, when template documents expose Word content-control/SDT tag, alias, id, or type anchors. Prefer paraId/search tools for ordinary prose.
 - For formatting requests, call apply_formatting or set_paragraph_style directly.
 - For tables, call insert_table with rows, columns, and data. Use paraId when the user points to a location; otherwise insert at the cursor.
 ${
   imageGenerationEnabled
     ? ' - For images, never invent base64. First call generate_image with a concise prompt, then call insert_image with the returned src, alt, width, and height.'
     : ' - If the user asks for generated images, explain that image generation is disabled for this server.'
 }
 - If the user asks for a table${imageGenerationEnabled ? ' or image' : ''} but the placement is ambiguous, use the current cursor/selection from the CONTEXT block.
 - Keep table sizes practical: usually 2–8 rows and 2–5 columns unless the user asks for more.

Workflow:
 1. Understand the user's requested edit and the current document context.
 2. Read just enough document content to choose precise paraIds.
 3. Execute the edit with the smallest number of tool calls that produces a polished result.
 4. Tell the user what you changed and where.

Tone:
 - Clear, professional, and concise.
 - Use Chinese when the user writes Chinese; otherwise match the user's language.
 - Do not expose implementation details unless the user asks.

You will see a CONTEXT block describing the user's current selection and page. Use it sparingly — most of your work happens against the whole document.`;

function isAllowedOrigin(origin: string | null): boolean {
  const allowList = process.env.ALLOWED_ORIGINS;
  if (!allowList) return true;
  if (!origin) return false;
  return allowList
    .split(',')
    .map((o) => o.trim())
    .includes(origin);
}

function formatContext(ctx: AgentContextSnapshot | undefined): string {
  if (!ctx) return '';
  const lines: string[] = [];
  if (ctx.totalPages) lines.push(`Document has ${ctx.totalPages} page(s).`);
  if (ctx.currentPage) lines.push(`User is viewing page ${ctx.currentPage}.`);
  const sel = ctx.selection;
  if (sel?.paraId) {
    if (sel.selectedText)
      lines.push(`User selection: "${sel.selectedText}" in paragraph ${sel.paraId}.`);
    else lines.push(`User cursor is in paragraph ${sel.paraId}.`);
  }
  return lines.length > 0 ? `\n\n[CONTEXT]\n${lines.join('\n')}` : '';
}

export async function POST(request: NextRequest) {
  if (!isAllowedOrigin(request.headers.get('origin'))) {
    return new Response('Origin not allowed', { status: 403 });
  }

  const { messages, context } = (await request.json()) as {
    messages: UIMessage[];
    context?: AgentContextSnapshot;
  };

  const result = streamText({
    model: chatOpenAI.chat(chatModel),
    system: SYSTEM_PROMPT + formatContext(context),
    messages: await convertToModelMessages(messages),
    tools: aiSdkTools,
    maxOutputTokens,
    timeout: requestTimeoutMs,
    abortSignal: request.signal,
    providerOptions: chatProviderOptions(),
    // AI SDK defaults to a single step — without this, the model never
    // gets a chance to read its own tool results and write a final reply.
    // 12 lets it read, edit, insert structures, and summarize without runaway loops.
    stopWhen: stepCountIs(12),
  });

  return result.toUIMessageStreamResponse();
}
