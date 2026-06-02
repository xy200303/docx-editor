# AI Word Editor — agent + editor reference example

The canonical "plug an agent into the editor" demo. A Next.js app that:

- Mounts `<DocxEditor>` with the controllable right-hand `agentPanel` slot.
- Wires `useDocxAgentTools` to an OpenAI-backed `/api/chat` route.
- Streams live tool calls into the running editor: comments, tracked rewrites,
  formatting, table insertion, and image insertion.
- Adds a demo-only `generate_image` server tool so the model can generate a
  base64 image and then call `insert_image` to place it in the DOCX.

## Run it

```bash
cp .env.example .env.local
# Fill OPENAI_API_KEY in .env.local
# Optional: set OPENAI_BASE_URL for an OpenAI-compatible gateway/proxy
bun install
bun run dev --filter agent-chat-demo
```

Open http://localhost:3002, type into the document or open a DOCX, then ask the
assistant to rewrite text, insert a table, or generate an image.

## What the code does

Three pieces are worth copying into your own app.

**1. Server route (`app/api/chat/route.ts`)** — proxy your LLM call. Built-in
docx tools come from `getAiSdkTools()`. The demo adds a server-executed
`generate_image` tool because image generation needs your API key:

```ts
const openai = createOpenAI({
  apiKey: process.env.OPENAI_API_KEY,
  baseURL: process.env.OPENAI_BASE_URL || process.env.OPENAI_API_BASE_URL,
});

const tools = {
  ...getAiSdkTools(),
  generate_image: tool({
    inputSchema: jsonSchema({
      /* prompt schema */
    }),
    execute: async ({ prompt }) => {
      const result = await generateImage({ model: openai.image('gpt-image-1'), prompt });
      return { src: `data:${result.image.mediaType};base64,${result.image.base64}` };
    },
  }),
};
```

**2. React page (`app/page.tsx`)** — the hook owns the live editor bridge:

```tsx
const { executeToolCall, getContext } = useDocxAgentTools({
  editorRef,
  author: 'AI Editor',
});
```

**3. Tool execution loop** — client-side editor tools are run locally through
`executeToolCall` and pushed back into the AI SDK conversation. The server-side
`generate_image` tool returns a data URL; the model then calls `insert_image`
with that `src`.

## Common prompts

- "Rewrite the selected text to be clearer."
- "Insert a 4-row project plan table here."
- "Generate and insert a simple product roadmap image."

## Repurposing

Swap the system prompt in `app/api/chat/route.ts`, filter tools with
`include`/`exclude` in `useDocxAgentTools`, and replace the panel UI with your
own chat surface. The editor bridge and tool catalog stay the same.
