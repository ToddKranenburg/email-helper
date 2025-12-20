# Email Helper – Automatic Must Know & Suggested Actions

This build automatically drops a “Must Know” summary and a single “Suggested Action” into the Secretary chat timeline for every ingested email thread. Actions run through a simple, extensible state machine that supports single-step, conversational, and multi-step flows (v1: archive | create_task | more_info | skip).

## Action model
- **ActionFlow**: per `(userId, threadId)` with `actionType`, `state`, optional `draftPayload` (JSON string), `lastMessageId`.
- **States**: `suggested → draft_ready → editing → executing → completed | failed`. Single-step actions jump straight to `completed`.
- **TranscriptMessage**: timeline rendering source with `type` (`must_know`, `suggested_action`, `draft_details`, `inline_editor`, `action_result`), `content`, `payload` (JSON string), timestamps.

## Lifecycle
1. Ingest (or first render) calls `ensureAutoSummaryCards` to generate and store:
   - `must_know` (plain text essentials only)
   - `suggested_action` (JSON `{actionType,userFacingPrompt}`)
   - `ActionFlow` seeded in `suggested`.
2. Timeline renders stored transcript messages; buttons invoke `/secretary/action/draft` or `/secretary/action/execute`.
3. **archive/skip/more_info**: single POST to `/secretary/action/execute`.
4. **create_task**:
   - `POST /secretary/action/draft` → `draft_details` transcript
   - Optional `mode=edit` opens `inline_editor`; `mode=save` persists edited draft
   - `POST /secretary/action/execute` with draft → Google Tasks → `action_result`.
5. Idempotency: if `ActionFlow.state=completed` and `actionType=skip` with unchanged `lastMessageId`, auto-summarize is skipped.

## Endpoints
- `POST /secretary/auto-summarize` → Must Know + Suggested Action (idempotent)
- `POST /secretary/action/draft` → generate/edit/save create_task draft
- `POST /secretary/action/execute` → archive | create_task | more_info | skip
- Existing chat/review/intent endpoints remain unchanged.

## Frontend timeline
- New transcript message types render in-line cards:
  - `must_know`: labeled badge
  - `suggested_action`: single primary button (Archive/Draft task/Tell me more/Skip)
  - `draft_details`: summary with “Create task” + “Edit”
  - `inline_editor`: in-timeline editable fields; primary “Create task”, secondary “Save draft”
  - `action_result`: confirmation or conversational response
- Always-on composer still routes typed intents; suggested buttons call the new APIs directly.

## Adding a new action type
1. Extend the shared action union in `src/actions/persistence.ts`.
2. Add state/execute handling in `src/web/routes.ts` (draft/execute endpoints).
3. Add timeline render logic + button handling in `src/web/public/secretary.js` (no special-casing outside the render function).
4. Include transcript serialization in `ensureAutoSummaryCards` or downstream draft/execute flows.

## Database & migrations
- Schema in `prisma/schema.prisma`.
- New migration: `prisma/migrations/20241004120000_action_flows`.
- Apply with `npx prisma migrate deploy` (or `npx prisma migrate dev` for local dev) and regenerate client with `npx prisma generate`.

## Dev/test notes
- Type-check: `npx tsc --noEmit`.
- The UI relies on stored transcript messages; if you have legacy summaries without transcripts, loading a thread will auto-generate Must Know + Suggested Action on first render.
