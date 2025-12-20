-- Create action/state tracking for suggested actions and timeline messages
CREATE TABLE "ActionFlow" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "userId" TEXT NOT NULL,
    "threadId" TEXT NOT NULL,
    "actionType" TEXT NOT NULL,
    "state" TEXT NOT NULL DEFAULT 'suggested',
    "draftPayload" TEXT,
    "lastMessageId" TEXT,
  "createdAt" TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
  "updatedAt" TIMESTAMP NOT NULL,
    CONSTRAINT "ActionFlow_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User" ("id") ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT "ActionFlow_threadId_userId_fkey" FOREIGN KEY ("threadId", "userId") REFERENCES "Thread" ("id", "userId") ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE UNIQUE INDEX "ActionFlow_userId_threadId_key" ON "ActionFlow"("userId", "threadId");
CREATE INDEX "ActionFlow_threadId_userId_updatedAt_idx" ON "ActionFlow"("threadId", "userId", "updatedAt");

CREATE TABLE "TranscriptMessage" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "userId" TEXT NOT NULL,
    "threadId" TEXT NOT NULL,
    "type" TEXT NOT NULL,
    "content" TEXT NOT NULL,
    "payload" TEXT,
  "createdAt" TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT "TranscriptMessage_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User" ("id") ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT "TranscriptMessage_threadId_userId_fkey" FOREIGN KEY ("threadId", "userId") REFERENCES "Thread" ("id", "userId") ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX "TranscriptMessage_threadId_userId_createdAt_idx" ON "TranscriptMessage"("threadId", "userId", "createdAt");
