PRAGMA defer_foreign_keys=ON;
PRAGMA foreign_keys=OFF;

CREATE TABLE "User" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "email" TEXT NOT NULL,
    "name" TEXT,
    "picture" TEXT,
    "createdAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" DATETIME NOT NULL
);
CREATE UNIQUE INDEX "User_email_key" ON "User"("email");

INSERT INTO "User" ("id", "email", "name", "picture", "updatedAt")
VALUES ('legacy-user', 'legacy@example.com', NULL, NULL, CURRENT_TIMESTAMP)
ON CONFLICT("id") DO NOTHING;

CREATE TABLE "new_Thread" (
    "id" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "subject" TEXT,
    "participants" TEXT NOT NULL,
    "lastMessageTs" DATETIME NOT NULL,
    "historyId" TEXT,
    "fromName" TEXT,
    "fromEmail" TEXT,
    "updatedAt" DATETIME NOT NULL,
    CONSTRAINT "Thread_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User" ("id") ON DELETE CASCADE ON UPDATE CASCADE,
    PRIMARY KEY ("id", "userId")
);

INSERT INTO "new_Thread" ("id", "userId", "subject", "participants", "lastMessageTs", "historyId", "fromName", "fromEmail", "updatedAt")
SELECT "id", 'legacy-user', "subject", "participants", "lastMessageTs", "historyId", "fromName", "fromEmail", "updatedAt"
FROM "Thread";

DROP TABLE "Thread";
ALTER TABLE "new_Thread" RENAME TO "Thread";
CREATE INDEX "Thread_lastMessageTs_idx" ON "Thread"("lastMessageTs");
CREATE INDEX "Thread_userId_idx" ON "Thread"("userId");

CREATE TABLE "new_Summary" (
    "id" TEXT NOT NULL,
    "threadId" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "lastMsgId" TEXT NOT NULL,
    "headline" TEXT NOT NULL DEFAULT '',
    "tldr" TEXT NOT NULL,
    "category" TEXT NOT NULL,
    "nextStep" TEXT NOT NULL,
    "confidence" TEXT NOT NULL,
    "createdAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT "Summary_threadId_userId_fkey" FOREIGN KEY ("threadId", "userId") REFERENCES "Thread" ("id", "userId") ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT "Summary_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

INSERT INTO "new_Summary" ("id", "threadId", "userId", "lastMsgId", "headline", "tldr", "category", "nextStep", "confidence", "createdAt")
SELECT "id", "threadId", 'legacy-user', "lastMsgId", "headline", "tldr", "category", "nextStep", "confidence", "createdAt"
FROM "Summary";

DROP TABLE "Summary";
ALTER TABLE "new_Summary" RENAME TO "Summary";
CREATE UNIQUE INDEX "Summary_userId_lastMsgId_key" ON "Summary"("userId", "lastMsgId");
CREATE INDEX "Summary_threadId_userId_idx" ON "Summary"("threadId", "userId");
CREATE INDEX "Summary_userId_createdAt_idx" ON "Summary"("userId", "createdAt");

CREATE TABLE "new_Processing" (
    "threadId" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "lastProcessedMsgId" TEXT NOT NULL,
    "updatedAt" DATETIME NOT NULL,
    CONSTRAINT "Processing_threadId_userId_fkey" FOREIGN KEY ("threadId", "userId") REFERENCES "Thread" ("id", "userId") ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT "Processing_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User" ("id") ON DELETE CASCADE ON UPDATE CASCADE,
    PRIMARY KEY ("threadId", "userId")
);

INSERT INTO "new_Processing" ("threadId", "userId", "lastProcessedMsgId", "updatedAt")
SELECT "threadId", 'legacy-user', "lastProcessedMsgId", "updatedAt" FROM "Processing";

DROP TABLE IF EXISTS "Processing";
ALTER TABLE "new_Processing" RENAME TO "Processing";

PRAGMA foreign_keys=ON;
PRAGMA defer_foreign_keys=OFF;
