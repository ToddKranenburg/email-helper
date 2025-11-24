-- RedefineTables
PRAGMA defer_foreign_keys=ON;
PRAGMA foreign_keys=OFF;
CREATE TABLE "new_Summary" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "threadId" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "lastMsgId" TEXT NOT NULL,
    "headline" TEXT NOT NULL DEFAULT '',
    "tldr" TEXT NOT NULL,
    "category" TEXT NOT NULL,
    "nextStep" TEXT NOT NULL,
    "convoText" TEXT NOT NULL DEFAULT '',
    "confidence" TEXT NOT NULL,
    "createdAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT "Summary_threadId_userId_fkey" FOREIGN KEY ("threadId", "userId") REFERENCES "Thread" ("id", "userId") ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT "Summary_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);
INSERT INTO "new_Summary" ("category", "confidence", "createdAt", "headline", "id", "lastMsgId", "nextStep", "threadId", "tldr", "userId") SELECT "category", "confidence", "createdAt", "headline", "id", "lastMsgId", "nextStep", "threadId", "tldr", "userId" FROM "Summary";
DROP TABLE "Summary";
ALTER TABLE "new_Summary" RENAME TO "Summary";
CREATE INDEX "Summary_threadId_userId_idx" ON "Summary"("threadId", "userId");
CREATE INDEX "Summary_userId_createdAt_idx" ON "Summary"("userId", "createdAt");
CREATE UNIQUE INDEX "Summary_userId_lastMsgId_key" ON "Summary"("userId", "lastMsgId");
PRAGMA foreign_keys=ON;
PRAGMA defer_foreign_keys=OFF;
