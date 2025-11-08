-- RedefineTables
PRAGMA defer_foreign_keys=ON;
PRAGMA foreign_keys=OFF;
CREATE TABLE "new_Summary" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "threadId" TEXT NOT NULL,
    "lastMsgId" TEXT NOT NULL,
    "headline" TEXT NOT NULL DEFAULT '',
    "tldr" TEXT NOT NULL,
    "category" TEXT NOT NULL,
    "nextStep" TEXT NOT NULL,
    "confidence" TEXT NOT NULL,
    "createdAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT "Summary_threadId_fkey" FOREIGN KEY ("threadId") REFERENCES "Thread" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);
INSERT INTO "new_Summary" ("category", "confidence", "createdAt", "id", "lastMsgId", "nextStep", "threadId", "tldr") SELECT "category", "confidence", "createdAt", "id", "lastMsgId", "nextStep", "threadId", "tldr" FROM "Summary";
DROP TABLE "Summary";
ALTER TABLE "new_Summary" RENAME TO "Summary";
CREATE UNIQUE INDEX "Summary_lastMsgId_key" ON "Summary"("lastMsgId");
CREATE INDEX "Summary_threadId_idx" ON "Summary"("threadId");
PRAGMA foreign_keys=ON;
PRAGMA defer_foreign_keys=OFF;
