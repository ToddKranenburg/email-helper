-- CreateTable
CREATE TABLE "Thread" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "subject" TEXT,
    "participants" TEXT NOT NULL,
    "lastMessageTs" DATETIME NOT NULL,
    "historyId" TEXT,
    "updatedAt" DATETIME NOT NULL
);

-- CreateTable
CREATE TABLE "Summary" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "threadId" TEXT NOT NULL,
    "tldr" TEXT NOT NULL,
    "category" TEXT NOT NULL,
    "nextStep" TEXT NOT NULL,
    "confidence" TEXT NOT NULL,
    "createdAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT "Summary_threadId_fkey" FOREIGN KEY ("threadId") REFERENCES "Thread" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- CreateTable
CREATE TABLE "Processing" (
    "threadId" TEXT NOT NULL PRIMARY KEY,
    "lastProcessedMsgId" TEXT NOT NULL,
    "updatedAt" DATETIME NOT NULL
);
