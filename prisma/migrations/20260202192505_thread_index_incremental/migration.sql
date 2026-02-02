-- AlterTable
ALTER TABLE "ActionFlow" ALTER COLUMN "createdAt" SET DATA TYPE TIMESTAMP(3),
ALTER COLUMN "updatedAt" SET DATA TYPE TIMESTAMP(3);

-- AlterTable
ALTER TABLE "Thread" ADD COLUMN     "contentVersion" TEXT,
ADD COLUMN     "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
ADD COLUMN     "extracted" JSONB,
ADD COLUMN     "gmailLabelIds" JSONB,
ADD COLUMN     "inPrimaryInbox" BOOLEAN NOT NULL DEFAULT true,
ADD COLUMN     "lastMessageId" TEXT,
ADD COLUMN     "lastScoredAt" TIMESTAMP(3),
ADD COLUMN     "priorityReason" TEXT,
ADD COLUMN     "priorityScore" DOUBLE PRECISION,
ADD COLUMN     "scoreVersion" TEXT NOT NULL DEFAULT 'priority-v1',
ADD COLUMN     "snippet" TEXT,
ADD COLUMN     "suggestedActionType" TEXT,
ADD COLUMN     "unreadCount" INTEGER NOT NULL DEFAULT 0,
ALTER COLUMN "participants" DROP NOT NULL;

-- AlterTable
ALTER TABLE "TranscriptMessage" ALTER COLUMN "createdAt" SET DATA TYPE TIMESTAMP(3);

-- CreateTable
CREATE TABLE "GmailAccount" (
    "userId" TEXT NOT NULL,
    "emailAddress" TEXT NOT NULL,
    "historyCursor" TEXT,
    "watchExpiration" TIMESTAMP(3),
    "lastSyncAt" TIMESTAMP(3),
    "lastInitialSyncAt" TIMESTAMP(3),
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "GmailAccount_pkey" PRIMARY KEY ("userId")
);

-- CreateTable
CREATE TABLE "ThreadContentCache" (
    "userId" TEXT NOT NULL,
    "threadId" TEXT NOT NULL,
    "contentText" TEXT NOT NULL,
    "contentVersion" TEXT NOT NULL,
    "updatedAt" TIMESTAMP(3) NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- CreateTable
CREATE TABLE "PrioritizationBatch" (
    "id" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "status" TEXT NOT NULL,
    "totalThreadsPlanned" INTEGER NOT NULL,
    "processedThreads" INTEGER NOT NULL,
    "deferredThreads" INTEGER NOT NULL,
    "startedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "finishedAt" TIMESTAMP(3),
    "trigger" TEXT NOT NULL,

    CONSTRAINT "PrioritizationBatch_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "DeferredPrioritization" (
    "id" TEXT NOT NULL,
    "userId" TEXT NOT NULL,
    "threadId" TEXT NOT NULL,
    "reason" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "DeferredPrioritization_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "ThreadContentCache_userId_updatedAt_idx" ON "ThreadContentCache"("userId", "updatedAt");

-- CreateIndex
CREATE UNIQUE INDEX "ThreadContentCache_userId_threadId_key" ON "ThreadContentCache"("userId", "threadId");

-- CreateIndex
CREATE INDEX "PrioritizationBatch_userId_startedAt_idx" ON "PrioritizationBatch"("userId", "startedAt");

-- CreateIndex
CREATE INDEX "DeferredPrioritization_userId_createdAt_idx" ON "DeferredPrioritization"("userId", "createdAt");

-- CreateIndex
CREATE UNIQUE INDEX "DeferredPrioritization_userId_threadId_key" ON "DeferredPrioritization"("userId", "threadId");

-- AddForeignKey
ALTER TABLE "GmailAccount" ADD CONSTRAINT "GmailAccount_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User"("id") ON DELETE CASCADE ON UPDATE CASCADE;
