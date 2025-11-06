-- AlterTable
ALTER TABLE "Thread" ADD COLUMN "fromEmail" TEXT;
ALTER TABLE "Thread" ADD COLUMN "fromName" TEXT;

-- CreateIndex
CREATE INDEX "Thread_lastMessageTs_idx" ON "Thread"("lastMessageTs");
