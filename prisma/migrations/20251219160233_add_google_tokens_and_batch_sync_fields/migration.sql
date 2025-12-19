-- AlterTable
ALTER TABLE "User" ADD COLUMN "lastActiveAt" DATETIME;
ALTER TABLE "User" ADD COLUMN "lastBatchSyncAt" DATETIME;
ALTER TABLE "User" ADD COLUMN "lastBatchSyncError" TEXT;
ALTER TABLE "User" ADD COLUMN "lastBatchSyncStatus" TEXT;

-- CreateTable
CREATE TABLE "GoogleToken" (
    "userId" TEXT NOT NULL PRIMARY KEY,
    "refreshToken" TEXT NOT NULL,
    "accessToken" TEXT,
    "scope" TEXT,
    "tokenType" TEXT,
    "expiryDate" DATETIME,
    "updatedAt" DATETIME NOT NULL,
    CONSTRAINT "GoogleToken_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);
