// Mirrors the backend's admin API response shapes (backend/src/store/ChangeStore.ts).

export interface DocumentSummary {
  docId: string;
  title: string | null;
  createdBy: string | null;
  creatorFirstName: string;
  creatorLastName: string;
  createdAt: number;
  headVersion: number;
  contributorCount: number;
  lastEditedAt: number | null;
}

export interface UserSummary {
  id: string;
  firstName: string;
  lastName: string;
  docsCreated: number;
  changeCount: number;
  lastActive: number | null;
}

export interface WebhookRecord {
  id: string;
  url: string;
  secret: string;
  events: string[];
  active: boolean;
  createdAt: number;
}

export interface DeliveryRecord {
  id: string;
  webhookId: string;
  docId: string | null;
  event: string;
  status: "success" | "failed";
  responseCode: number | null;
  attempts: number;
  lastError: string | null;
  createdAt: number;
}

export interface ActivityEntry {
  seq: number;
  userId: string | null;
  firstName: string;
  lastName: string;
  ts: number;
  origin: string;
  opCount: number;
}

export interface DocumentActivity {
  docId: string;
  createdBy: string | null;
  creatorFirstName: string;
  creatorLastName: string;
  createdAt: number;
  entries: ActivityEntry[];
}

export interface UploadResult {
  docId: string;
  version: number;
  warnings: { code: string; message: string }[];
}
