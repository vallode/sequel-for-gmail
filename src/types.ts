export interface UserSettings {
  days: number;
  excludeReplied: boolean;
  excludedDomains: string;
  staleDays: number;
  excludeInternal: boolean;
  autoFollowUpEnabled: boolean;
  autoFollowUpSteps: AutoFollowUpStep[];
}

export interface PendingEmail {
  threadId: string;
  subject: string;
  to: string;
  date: Date;
}

export interface PendingEmailSerialized {
  threadId: string;
  subject: string;
  to: string;
  dateMs: number;
}

export interface ActionEvent {
  formInput: Record<string, string>;
  parameters: Record<string, string>;
}

export interface DismissedEmail {
  threadId: string;
  subject: string;
  to: string;
  dateMs: number;
}

export interface ThreadMessageInput {
  from: string;
  to: string;
  date: Date | GoogleAppsScript.Base.Date;
  subject: string;
}

export interface ThreadInput {
  threadId: string;
  messages: ThreadMessageInput[];
}

export interface FollowUpFilterOptions {
  cutoff: Date;
  staleCutoff: Date;
  excludeReplied: boolean;
  excludeInternal: boolean;
  excludedDomainList: string[];
  myEmail: string;
  myDomain: string;
}

export interface FollowUpQueryResult {
  emails: PendingEmail[];
  threadsScanned: number;
  searchTruncated: boolean;
  fetchedAt: Date;
  /** Set when the list could not be loaded; never cached */
  error?: string;
}

export interface AutoFollowUpStep {
  /** Days to wait after the previous message the user sent (original or nudge) */
  waitDays: number;
  template: string;
}

export interface AutoFollowUpMessageInput extends ThreadMessageInput {
  cc: string;
  isDraft: boolean;
}

export interface AutoFollowUpThreadInput {
  threadId: string;
  messages: AutoFollowUpMessageInput[];
}

export interface AutoFollowUpOptions {
  now: Date;
  steps: AutoFollowUpStep[];
  excludeInternal: boolean;
  excludedDomainList: string[];
  myEmail: string;
  myDomain: string;
}

export type AutoFollowUpStopReason =
  | "replied"
  | "sequence-complete"
  | "no-recipient"
  | "excluded";

export type AutoFollowUpWaitReason = "not-due" | "draft-pending";

export type AutoFollowUpDecision =
  | {
      action: "draft";
      threadId: string;
      stepIndex: number;
      to: string;
      cc: string;
      body: string;
    }
  | { action: "stop"; threadId: string; reason: AutoFollowUpStopReason }
  | { action: "wait"; threadId: string; reason: AutoFollowUpWaitReason };
