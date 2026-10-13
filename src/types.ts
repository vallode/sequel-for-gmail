export interface UserSettings {
  days: number;
  excludeReplied: boolean;
  autoLabel: string;
  excludedDomains: string;
  staleDays: number;
  excludeInternal: boolean;
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
}
