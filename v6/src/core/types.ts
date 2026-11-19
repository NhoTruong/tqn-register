// Kiểu dữ liệu nghiệp vụ V6 – cùng tên trường với db/schema_v0.sql để sau chuyển PostgreSQL không đổi nghiệp vụ.
export type DocType = 'RFI' | 'RFA' | 'RFC' | 'SDS' | 'MSS' | 'MAS';
export type Role = 'ADMIN' | 'DC' | 'CONTRACTOR' | 'REVIEWER' | 'VIEWER';
export type ActorKind = 'HUMAN' | 'AI' | 'SYSTEM';

export interface Actor { email: string; kind: ActorKind }

export interface ReplyOption { code: string; text: string; result: 'CLOSE' | 'NEED_INFO' | 'FORWARD' | 'RESUBMIT' | 'REJECT' }

export interface TypeConfig {
  pattern: string;          // vd "{yymmdd}-{project}-{contractor}-{type}-{seq:3}"
  slaDays: number;          // số ngày bên duyệt phải trả lời
  reviewer: string;         // nhãn bên duyệt (CĐT)
  options: ReplyOption[];   // đúng chữ trên phiếu của dự án
}

export interface ProjectConfig {
  projectId: string;
  code: string;             // COA
  contractor: string;       // TBC
  timezone: string;
  configVersion: number;
  members: Record<string, Role[]>;   // email (chữ thường) → vai trò
  types: Partial<Record<DocType, TypeConfig>>;
}

export type DocStatus = 'DRAFT' | 'SUBMITTED' | 'NEED_INFO' | 'FORWARDED' | 'CLOSED' | 'REJECTED' | 'RESUBMIT_REQUIRED';

export interface FileRef { name: string; url: string; source: 'DRIVE' | 'ONEDRIVE' | 'GMAIL' | 'LINK'; sha256?: string }

export interface DocumentRec {
  documentId: string;       // ULID – khóa nghiệp vụ, không phải số dòng
  projectId: string;
  docType: DocType;
  docNo: string;            // mã phát hành do Numbering Engine cấp
  title: string;
  status: DocStatus;
  currentRev: number;
  owner: string;            // email người phụ trách (bên nhà thầu)
  holder: string;           // bên đang giữ việc: CONTRACTOR | REVIEWER label
  dueAt: string | null;     // ISO
  createdAt: string;
  updatedAt: string;
}

export interface RevisionRec {
  documentId: string;
  revNo: number;
  submittedAt: string | null;
  repliedAt: string | null;
  replyCode: string | null;
  replyText: string | null;
  replyNote: string | null;
  files: FileRef[];
}

export interface EventRec {
  eventId: string;
  at: string;
  actor: string;
  actorKind: ActorKind;
  action: string;
  documentId: string | null;
  revNo: number | null;
  idemKey: string | null;
  detail: string;           // JSON ngắn
  ok: boolean;
}
