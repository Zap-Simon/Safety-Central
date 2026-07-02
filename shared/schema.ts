import { pgTable, text, serial, integer, boolean, timestamp, uniqueIndex } from "drizzle-orm/pg-core";
import { sql } from "drizzle-orm";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

// Field Training (Able to Use) - 5-Level Competency System Constants
export const FIELD_COMPETENCY_LEVELS = {
  NOT_TRAINED: "Not Trained",
  IN_TRAINING_SUPERVISED: "In Training (Supervised)",
  COMPETENT_SUPERVISED: "Competent – Supervised", 
  COMPETENT_SOP_MODULE: "Competent – SOP/Module",
  EXPERT: "Expert"
} as const;

// Admin Training - Simplified 2-Level Competency System Constants
export const ADMIN_COMPETENCY_LEVELS = {
  NOT_TRAINED: "Not Trained",
  TRAINED: "Trained"
} as const;

// Legacy export for backward compatibility
export const COMPETENCY_LEVELS = FIELD_COMPETENCY_LEVELS;

export const FIELD_COMPETENCY_LEVEL_VALUES = Object.values(FIELD_COMPETENCY_LEVELS);
export const ADMIN_COMPETENCY_LEVEL_VALUES = Object.values(ADMIN_COMPETENCY_LEVELS);
export const COMPETENCY_LEVEL_VALUES = FIELD_COMPETENCY_LEVEL_VALUES; // Legacy

// Helper function to check if competency level allows "Able to Use" (Field Training only)
export function isAbleToUse(competencyLevel: string): boolean {
  return competencyLevel === FIELD_COMPETENCY_LEVELS.COMPETENT_SOP_MODULE || 
         competencyLevel === FIELD_COMPETENCY_LEVELS.EXPERT;
}

// Helper function to get competency levels based on training type
export function getCompetencyLevels(isFieldTraining: boolean) {
  return isFieldTraining ? FIELD_COMPETENCY_LEVEL_VALUES : ADMIN_COMPETENCY_LEVEL_VALUES;
}

export const users = pgTable("users", {
  id: serial("id").primaryKey(),
  username: text("username").notNull().unique(),
  password: text("password").notNull(),
});

export const meetingLocks = pgTable("meeting_locks", {
  id: serial("id").primaryKey(),
  meetingDate: text("meeting_date").notNull().unique(),
  isLocked: boolean("is_locked").notNull().default(false),
  isClosed: boolean("is_closed").notNull().default(false),
  lockedAt: timestamp("locked_at"),
  lockedBy: text("locked_by"),
  closedAt: timestamp("closed_at"),
  closedBy: text("closed_by"),
});

// Action Items - Local storage for action tracking data (linked to SharePoint items)
export const actionItems = pgTable("action_items", {
  id: serial("id").primaryKey(),
  listType: text("list_type").notNull(), // "BusinessIdeas", "SafetyIdeas", "NearMiss"
  sharePointItemId: text("sharepoint_item_id").notNull(), // SharePoint list item ID
  actionPriority: text("action_priority"), // "High", "Medium", "Low"
  actionStatus: text("action_status"), // "Not Started", "In Progress", "On Hold", "Completed"
  actionAssignedTo: text("action_assigned_to"),
  actionStartDate: text("action_start_date"), // ISO date string
  actionDueDate: text("action_due_date"), // ISO date string
  reconsiderDate: text("reconsider_date"), // ISO date — when an "On Hold" item should be revisited (re-surfaces in that meeting)
  actionNotes: text("action_notes"),
  meetingNotes: text("meeting_notes"), // Local backup of meeting notes (also saved to SharePoint)
  createdAt: timestamp("created_at").notNull().defaultNow(),
  updatedAt: timestamp("updated_at").notNull().defaultNow(),
}, (table) => ({
  uniqueItem: uniqueIndex("unique_action_item").on(table.listType, table.sharePointItemId)
}));

export const actionActivityLog = pgTable("action_activity_log", {
  id: serial("id").primaryKey(),
  listType: text("list_type").notNull(),
  sharePointItemId: text("sharepoint_item_id").notNull(),
  entryType: text("entry_type").notNull(), // 'note' | 'status' | 'priority' | 'due_date' | 'assigned' | 'start_date'
  content: text("content").notNull(),
  author: text("author"),
  createdAt: timestamp("created_at").notNull().defaultNow(),
});

export const insertActionActivityLogSchema = createInsertSchema(actionActivityLog).omit({
  id: true,
  createdAt: true,
});

export type ActionActivityLog = typeof actionActivityLog.$inferSelect;
export type InsertActionActivityLog = z.infer<typeof insertActionActivityLogSchema>;

export const meetingAttendance = pgTable("meeting_attendance", {
  id: serial("id").primaryKey(),
  meetingDate: text("meeting_date").notNull(),
  attendeeName: text("attendee_name").notNull(),
  isPresent: boolean("is_present").notNull().default(false),
  signatureStatus: text("signature_status"), // 'signed' | 'remote' | 'absent'
  signatureData: text("signature_data"), // base64 PNG data URL (for 'signed' status)
  signedAt: text("signed_at"), // ISO date string
  guestTitle: text("guest_title"), // free-text title/company for non-365 guests (null for roster members)
});

export const cardOrdering = pgTable("card_ordering", {
  id: serial("id").primaryKey(),
  cardId: text("card_id").notNull(),
  position: integer("position").notNull(),
  createdAt: timestamp("created_at").notNull().defaultNow(),
  updatedAt: timestamp("updated_at").notNull().defaultNow(),
});

// Tracks one-time data migrations that run on server startup so each runs once
// per database (dev and production each apply it the first time they boot).
export const appMigrations = pgTable("app_migrations", {
  id: serial("id").primaryKey(),
  key: text("key").notNull().unique(),
  appliedAt: timestamp("applied_at").notNull().defaultNow(),
});

// Small shared key-value settings store (e.g. the planned date of the next
// meeting when no upcoming meeting exists yet). Shared across all users.
export const appSettings = pgTable("app_settings", {
  key: text("key").primaryKey(),
  value: text("value").notNull(),
});

export type AppSetting = typeof appSettings.$inferSelect;

export const insertUserSchema = createInsertSchema(users).pick({
  username: true,
  password: true,
});

export const insertMeetingLockSchema = createInsertSchema(meetingLocks).pick({
  meetingDate: true,
  isLocked: true,
  isClosed: true,
  lockedBy: true,
  closedBy: true,
});

export const insertActionItemSchema = createInsertSchema(actionItems).omit({
  id: true,
  createdAt: true,
  updatedAt: true,
});

export const insertMeetingAttendanceSchema = createInsertSchema(meetingAttendance).pick({
  meetingDate: true,
  attendeeName: true,
  isPresent: true,
});

export const insertCardOrderingSchema = createInsertSchema(cardOrdering).pick({
  cardId: true,
  position: true,
});

export type InsertUser = z.infer<typeof insertUserSchema>;
export type User = typeof users.$inferSelect;
export type MeetingLock = typeof meetingLocks.$inferSelect;
export type InsertMeetingLock = z.infer<typeof insertMeetingLockSchema>;
export type ActionItem = typeof actionItems.$inferSelect;
export type InsertActionItem = z.infer<typeof insertActionItemSchema>;
export type MeetingAttendance = typeof meetingAttendance.$inferSelect;
export type InsertMeetingAttendance = z.infer<typeof insertMeetingAttendanceSchema>;
export type CardOrdering = typeof cardOrdering.$inferSelect;
export type InsertCardOrdering = z.infer<typeof insertCardOrderingSchema>;

// Excel integration types
export const worksheetDataSchema = z.object({
  name: z.string(),
  headers: z.array(z.string()),
  rows: z.array(z.array(z.union([z.string(), z.number(), z.null()]))),
  types: z.array(z.string()).optional(),
});

export const workbookDataSchema = z.object({
  fileId: z.string(),
  name: z.string(),
  lastModified: z.string(),
  sheets: z.array(worksheetDataSchema),
});

export const sharePointExcelFileSchema = z.object({
  id: z.string(),
  name: z.string(),
  size: z.number(),
  lastModified: z.string(),
  webUrl: z.string(),
  type: z.enum(['xlsx', 'xls', 'csv']),
});

export type WorksheetData = z.infer<typeof worksheetDataSchema>;
export type WorkbookData = z.infer<typeof workbookDataSchema>;
export type SharePointExcelFile = z.infer<typeof sharePointExcelFileSchema>;

// Skills Matrix Tables - Replacing Excel-based system

// Staff table - Central staff information linked to Azure AD
export const staff = pgTable("staff", {
  id: serial("id").primaryKey(),
  azureAdObjectId: text("azure_ad_object_id").unique(), // Link to Azure AD
  name: text("name").notNull(),
  email: text("email").notNull().unique(),
  jobTitle: text("job_title"),
  role: text("role").notNull().default("field"), // "admin", "supervisor", "field", "apprentice"
  roleRank: integer("role_rank").notNull().default(3), // 1=admin, 2=supervisor, 3=field, 4=apprentice
  isFieldStaff: boolean("is_field_staff").notNull().default(true), // Can access Field/Able to Use training
  isAdministrationStaff: boolean("is_administration_staff").notNull().default(false), // Can access Administration training
  startDate: timestamp("start_date"),
  isActive: boolean("is_active").notNull().default(true),
  createdAt: timestamp("created_at").notNull().defaultNow(),
  updatedAt: timestamp("updated_at").notNull().defaultNow(),
});

// Skills/Competencies definitions
export const skills = pgTable("skills", {
  id: serial("id").primaryKey(),
  category: text("category").notNull(), // "Heights", "Glass & Glazing", "PPE", "Equipment"
  name: text("name").notNull(),
  description: text("description"),
  requiresCertification: boolean("requires_certification").notNull().default(false),
  validityPeriod: integer("validity_period_months"), // How many months before renewal needed
  isActive: boolean("is_active").notNull().default(true),
  createdAt: timestamp("created_at").notNull().defaultNow(),
});

// Training Records - Staff x Training Module relationships with competency levels
export const trainingRecords = pgTable("training_records", {
  id: serial("id").primaryKey(),
  staffId: integer("staff_id").notNull().references(() => staff.id),
  skillId: integer("skill_id").notNull().references(() => trainingModules.id), // Updated to reference training_modules
  competencyLevel: text("competency_level").notNull(), // "Trainee", "Competent", "Instructor"
  achievedDate: timestamp("achieved_date"),
  appliedDate: timestamp("applied_date"), // When training was applied/started
  expiryDate: timestamp("expiry_date"),
  assessorName: text("assessor_name"),
  trainingProvider: text("training_provider"),
  certificateNumber: text("certificate_number"),
  notes: text("notes"),
  photoEvidenceUrl: text("photo_evidence_url"), // URL to photo evidence in SharePoint
  photoEvidenceFilename: text("photo_evidence_filename"), // Original filename
  ableToUse: boolean("able_to_use").notNull().default(false), // H&S Policy "Able to Use" status
  status: text("status").notNull().default("Active"), // "Active", "Expired", "Pending", "Superseded"
  previousRecordId: integer("previous_record_id").references(() => trainingRecords.id, { onDelete: "set null" }), // Versioning system
  version: integer("version").notNull().default(1), // Version number for audit trail
  createdAt: timestamp("created_at").notNull().defaultNow(),
  updatedAt: timestamp("updated_at").notNull().defaultNow(),
}, (table) => ({
  // Partial unique index: only one Active record per staff_id + skill_id
  uniqueActiveRecord: uniqueIndex("unique_active_training_record")
    .on(table.staffId, table.skillId)
    .where(sql`status = 'Active'`)
}));

// PPE Register - Tracking PPE issue and expiry
export const ppeRecords = pgTable("ppe_records", {
  id: serial("id").primaryKey(),
  staffId: integer("staff_id").notNull().references(() => staff.id),
  ppeType: text("ppe_type").notNull(), // "Hard Hat", "Hi-Vis Vest", "Safety Glasses", etc.
  brand: text("brand"),
  size: text("size"),
  issueDate: timestamp("issue_date"),
  expiryDate: timestamp("expiry_date"),
  condition: text("condition"), // "Good", "Fair", "Needs Replacement"
  location: text("location"), // Where issued/stored
  status: text("status").notNull().default("Issued"), // "Issued", "Expired", "Lost", "Returned"
  notes: text("notes"),
  createdAt: timestamp("created_at").notNull().defaultNow(),
  updatedAt: timestamp("updated_at").notNull().defaultNow(),
});

// Induction Register - Tracking completed inductions
export const inductionRecords = pgTable("induction_records", {
  id: serial("id").primaryKey(),
  staffId: integer("staff_id").notNull().references(() => staff.id),
  inductionType: text("induction_type").notNull(), // "Site Safety", "Company Policies", etc.
  completionDate: timestamp("completion_date").notNull(),
  location: text("location"),
  conductedBy: text("conducted_by"),
  expiryDate: timestamp("expiry_date"), // Some inductions may need renewal
  certificateIssued: boolean("certificate_issued").notNull().default(false),
  notes: text("notes"),
  status: text("status").notNull().default("Completed"), // "Completed", "Expired", "Pending"
  createdAt: timestamp("created_at").notNull().defaultNow(),
});

// Equipment Authorizations - "Able to Use" matrix (Enhanced with module references)
export const equipmentAuthorizations = pgTable("equipment_authorizations", {
  id: serial("id").primaryKey(),
  staffId: integer("staff_id").notNull().references(() => staff.id),
  moduleId: integer("module_id").references(() => trainingModules.id), // Link to training module
  moduleToolId: integer("module_tool_id").references(() => moduleTools.id), // Link to specific tool/equipment
  equipmentType: text("equipment_type").notNull(), // "Scaffold", "Ladder", "Cutting Tools", etc.
  equipmentModel: text("equipment_model"),
  authorizedDate: timestamp("authorized_date").notNull(),
  expiryDate: timestamp("expiry_date"),
  authorizedBy: text("authorized_by").notNull(),
  competencyLevel: text("competency_level"), // "Basic", "Advanced", "Instructor"
  restrictions: text("restrictions"), // Any usage restrictions
  status: text("status").notNull().default("Active"), // "Active", "Expired", "Suspended"
  createdAt: timestamp("created_at").notNull().defaultNow(),
  updatedAt: timestamp("updated_at").notNull().defaultNow(),
});

// Photo Assets - Staff ID photos stored in SharePoint
export const photoAssets = pgTable("photo_assets", {
  id: serial("id").primaryKey(),
  staffId: integer("staff_id").notNull().references(() => staff.id),
  filename: text("filename").notNull(),
  sharePointDriveItemId: text("sharepoint_drive_item_id").unique(),
  sharePointWebUrl: text("sharepoint_web_url"),
  thumbnailUrl: text("thumbnail_url"),
  fileSize: integer("file_size"),
  mimeType: text("mime_type"),
  uploadedAt: timestamp("uploaded_at").notNull().defaultNow(),
  uploadedBy: text("uploaded_by"),
  isActive: boolean("is_active").notNull().default(true),
});

// Glove Register - Tracking safety glove distribution
export const gloveRecords = pgTable("glove_records", {
  id: serial("id").primaryKey(),
  staffId: integer("staff_id").notNull().references(() => staff.id),
  gloveType: text("glove_type").notNull(), // "GV28", "GV35", etc. from your image
  gloveSize: text("glove_size").notNull(), // "S", "M", "L", "XL", "XXL"
  issueDate: timestamp("issue_date").notNull(),
  condition: text("condition").notNull(), // From your image: "Good condition", "Needs replacement"
  returnDate: timestamp("return_date"),
  reasonForReturn: text("reason_for_return"),
  status: text("status").notNull().default("Issued"), // "Issued", "Returned", "Lost"
  notes: text("notes"),
  createdAt: timestamp("created_at").notNull().defaultNow(),
});

// Insert schemas for all new tables
export const insertStaffSchema = createInsertSchema(staff).omit({
  id: true,
  createdAt: true,
  updatedAt: true,
});

export const insertSkillSchema = createInsertSchema(skills).omit({
  id: true,
  createdAt: true,
});

export const insertTrainingRecordSchema = createInsertSchema(trainingRecords).omit({
  id: true,
  createdAt: true,
  updatedAt: true,
  version: true, // Auto-managed by versioning system
}).extend({
  // Coerce date strings to Date objects
  achievedDate: z.coerce.date().nullable().optional(),
  appliedDate: z.coerce.date().nullable().optional(),
  expiryDate: z.coerce.date().nullable().optional(),
  // Allow previousRecordId for versioning system (optional)
  previousRecordId: z.number().nullable().optional(),
});

export const insertPpeRecordSchema = createInsertSchema(ppeRecords).omit({
  id: true,
  createdAt: true,
  updatedAt: true,
});

export const insertInductionRecordSchema = createInsertSchema(inductionRecords).omit({
  id: true,
  createdAt: true,
});

export const insertEquipmentAuthorizationSchema = createInsertSchema(equipmentAuthorizations).omit({
  id: true,
  createdAt: true,
  updatedAt: true,
});

export const insertPhotoAssetSchema = createInsertSchema(photoAssets).omit({
  id: true,
  uploadedAt: true,
});

export const insertGloveRecordSchema = createInsertSchema(gloveRecords).omit({
  id: true,
  createdAt: true,
});

// "Able to Use" Training System - Module-centric approach (Table Definitions First)

// Training Classifications (Foundation, Equipment, Safety Critical, Operations)
export const trainingClassifications = pgTable("training_classifications", {
  id: serial("id").primaryKey(),
  key: text("key").notNull().unique(), // "foundation", "equipment", "safety_critical", "operations"
  name: text("name").notNull(), // "Foundation", "Equipment", "Safety Critical", "Operations"
  description: text("description"),
  audience: text("audience").notNull().default("both"), // "field", "administration", "both"
  displayOrder: integer("display_order").notNull().default(0),
  isActive: boolean("is_active").notNull().default(true),
  createdAt: timestamp("created_at").notNull().defaultNow(),
});

// Training Modules (Basic Hand Tools, Power Tools, Lifting Equipment, etc.)
export const trainingModules = pgTable("training_modules", {
  id: serial("id").primaryKey(),
  classificationId: integer("classification_id").notNull().references(() => trainingClassifications.id),
  code: text("code").notNull().unique(), // "BHT001", "PWR002", "LFT003"
  name: text("name").notNull(), // "Basic Hand Tools", "Power Tools", "Lifting Equipment"
  description: text("description"),
  audience: text("audience").notNull().default("both"), // "field", "administration", "both"
  validityMonths: integer("validity_months"), // How many months before renewal needed
  requiresAssessment: boolean("requires_assessment").notNull().default(true),
  requiresAuthorization: boolean("requires_authorization").notNull().default(true),
  isSafetyCritical: boolean("is_safety_critical").notNull().default(false), // Requires formal certification and periodic reassessment
  requiresCertification: boolean("requires_certification").notNull().default(false), // Formal certification needed beyond training
  sopUrl: text("sop_url"), // Link to Standard Operating Procedure document
  trainingVideoUrl: text("training_video_url"), // Link to training video/materials
  displayOrder: integer("display_order").notNull().default(0),
  isActive: boolean("is_active").notNull().default(true),
  createdAt: timestamp("created_at").notNull().defaultNow(),
  updatedAt: timestamp("updated_at").notNull().defaultNow(),
});

// Module Prerequisites - Define which modules must be completed first
export const modulePrerequisites = pgTable("module_prerequisites", {
  id: serial("id").primaryKey(),
  moduleId: integer("module_id").notNull().references(() => trainingModules.id),
  prerequisiteModuleId: integer("prerequisite_module_id").notNull().references(() => trainingModules.id),
  createdAt: timestamp("created_at").notNull().defaultNow(),
});

// Module Tools/Equipment - Specific tools and equipment covered in each module
export const moduleTools = pgTable("module_tools", {
  id: serial("id").primaryKey(),
  moduleId: integer("module_id").notNull().references(() => trainingModules.id),
  toolName: text("tool_name").notNull(), // "Hammer", "Drill", "Angle Grinder"
  equipmentType: text("equipment_type"), // "Hand Tool", "Power Tool", "Safety Equipment"
  model: text("model"), // Specific model/brand if applicable
  description: text("description"),
  displayOrder: integer("display_order").notNull().default(0),
  isActive: boolean("is_active").notNull().default(true),
  createdAt: timestamp("created_at").notNull().defaultNow(),
});

// Module Materials - SOPs, training videos, checklists linked to SharePoint
export const moduleMaterials = pgTable("module_materials", {
  id: serial("id").primaryKey(),
  moduleId: integer("module_id").notNull().references(() => trainingModules.id),
  type: text("type").notNull(), // "SOP", "Video", "Checklist", "Document", "Assessment"
  title: text("title").notNull(),
  description: text("description"),
  sharePointDriveItemId: text("sharepoint_drive_item_id"), // Link to SharePoint document
  sharePointWebUrl: text("sharepoint_web_url"), // Full SharePoint URL
  externalUrl: text("external_url"), // For external training materials
  displayOrder: integer("display_order").notNull().default(0),
  isActive: boolean("is_active").notNull().default(true),
  createdAt: timestamp("created_at").notNull().defaultNow(),
  updatedAt: timestamp("updated_at").notNull().defaultNow(),
});

// Staff Module Progress - Tracks progress through training modules
export const staffModuleProgress = pgTable("staff_module_progress", {
  id: serial("id").primaryKey(),
  staffId: integer("staff_id").notNull().references(() => staff.id),
  moduleId: integer("module_id").notNull().references(() => trainingModules.id),
  status: text("status").notNull().default("Not Started"), // "Not Started", "In Training", "Assessed-Pass", "Assessed-Fail", "Authorized"
  competencyLevel: text("competency_level").notNull().default("Not Trained"), // New 5-level system
  trainedAgainstSop: boolean("trained_against_sop").notNull().default(false), // Whether training included SOP/module materials
  sopAcknowledgedAt: timestamp("sop_acknowledged_at"), // When SOP was acknowledged
  sopVersion: text("sop_version"), // Version of SOP used in training
  appliedDate: timestamp("applied_date"), // When training started
  completedDate: timestamp("completed_date"), // When training/assessment completed
  authorizedDate: timestamp("authorized_date"), // When authorized for use
  expiryDate: timestamp("expiry_date"), // When authorization expires
  assessorName: text("assessor_name"),
  assessorStaffId: integer("assessor_staff_id").references(() => staff.id),
  evidenceUrl: text("evidence_url"), // URL to evidence files in SharePoint
  evidenceFilename: text("evidence_filename"),
  notes: text("notes"),
  createdAt: timestamp("created_at").notNull().defaultNow(),
  updatedAt: timestamp("updated_at").notNull().defaultNow(),
});

// Insert schemas for "Able to Use" training system tables
export const insertTrainingClassificationSchema = createInsertSchema(trainingClassifications).omit({
  id: true,
  createdAt: true,
});

export const insertTrainingModuleSchema = createInsertSchema(trainingModules).omit({
  id: true,
  createdAt: true,
  updatedAt: true,
});

export const insertModulePrerequisiteSchema = createInsertSchema(modulePrerequisites).omit({
  id: true,
  createdAt: true,
});

export const insertModuleToolSchema = createInsertSchema(moduleTools).omit({
  id: true,
  createdAt: true,
});

export const insertModuleMaterialSchema = createInsertSchema(moduleMaterials).omit({
  id: true,
  createdAt: true,
  updatedAt: true,
});

export const insertStaffModuleProgressSchema = createInsertSchema(staffModuleProgress).omit({
  id: true,
  createdAt: true,
  updatedAt: true,
}).extend({
  competencyLevel: z.enum([
    // Field training levels
    FIELD_COMPETENCY_LEVELS.NOT_TRAINED,
    FIELD_COMPETENCY_LEVELS.IN_TRAINING_SUPERVISED,
    FIELD_COMPETENCY_LEVELS.COMPETENT_SUPERVISED,
    FIELD_COMPETENCY_LEVELS.COMPETENT_SOP_MODULE,
    FIELD_COMPETENCY_LEVELS.EXPERT,
    // Admin training levels
    ADMIN_COMPETENCY_LEVELS.NOT_TRAINED, // Same as field "Not Trained"
    ADMIN_COMPETENCY_LEVELS.TRAINED
  ]).default(FIELD_COMPETENCY_LEVELS.NOT_TRAINED)
});

// Update equipment authorization schema to include new fields
export const insertEquipmentAuthorizationSchemaNew = createInsertSchema(equipmentAuthorizations).omit({
  id: true,
  createdAt: true,
  updatedAt: true,
});

// Type exports for all new tables
export type Staff = typeof staff.$inferSelect;
export type InsertStaff = z.infer<typeof insertStaffSchema>;
export type Skill = typeof skills.$inferSelect;
export type InsertSkill = z.infer<typeof insertSkillSchema>;
export type TrainingRecord = typeof trainingRecords.$inferSelect;
export type InsertTrainingRecord = z.infer<typeof insertTrainingRecordSchema>;
export type PpeRecord = typeof ppeRecords.$inferSelect;
export type InsertPpeRecord = z.infer<typeof insertPpeRecordSchema>;
export type InductionRecord = typeof inductionRecords.$inferSelect;
export type InsertInductionRecord = z.infer<typeof insertInductionRecordSchema>;
export type EquipmentAuthorization = typeof equipmentAuthorizations.$inferSelect;
export type InsertEquipmentAuthorization = z.infer<typeof insertEquipmentAuthorizationSchema>;
export type PhotoAsset = typeof photoAssets.$inferSelect;
export type InsertPhotoAsset = z.infer<typeof insertPhotoAssetSchema>;
export type GloveRecord = typeof gloveRecords.$inferSelect;
export type InsertGloveRecord = z.infer<typeof insertGloveRecordSchema>;

// CSC Workflow Modules - Field Service Management System
export const CSC_WORKFLOW_MODULES = {
  CSC_WF_001: {
    code: "CSC-WF-001",
    name: "Direct Client Workflow",
    description: "Standard operating procedure for handling direct client jobs and service requests",
    classification: "office-procedures",
    audience: "administration"
  },
  CSC_WF_002: {
    code: "CSC-WF-002", 
    name: "Insurance Master Glaziers Workflow",
    description: "Procedures for insurance-related glazing work and master glazier requirements",
    classification: "office-procedures",
    audience: "administration"
  },
  CSC_WF_003: {
    code: "CSC-WF-003",
    name: "Government & Builders Workflow", 
    description: "Compliance and procedures for government contracts and builder partnerships",
    classification: "office-procedures",
    audience: "administration"
  },
  CSC_WF_004: {
    code: "CSC-WF-004",
    name: "Property Management Workflow",
    description: "Standard procedures for property management company contracts and maintenance work",
    classification: "office-procedures", 
    audience: "administration"
  }
} as const;

// CSC Workflow Data Schema for popup display
export const cscWorkflowSchema = z.object({
  code: z.string(),
  name: z.string(),
  description: z.string(),
  content: z.string().optional(),
  steps: z.array(z.object({
    stepNumber: z.number(),
    title: z.string(),
    description: z.string(),
    requirements: z.array(z.string()).optional(),
    notes: z.string().optional()
  })).optional(),
  requirements: z.array(z.string()).optional(),
  documentation: z.array(z.object({
    title: z.string(),
    url: z.string().optional(),
    type: z.enum(["SOP", "Form", "Checklist", "Reference"])
  })).optional(),
  contacts: z.array(z.object({
    role: z.string(),
    name: z.string().optional(),
    phone: z.string().optional(),
    email: z.string().optional()
  })).optional()
});

// Type exports for "Able to Use" training system
export type TrainingClassification = typeof trainingClassifications.$inferSelect;
export type InsertTrainingClassification = z.infer<typeof insertTrainingClassificationSchema>;
export type TrainingModule = typeof trainingModules.$inferSelect;
export type InsertTrainingModule = z.infer<typeof insertTrainingModuleSchema>;
export type ModulePrerequisite = typeof modulePrerequisites.$inferSelect;
export type InsertModulePrerequisite = z.infer<typeof insertModulePrerequisiteSchema>;
export type ModuleTool = typeof moduleTools.$inferSelect;
export type InsertModuleTool = z.infer<typeof insertModuleToolSchema>;
export type ModuleMaterial = typeof moduleMaterials.$inferSelect;
export type InsertModuleMaterial = z.infer<typeof insertModuleMaterialSchema>;
export type StaffModuleProgress = typeof staffModuleProgress.$inferSelect;
export type InsertStaffModuleProgress = z.infer<typeof insertStaffModuleProgressSchema>;

// CSC Workflow types
export type CSCWorkflow = z.infer<typeof cscWorkflowSchema>;

// ─── Near Miss Investigations ───────────────────────────────────────────────
export const nearMissInvestigations = pgTable("near_miss_investigations", {
  id: serial("id").primaryKey(),
  nearMissItemId: text("near_miss_item_id").notNull(),
  itemTitle: text("item_title").notNull().default(""),
  meetingDate: text("meeting_date").notNull().default(""),
  investigatorName: text("investigator_name").notNull().default(""),
  siteJob: text("site_job").notNull().default(""),
  eventDate: text("event_date").notNull().default(""),
  eventTime: text("event_time").notNull().default(""),
  eventType: text("event_type").notNull().default("Near Miss"),
  involvedPersons: text("involved_persons").notNull().default(""),
  witnesses: text("witnesses").notNull().default(""),
  eventDescription: text("event_description").notNull().default(""),
  contributingFactors: text("contributing_factors").notNull().default(""),
  hazards: text("hazards").notNull().default("[]"),
  likelihood: text("likelihood").notNull().default(""),
  consequence: text("consequence").notNull().default(""),
  riskLevel: text("risk_level").notNull().default(""),
  treatmentGiven: text("treatment_given").notNull().default(""),
  resultingActions: text("resulting_actions").notNull().default("[]"),
  // Investigator sign-off
  investigatorSignature: text("investigator_signature"),
  investigatorSignedAt: text("investigator_signed_at"),
  // Approver / Manager sign-off (legacy column names kept to preserve existing data)
  directorName: text("director_name"),
  directorSignature: text("director_signature"),
  signedAt: text("signed_at"),
  status: text("status").notNull().default("Draft"),
  createdAt: timestamp("created_at").notNull().defaultNow(),
  updatedAt: timestamp("updated_at").notNull().defaultNow(),
});

export const insertNearMissInvestigationSchema = createInsertSchema(nearMissInvestigations).omit({
  id: true,
  createdAt: true,
  updatedAt: true,
});

export type NearMissInvestigation = typeof nearMissInvestigations.$inferSelect;
export type InsertNearMissInvestigation = z.infer<typeof insertNearMissInvestigationSchema>;

// ─── Investigation Progress Notes — time-stamped history while an investigation is in progress ───
export const investigationProgressNotes = pgTable("investigation_progress_notes", {
  id: serial("id").primaryKey(),
  nearMissItemId: text("near_miss_item_id").notNull(),
  content: text("content").notNull(),
  author: text("author"),
  createdAt: timestamp("created_at").notNull().defaultNow(),
});

export const insertInvestigationProgressNoteSchema = createInsertSchema(investigationProgressNotes).omit({
  id: true,
  createdAt: true,
});

export type InvestigationProgressNote = typeof investigationProgressNotes.$inferSelect;
export type InsertInvestigationProgressNote = z.infer<typeof insertInvestigationProgressNoteSchema>;

// ─── Order Items — Teams Whiteboard field ordering pad ───────────────────────
export const orderItems = pgTable("order_items", {
  id: serial("id").primaryKey(),
  itemName: text("item_name").notNull(),
  addedBy: text("added_by").notNull(),
  addedAt: timestamp("added_at").notNull().defaultNow(),
  status: text("status").notNull().default("active"), // "active" | "ordered" | "archived"
  orderedAt: timestamp("ordered_at"),
  orderedBy: text("ordered_by"),
});

export const insertOrderItemSchema = createInsertSchema(orderItems).omit({
  id: true,
  addedAt: true,
  orderedAt: true,
  orderedBy: true,
});

export type OrderItem = typeof orderItems.$inferSelect;
export type InsertOrderItem = z.infer<typeof insertOrderItemSchema>;

// Enhanced equipment authorization type with module references
export type EquipmentAuthorizationNew = typeof equipmentAuthorizations.$inferSelect;
export type InsertEquipmentAuthorizationNew = z.infer<typeof insertEquipmentAuthorizationSchemaNew>;
