import { 
  users, 
  meetingLocks, 
  meetingAttendance,
  actionItems,
  actionActivityLog,
  staff,
  skills,
  trainingRecords,
  ppeRecords,
  inductionRecords,
  equipmentAuthorizations,
  photoAssets,
  gloveRecords,
  trainingClassifications,
  trainingModules,
  modulePrerequisites,
  moduleTools,
  moduleMaterials,
  staffModuleProgress,
  orderItems,
  type User, 
  type InsertUser,
  type MeetingLock,
  type InsertMeetingLock,
  type ActionItem,
  type InsertActionItem,
  type MeetingAttendance,
  type InsertMeetingAttendance,
  type Staff,
  type InsertStaff,
  type Skill,
  type InsertSkill,
  type TrainingRecord,
  type InsertTrainingRecord,
  type PpeRecord,
  type InsertPpeRecord,
  type InductionRecord,
  type InsertInductionRecord,
  type EquipmentAuthorization,
  type InsertEquipmentAuthorization,
  type PhotoAsset,
  type InsertPhotoAsset,
  type GloveRecord,
  type InsertGloveRecord,
  type TrainingClassification,
  type InsertTrainingClassification,
  type TrainingModule,
  type InsertTrainingModule,
  type ModulePrerequisite,
  type InsertModulePrerequisite,
  type ModuleTool,
  type InsertModuleTool,
  type ModuleMaterial,
  type InsertModuleMaterial,
  type StaffModuleProgress,
  type InsertStaffModuleProgress,
  type OrderItem,
  type InsertOrderItem,
  type ActionActivityLog,
  type InsertActionActivityLog,
} from "@shared/schema";
import { db } from "./db";
import { eq, and, gte, lte } from "drizzle-orm";

// modify the interface with any CRUD methods
// you might need

export interface IStorage {
  getUser(id: number): Promise<User | undefined>;
  getUserByUsername(username: string): Promise<User | undefined>;
  createUser(user: InsertUser): Promise<User>;
  
  // Meeting lock methods
  getMeetingLock(meetingDate: string): Promise<MeetingLock | undefined>;
  setMeetingLock(lock: InsertMeetingLock): Promise<MeetingLock>;
  updateMeetingLock(meetingDate: string, isLocked: boolean, lockedBy?: string): Promise<MeetingLock>;
  updateMeetingClosed(meetingDate: string, isClosed: boolean, closedBy?: string): Promise<MeetingLock>;
  
  // Meeting attendance methods
  getMeetingAttendance(meetingDate: string): Promise<MeetingAttendance[]>;
  setMeetingAttendance(attendance: InsertMeetingAttendance): Promise<MeetingAttendance>;
  updateMeetingAttendance(meetingDate: string, attendeeName: string, isPresent: boolean): Promise<MeetingAttendance>;
  getAllMeetingAttendance(): Promise<Record<string, string[]>>;
  updateMeetingSignature(meetingDate: string, attendeeName: string, status: string, signatureData: string | null, signedAt: string): Promise<MeetingAttendance>;
  getMeetingSignatures(meetingDate: string): Promise<Record<string, { status: string; signatureData: string | null; signedAt: string }>>;
  getAllMeetingSignatures(): Promise<Record<string, Record<string, { status: string; signatureData: string | null; signedAt: string }>>>;
  
  // Action Items methods (local database for action tracking)
  getActionItem(listType: string, sharePointItemId: string): Promise<ActionItem | undefined>;
  getAllActionItems(): Promise<ActionItem[]>;
  getActionItemsByListType(listType: string): Promise<ActionItem[]>;
  upsertActionItem(item: InsertActionItem): Promise<ActionItem>;

  // Action Activity Log methods
  addActivityEntry(entry: InsertActionActivityLog): Promise<ActionActivityLog>;
  getActivityLog(listType: string, sharePointItemId: string): Promise<ActionActivityLog[]>;
  
  // Skills Matrix methods
  
  // Staff methods
  getAllStaff(): Promise<Staff[]>;
  getStaff(id: number): Promise<Staff | undefined>;
  getStaffByAzureId(azureAdObjectId: string): Promise<Staff | undefined>;
  getStaffByEmail(email: string): Promise<Staff | undefined>;
  createStaff(staff: InsertStaff): Promise<Staff>;
  updateStaff(id: number, updates: Partial<InsertStaff>): Promise<Staff>;
  deleteStaff(id: number): Promise<void>;
  
  // Skills methods
  getAllSkills(): Promise<Skill[]>;
  getSkill(id: number): Promise<Skill | undefined>;
  getSkillsByCategory(category: string): Promise<Skill[]>;
  createSkill(skill: InsertSkill): Promise<Skill>;
  updateSkill(id: number, updates: Partial<InsertSkill>): Promise<Skill>;
  
  // Training Records methods
  getAllTrainingRecords(): Promise<TrainingRecord[]>;
  getTrainingRecord(id: number): Promise<TrainingRecord | undefined>;
  getTrainingRecordsByStaff(staffId: number): Promise<TrainingRecord[]>;
  getTrainingRecordsBySkill(skillId: number): Promise<TrainingRecord[]>;
  getTrainingMatrix(): Promise<Array<{ staffId: number; staffName: string; skillId: number; skillName: string; competencyLevel: string; status: string; expiryDate: Date | null }>>;
  createTrainingRecord(record: InsertTrainingRecord): Promise<TrainingRecord>;
  updateTrainingRecord(id: number, updates: Partial<InsertTrainingRecord>): Promise<TrainingRecord>;
  
  // PPE Records methods
  getAllPpeRecords(): Promise<PpeRecord[]>;
  getPpeRecord(id: number): Promise<PpeRecord | undefined>;
  getPpeRecordsByStaff(staffId: number): Promise<PpeRecord[]>;
  getExpiringPpe(withinDays: number): Promise<PpeRecord[]>;
  createPpeRecord(record: InsertPpeRecord): Promise<PpeRecord>;
  updatePpeRecord(id: number, updates: Partial<InsertPpeRecord>): Promise<PpeRecord>;
  
  // Induction Records methods
  getAllInductionRecords(): Promise<InductionRecord[]>;
  getInductionRecord(id: number): Promise<InductionRecord | undefined>;
  getInductionRecordsByStaff(staffId: number): Promise<InductionRecord[]>;
  createInductionRecord(record: InsertInductionRecord): Promise<InductionRecord>;
  updateInductionRecord(id: number, updates: Partial<InsertInductionRecord>): Promise<InductionRecord>;
  
  // Equipment Authorization methods
  getAllEquipmentAuthorizations(): Promise<EquipmentAuthorization[]>;
  getEquipmentAuthorization(id: number): Promise<EquipmentAuthorization | undefined>;
  getEquipmentAuthorizationsByStaff(staffId: number): Promise<EquipmentAuthorization[]>;
  getEquipmentAuthorizationsByType(equipmentType: string): Promise<EquipmentAuthorization[]>;
  createEquipmentAuthorization(auth: InsertEquipmentAuthorization): Promise<EquipmentAuthorization>;
  updateEquipmentAuthorization(id: number, updates: Partial<InsertEquipmentAuthorization>): Promise<EquipmentAuthorization>;
  
  // Photo Asset methods
  getPhotoAsset(id: number): Promise<PhotoAsset | undefined>;
  getPhotoAssetByStaff(staffId: number): Promise<PhotoAsset | undefined>;
  createPhotoAsset(asset: InsertPhotoAsset): Promise<PhotoAsset>;
  updatePhotoAsset(id: number, updates: Partial<InsertPhotoAsset>): Promise<PhotoAsset>;
  deletePhotoAsset(id: number): Promise<void>;
  
  // Glove Records methods
  getAllGloveRecords(): Promise<GloveRecord[]>;
  getGloveRecord(id: number): Promise<GloveRecord | undefined>;
  getGloveRecordsByStaff(staffId: number): Promise<GloveRecord[]>;
  createGloveRecord(record: InsertGloveRecord): Promise<GloveRecord>;
  updateGloveRecord(id: number, updates: Partial<InsertGloveRecord>): Promise<GloveRecord>;
  
  // "Able to Use" Training System methods
  
  // Training Classification methods
  getAllTrainingClassifications(audience?: string): Promise<TrainingClassification[]>;
  getTrainingClassification(id: number): Promise<TrainingClassification | undefined>;
  getTrainingClassificationByKey(key: string): Promise<TrainingClassification | undefined>;
  createTrainingClassification(classification: InsertTrainingClassification): Promise<TrainingClassification>;
  updateTrainingClassification(id: number, updates: Partial<InsertTrainingClassification>): Promise<TrainingClassification>;
  
  // Training Module methods
  getAllTrainingModules(): Promise<TrainingModule[]>;
  getTrainingModule(id: number): Promise<TrainingModule | undefined>;
  getTrainingModulesByClassification(classificationId: number): Promise<TrainingModule[]>;
  getTrainingModuleByCode(code: string): Promise<TrainingModule | undefined>;
  createTrainingModule(module: InsertTrainingModule): Promise<TrainingModule>;
  updateTrainingModule(id: number, updates: Partial<InsertTrainingModule>): Promise<TrainingModule>;
  
  // Module Prerequisite methods
  getModulePrerequisites(moduleId: number): Promise<ModulePrerequisite[]>;
  createModulePrerequisite(prerequisite: InsertModulePrerequisite): Promise<ModulePrerequisite>;
  deleteModulePrerequisite(moduleId: number, prerequisiteModuleId: number): Promise<void>;
  
  // Module Tool methods
  getModuleTools(moduleId: number): Promise<ModuleTool[]>;
  createModuleTool(tool: InsertModuleTool): Promise<ModuleTool>;
  updateModuleTool(id: number, updates: Partial<InsertModuleTool>): Promise<ModuleTool>;
  deleteModuleTool(id: number): Promise<void>;
  
  // Module Material methods
  getModuleMaterials(moduleId: number): Promise<ModuleMaterial[]>;
  createModuleMaterial(material: InsertModuleMaterial): Promise<ModuleMaterial>;
  updateModuleMaterial(id: number, updates: Partial<InsertModuleMaterial>): Promise<ModuleMaterial>;
  deleteModuleMaterial(id: number): Promise<void>;
  
  // Staff Module Progress methods
  getAllStaffModuleProgress(): Promise<StaffModuleProgress[]>;
  getStaffModuleProgress(staffId: number): Promise<StaffModuleProgress[]>;
  getModuleProgressByStaffAndModule(staffId: number, moduleId: number): Promise<StaffModuleProgress | undefined>;
  getModuleProgress(moduleId: number): Promise<StaffModuleProgress[]>;
  createStaffModuleProgress(progress: InsertStaffModuleProgress): Promise<StaffModuleProgress>;
  updateStaffModuleProgress(id: number, updates: Partial<InsertStaffModuleProgress>): Promise<StaffModuleProgress>;
  
  // Complex queries for training matrix
  getTrainingModuleMatrix(): Promise<Array<{ 
    staffId: number; 
    staffName: string; 
    moduleId: number; 
    moduleName: string; 
    moduleCode: string;
    classificationName: string;
    status: string; 
    competencyLevel: string | null;
    completedDate: Date | null;
    expiryDate: Date | null;
  }>>;

  // Order Items methods
  getActiveOrderItems(): Promise<OrderItem[]>;
  createOrderItem(item: InsertOrderItem): Promise<OrderItem>;
  updateOrderItemStatus(id: number, status: string, orderedBy?: string): Promise<OrderItem>;
  clearActiveOrderItems(): Promise<number>;
}




export class DatabaseStorage implements IStorage {
  async getUser(id: number): Promise<User | undefined> {
    const [user] = await db.select().from(users).where(eq(users.id, id));
    return user || undefined;
  }

  async getUserByUsername(username: string): Promise<User | undefined> {
    const [user] = await db.select().from(users).where(eq(users.username, username));
    return user || undefined;
  }

  async createUser(insertUser: InsertUser): Promise<User> {
    const [user] = await db
      .insert(users)
      .values(insertUser)
      .returning();
    return user;
  }

  // Meeting lock methods
  async getMeetingLock(meetingDate: string): Promise<MeetingLock | undefined> {
    const [lock] = await db.select().from(meetingLocks).where(eq(meetingLocks.meetingDate, meetingDate));
    return lock || undefined;
  }

  async setMeetingLock(lock: InsertMeetingLock): Promise<MeetingLock> {
    const [meetingLock] = await db
      .insert(meetingLocks)
      .values({
        meetingDate: lock.meetingDate,
        isLocked: lock.isLocked ?? false,
        lockedAt: lock.isLocked ? new Date() : null,
        lockedBy: lock.lockedBy || null
      })
      .returning();
    return meetingLock;
  }

  async updateMeetingLock(meetingDate: string, isLocked: boolean, lockedBy?: string): Promise<MeetingLock> {
    const existing = await this.getMeetingLock(meetingDate);
    if (existing) {
      const [updated] = await db
        .update(meetingLocks)
        .set({
          isLocked,
          lockedAt: isLocked ? new Date() : null,
          lockedBy: lockedBy || null
        })
        .where(eq(meetingLocks.meetingDate, meetingDate))
        .returning();
      return updated;
    } else {
      return this.setMeetingLock({ meetingDate, isLocked, lockedBy });
    }
  }

  async updateMeetingClosed(meetingDate: string, isClosed: boolean, closedBy?: string): Promise<MeetingLock> {
    const existing = await this.getMeetingLock(meetingDate);
    if (existing) {
      const [updated] = await db
        .update(meetingLocks)
        .set({
          isClosed,
          closedAt: isClosed ? new Date() : null,
          closedBy: closedBy || null
        })
        .where(eq(meetingLocks.meetingDate, meetingDate))
        .returning();
      return updated;
    } else {
      // Create new record with just isClosed set
      const [meetingLock] = await db
        .insert(meetingLocks)
        .values({
          meetingDate,
          isLocked: false,
          isClosed,
          lockedAt: null,
          lockedBy: null,
          closedAt: isClosed ? new Date() : null,
          closedBy: closedBy || null
        })
        .returning();
      return meetingLock;
    }
  }

  // Meeting attendance methods
  async getMeetingAttendance(meetingDate: string): Promise<MeetingAttendance[]> {
    return await db.select().from(meetingAttendance).where(eq(meetingAttendance.meetingDate, meetingDate));
  }

  async setMeetingAttendance(attendance: InsertMeetingAttendance): Promise<MeetingAttendance> {
    const [record] = await db
      .insert(meetingAttendance)
      .values({
        meetingDate: attendance.meetingDate,
        attendeeName: attendance.attendeeName,
        isPresent: attendance.isPresent ?? false
      })
      .returning();
    return record;
  }

  async updateMeetingAttendance(meetingDate: string, attendeeName: string, isPresent: boolean): Promise<MeetingAttendance> {
    // Try to find existing record
    const existingRecords = await db
      .select()
      .from(meetingAttendance)
      .where(eq(meetingAttendance.meetingDate, meetingDate));
    
    const existing = existingRecords.find(record => record.attendeeName === attendeeName);

    if (existing) {
      const [updated] = await db
        .update(meetingAttendance)
        .set({ isPresent })
        .where(eq(meetingAttendance.id, existing.id))
        .returning();
      return updated;
    } else {
      return this.setMeetingAttendance({ meetingDate, attendeeName, isPresent });
    }
  }

  async getAllMeetingAttendance(): Promise<Record<string, string[]>> {
    const records = await db.select().from(meetingAttendance);
    const result: Record<string, string[]> = {};
    for (const record of records) {
      if (record.isPresent) {
        if (!result[record.meetingDate]) {
          result[record.meetingDate] = [];
        }
        result[record.meetingDate].push(record.attendeeName);
      }
    }
    return result;
  }

  async updateMeetingSignature(meetingDate: string, attendeeName: string, status: string, signatureData: string | null, signedAt: string): Promise<MeetingAttendance> {
    const existingRecords = await db
      .select()
      .from(meetingAttendance)
      .where(eq(meetingAttendance.meetingDate, meetingDate));
    
    const existing = existingRecords.find(r => r.attendeeName === attendeeName);

    if (existing) {
      const [updated] = await db
        .update(meetingAttendance)
        .set({ signatureStatus: status, signatureData, signedAt })
        .where(eq(meetingAttendance.id, existing.id))
        .returning();
      return updated;
    } else {
      const [created] = await db
        .insert(meetingAttendance)
        .values({ meetingDate, attendeeName, isPresent: true, signatureStatus: status, signatureData, signedAt })
        .returning();
      return created;
    }
  }

  async getMeetingSignatures(meetingDate: string): Promise<Record<string, { status: string; signatureData: string | null; signedAt: string }>> {
    const records = await db.select().from(meetingAttendance).where(eq(meetingAttendance.meetingDate, meetingDate));
    const result: Record<string, { status: string; signatureData: string | null; signedAt: string }> = {};
    for (const record of records) {
      if (record.signatureStatus) {
        result[record.attendeeName] = {
          status: record.signatureStatus,
          signatureData: record.signatureData ?? null,
          signedAt: record.signedAt ?? ''
        };
      }
    }
    return result;
  }

  async getAllMeetingSignatures(): Promise<Record<string, Record<string, { status: string; signatureData: string | null; signedAt: string }>>> {
    const records = await db.select().from(meetingAttendance);
    const result: Record<string, Record<string, { status: string; signatureData: string | null; signedAt: string }>> = {};
    for (const record of records) {
      if (record.signatureStatus) {
        if (!result[record.meetingDate]) {
          result[record.meetingDate] = {};
        }
        result[record.meetingDate][record.attendeeName] = {
          status: record.signatureStatus,
          signatureData: record.signatureData ?? null,
          signedAt: record.signedAt ?? ''
        };
      }
    }
    return result;
  }

  // Action Items methods (local database for action tracking)
  async getActionItem(listType: string, sharePointItemId: string): Promise<ActionItem | undefined> {
    const [item] = await db
      .select()
      .from(actionItems)
      .where(and(eq(actionItems.listType, listType), eq(actionItems.sharePointItemId, sharePointItemId)));
    return item || undefined;
  }

  async getAllActionItems(): Promise<ActionItem[]> {
    return await db.select().from(actionItems);
  }

  async getActionItemsByListType(listType: string): Promise<ActionItem[]> {
    return await db.select().from(actionItems).where(eq(actionItems.listType, listType));
  }

  async upsertActionItem(item: InsertActionItem): Promise<ActionItem> {
    const existing = await this.getActionItem(item.listType, item.sharePointItemId);
    
    if (existing) {
      const [updated] = await db
        .update(actionItems)
        .set({ 
          ...item,
          updatedAt: new Date() 
        })
        .where(eq(actionItems.id, existing.id))
        .returning();
      return updated;
    } else {
      const [created] = await db
        .insert(actionItems)
        .values(item)
        .returning();
      return created;
    }
  }

  // Action Activity Log methods
  async addActivityEntry(entry: InsertActionActivityLog): Promise<ActionActivityLog> {
    const [created] = await db
      .insert(actionActivityLog)
      .values(entry)
      .returning();
    return created;
  }

  async getActivityLog(listType: string, sharePointItemId: string): Promise<ActionActivityLog[]> {
    return await db
      .select()
      .from(actionActivityLog)
      .where(and(
        eq(actionActivityLog.listType, listType),
        eq(actionActivityLog.sharePointItemId, sharePointItemId)
      ))
      .orderBy(actionActivityLog.createdAt);
  }

  // Skills Matrix Methods
  
  // Staff methods
  async getAllStaff(): Promise<Staff[]> {
    return await db.select().from(staff);
  }

  async getStaff(id: number): Promise<Staff | undefined> {
    const [staffMember] = await db.select().from(staff).where(eq(staff.id, id));
    return staffMember || undefined;
  }

  async getStaffByAzureId(azureAdObjectId: string): Promise<Staff | undefined> {
    const [staffMember] = await db.select().from(staff).where(eq(staff.azureAdObjectId, azureAdObjectId));
    return staffMember || undefined;
  }

  async getStaffByEmail(email: string): Promise<Staff | undefined> {
    const [staffMember] = await db.select().from(staff).where(eq(staff.email, email));
    return staffMember || undefined;
  }

  async createStaff(insertStaff: InsertStaff): Promise<Staff> {
    const [newStaff] = await db
      .insert(staff)
      .values(insertStaff)
      .returning();
    return newStaff;
  }

  async updateStaff(id: number, updates: Partial<InsertStaff>): Promise<Staff> {
    const [updatedStaff] = await db
      .update(staff)
      .set({ ...updates, updatedAt: new Date() })
      .where(eq(staff.id, id))
      .returning();
    return updatedStaff;
  }

  async deleteStaff(id: number): Promise<void> {
    await db.delete(staff).where(eq(staff.id, id));
  }

  // Skills methods
  async getAllSkills(): Promise<Skill[]> {
    return await db.select().from(skills);
  }

  async getSkill(id: number): Promise<Skill | undefined> {
    const [skill] = await db.select().from(skills).where(eq(skills.id, id));
    return skill || undefined;
  }

  async getSkillsByCategory(category: string): Promise<Skill[]> {
    return await db.select().from(skills).where(eq(skills.category, category));
  }

  async createSkill(insertSkill: InsertSkill): Promise<Skill> {
    const [newSkill] = await db
      .insert(skills)
      .values(insertSkill)
      .returning();
    return newSkill;
  }

  async updateSkill(id: number, updates: Partial<InsertSkill>): Promise<Skill> {
    const [updatedSkill] = await db
      .update(skills)
      .set(updates)
      .where(eq(skills.id, id))
      .returning();
    return updatedSkill;
  }

  // Training Records methods
  async getAllTrainingRecords(): Promise<TrainingRecord[]> {
    return await db.select().from(trainingRecords);
  }

  async getTrainingRecord(id: number): Promise<TrainingRecord | undefined> {
    const [record] = await db.select().from(trainingRecords).where(eq(trainingRecords.id, id));
    return record || undefined;
  }

  async getTrainingRecordsByStaff(staffId: number): Promise<TrainingRecord[]> {
    return await db.select().from(trainingRecords).where(eq(trainingRecords.staffId, staffId));
  }

  async getTrainingRecordsBySkill(skillId: number): Promise<TrainingRecord[]> {
    return await db.select().from(trainingRecords).where(eq(trainingRecords.skillId, skillId));
  }

  async getTrainingMatrix(): Promise<Array<{ staffId: number; staffName: string; skillId: number; skillName: string; competencyLevel: string; status: string; expiryDate: Date | null }>> {
    const result = await db
      .select({
        staffId: trainingRecords.staffId,
        staffName: staff.name,
        skillId: trainingRecords.skillId,
        skillName: skills.name,
        competencyLevel: trainingRecords.competencyLevel,
        status: trainingRecords.status,
        expiryDate: trainingRecords.expiryDate
      })
      .from(trainingRecords)
      .innerJoin(staff, eq(trainingRecords.staffId, staff.id))
      .innerJoin(skills, eq(trainingRecords.skillId, skills.id));
    
    return result;
  }

  async createTrainingRecord(record: InsertTrainingRecord): Promise<TrainingRecord> {
    return await db.transaction(async (tx) => {
      // Check if there's already an Active record for this staff+skill combination
      const [existingActiveRecord] = await tx
        .select()
        .from(trainingRecords)
        .where(
          and(
            eq(trainingRecords.staffId, record.staffId),
            eq(trainingRecords.skillId, record.skillId),
            eq(trainingRecords.status, 'Active')
          )
        );
      
      if (existingActiveRecord) {
        // Supersede the existing Active record automatically
        await tx
          .update(trainingRecords)
          .set({ status: 'Superseded', updatedAt: new Date() })
          .where(eq(trainingRecords.id, existingActiveRecord.id));
        
        // Create new record with incremented version and previousRecordId
        const newVersion = (existingActiveRecord.version || 1) + 1;
        const [newRecord] = await tx
          .insert(trainingRecords)
          .values({
            ...record,
            previousRecordId: existingActiveRecord.id,
            version: newVersion,
            status: 'Active' // Always set to Active regardless of client input
          })
          .returning();
        
        return newRecord;
      } else {
        // Normal creation - no existing Active record
        const [newRecord] = await tx
          .insert(trainingRecords)
          .values({
            ...record,
            version: 1, // First version
            status: 'Active' // Always set to Active regardless of client input
          })
          .returning();
        return newRecord;
      }
    });
  }

  async updateTrainingRecord(id: number, updates: Partial<InsertTrainingRecord>): Promise<TrainingRecord> {
    const [updatedRecord] = await db
      .update(trainingRecords)
      .set({ ...updates, updatedAt: new Date() })
      .where(eq(trainingRecords.id, id))
      .returning();
    return updatedRecord;
  }

  // PPE Records methods
  async getAllPpeRecords(): Promise<PpeRecord[]> {
    return await db.select().from(ppeRecords);
  }

  async getPpeRecord(id: number): Promise<PpeRecord | undefined> {
    const [record] = await db.select().from(ppeRecords).where(eq(ppeRecords.id, id));
    return record || undefined;
  }

  async getPpeRecordsByStaff(staffId: number): Promise<PpeRecord[]> {
    return await db.select().from(ppeRecords).where(eq(ppeRecords.staffId, staffId));
  }

  async getExpiringPpe(withinDays: number): Promise<PpeRecord[]> {
    const futureDate = new Date();
    futureDate.setDate(futureDate.getDate() + withinDays);
    
    return await db
      .select()
      .from(ppeRecords)
      .where(
        and(
          gte(ppeRecords.expiryDate, new Date()),
          lte(ppeRecords.expiryDate, futureDate)
        )
      );
  }

  async createPpeRecord(record: InsertPpeRecord): Promise<PpeRecord> {
    const [newRecord] = await db
      .insert(ppeRecords)
      .values(record)
      .returning();
    return newRecord;
  }

  async updatePpeRecord(id: number, updates: Partial<InsertPpeRecord>): Promise<PpeRecord> {
    const [updatedRecord] = await db
      .update(ppeRecords)
      .set({ ...updates, updatedAt: new Date() })
      .where(eq(ppeRecords.id, id))
      .returning();
    return updatedRecord;
  }

  // Induction Records methods
  async getAllInductionRecords(): Promise<InductionRecord[]> {
    return await db.select().from(inductionRecords);
  }

  async getInductionRecord(id: number): Promise<InductionRecord | undefined> {
    const [record] = await db.select().from(inductionRecords).where(eq(inductionRecords.id, id));
    return record || undefined;
  }

  async getInductionRecordsByStaff(staffId: number): Promise<InductionRecord[]> {
    return await db.select().from(inductionRecords).where(eq(inductionRecords.staffId, staffId));
  }

  async createInductionRecord(record: InsertInductionRecord): Promise<InductionRecord> {
    const [newRecord] = await db
      .insert(inductionRecords)
      .values(record)
      .returning();
    return newRecord;
  }

  async updateInductionRecord(id: number, updates: Partial<InsertInductionRecord>): Promise<InductionRecord> {
    const [updatedRecord] = await db
      .update(inductionRecords)
      .set(updates)
      .where(eq(inductionRecords.id, id))
      .returning();
    return updatedRecord;
  }

  // Equipment Authorization methods
  async getAllEquipmentAuthorizations(): Promise<EquipmentAuthorization[]> {
    return await db.select().from(equipmentAuthorizations);
  }

  async getEquipmentAuthorization(id: number): Promise<EquipmentAuthorization | undefined> {
    const [auth] = await db.select().from(equipmentAuthorizations).where(eq(equipmentAuthorizations.id, id));
    return auth || undefined;
  }

  async getEquipmentAuthorizationsByStaff(staffId: number): Promise<EquipmentAuthorization[]> {
    return await db.select().from(equipmentAuthorizations).where(eq(equipmentAuthorizations.staffId, staffId));
  }

  async getEquipmentAuthorizationsByType(equipmentType: string): Promise<EquipmentAuthorization[]> {
    return await db.select().from(equipmentAuthorizations).where(eq(equipmentAuthorizations.equipmentType, equipmentType));
  }

  async createEquipmentAuthorization(auth: InsertEquipmentAuthorization): Promise<EquipmentAuthorization> {
    const [newAuth] = await db
      .insert(equipmentAuthorizations)
      .values(auth)
      .returning();
    return newAuth;
  }

  async updateEquipmentAuthorization(id: number, updates: Partial<InsertEquipmentAuthorization>): Promise<EquipmentAuthorization> {
    const [updatedAuth] = await db
      .update(equipmentAuthorizations)
      .set({ ...updates, updatedAt: new Date() })
      .where(eq(equipmentAuthorizations.id, id))
      .returning();
    return updatedAuth;
  }

  // Photo Asset methods
  async getPhotoAsset(id: number): Promise<PhotoAsset | undefined> {
    const [asset] = await db.select().from(photoAssets).where(eq(photoAssets.id, id));
    return asset || undefined;
  }

  async getPhotoAssetByStaff(staffId: number): Promise<PhotoAsset | undefined> {
    const [asset] = await db.select().from(photoAssets).where(eq(photoAssets.staffId, staffId));
    return asset || undefined;
  }

  async createPhotoAsset(asset: InsertPhotoAsset): Promise<PhotoAsset> {
    const [newAsset] = await db
      .insert(photoAssets)
      .values(asset)
      .returning();
    return newAsset;
  }

  async updatePhotoAsset(id: number, updates: Partial<InsertPhotoAsset>): Promise<PhotoAsset> {
    const [updatedAsset] = await db
      .update(photoAssets)
      .set(updates)
      .where(eq(photoAssets.id, id))
      .returning();
    return updatedAsset;
  }

  async deletePhotoAsset(id: number): Promise<void> {
    await db.delete(photoAssets).where(eq(photoAssets.id, id));
  }

  // Glove Records methods
  async getAllGloveRecords(): Promise<GloveRecord[]> {
    return await db.select().from(gloveRecords);
  }

  async getGloveRecord(id: number): Promise<GloveRecord | undefined> {
    const [record] = await db.select().from(gloveRecords).where(eq(gloveRecords.id, id));
    return record || undefined;
  }

  async getGloveRecordsByStaff(staffId: number): Promise<GloveRecord[]> {
    return await db.select().from(gloveRecords).where(eq(gloveRecords.staffId, staffId));
  }

  async createGloveRecord(record: InsertGloveRecord): Promise<GloveRecord> {
    const [newRecord] = await db
      .insert(gloveRecords)
      .values(record)
      .returning();
    return newRecord;
  }

  async updateGloveRecord(id: number, updates: Partial<InsertGloveRecord>): Promise<GloveRecord> {
    const [updatedRecord] = await db
      .update(gloveRecords)
      .set(updates)
      .where(eq(gloveRecords.id, id))
      .returning();
    return updatedRecord;
  }

  // "Able to Use" Training System methods

  // Training Classification methods
  async getAllTrainingClassifications(audience?: string): Promise<TrainingClassification[]> {
    if (audience) {
      return await db.select().from(trainingClassifications)
        .where(eq(trainingClassifications.audience, audience))
        .orderBy(trainingClassifications.displayOrder);
    }
    return await db.select().from(trainingClassifications).orderBy(trainingClassifications.displayOrder);
  }

  async getTrainingClassification(id: number): Promise<TrainingClassification | undefined> {
    const [classification] = await db.select().from(trainingClassifications).where(eq(trainingClassifications.id, id));
    return classification || undefined;
  }

  async getTrainingClassificationByKey(key: string): Promise<TrainingClassification | undefined> {
    const [classification] = await db.select().from(trainingClassifications).where(eq(trainingClassifications.key, key));
    return classification || undefined;
  }

  async createTrainingClassification(classification: InsertTrainingClassification): Promise<TrainingClassification> {
    const [newClassification] = await db
      .insert(trainingClassifications)
      .values(classification)
      .returning();
    return newClassification;
  }

  async updateTrainingClassification(id: number, updates: Partial<InsertTrainingClassification>): Promise<TrainingClassification> {
    const [updatedClassification] = await db
      .update(trainingClassifications)
      .set(updates)
      .where(eq(trainingClassifications.id, id))
      .returning();
    return updatedClassification;
  }

  // Training Module methods
  async getAllTrainingModules(): Promise<TrainingModule[]> {
    return await db.select().from(trainingModules).orderBy(trainingModules.displayOrder);
  }

  async getTrainingModule(id: number): Promise<TrainingModule | undefined> {
    const [module] = await db.select().from(trainingModules).where(eq(trainingModules.id, id));
    return module || undefined;
  }

  async getTrainingModulesByClassification(classificationId: number): Promise<TrainingModule[]> {
    return await db
      .select()
      .from(trainingModules)
      .where(eq(trainingModules.classificationId, classificationId))
      .orderBy(trainingModules.displayOrder);
  }

  async getTrainingModuleByCode(code: string): Promise<TrainingModule | undefined> {
    const [module] = await db.select().from(trainingModules).where(eq(trainingModules.code, code));
    return module || undefined;
  }

  async createTrainingModule(module: InsertTrainingModule): Promise<TrainingModule> {
    const [newModule] = await db
      .insert(trainingModules)
      .values(module)
      .returning();
    return newModule;
  }

  async updateTrainingModule(id: number, updates: Partial<InsertTrainingModule>): Promise<TrainingModule> {
    const [updatedModule] = await db
      .update(trainingModules)
      .set({ ...updates, updatedAt: new Date() })
      .where(eq(trainingModules.id, id))
      .returning();
    return updatedModule;
  }

  // Module Prerequisite methods
  async getModulePrerequisites(moduleId: number): Promise<ModulePrerequisite[]> {
    return await db
      .select()
      .from(modulePrerequisites)
      .where(eq(modulePrerequisites.moduleId, moduleId));
  }

  async createModulePrerequisite(prerequisite: InsertModulePrerequisite): Promise<ModulePrerequisite> {
    const [newPrerequisite] = await db
      .insert(modulePrerequisites)
      .values(prerequisite)
      .returning();
    return newPrerequisite;
  }

  async deleteModulePrerequisite(moduleId: number, prerequisiteModuleId: number): Promise<void> {
    await db
      .delete(modulePrerequisites)
      .where(
        and(
          eq(modulePrerequisites.moduleId, moduleId),
          eq(modulePrerequisites.prerequisiteModuleId, prerequisiteModuleId)
        )
      );
  }

  // Module Tool methods
  async getModuleTools(moduleId: number): Promise<ModuleTool[]> {
    return await db
      .select()
      .from(moduleTools)
      .where(eq(moduleTools.moduleId, moduleId))
      .orderBy(moduleTools.displayOrder);
  }

  async createModuleTool(tool: InsertModuleTool): Promise<ModuleTool> {
    const [newTool] = await db
      .insert(moduleTools)
      .values(tool)
      .returning();
    return newTool;
  }

  async updateModuleTool(id: number, updates: Partial<InsertModuleTool>): Promise<ModuleTool> {
    const [updatedTool] = await db
      .update(moduleTools)
      .set(updates)
      .where(eq(moduleTools.id, id))
      .returning();
    return updatedTool;
  }

  async deleteModuleTool(id: number): Promise<void> {
    await db.delete(moduleTools).where(eq(moduleTools.id, id));
  }

  // Module Material methods
  async getModuleMaterials(moduleId: number): Promise<ModuleMaterial[]> {
    return await db
      .select()
      .from(moduleMaterials)
      .where(eq(moduleMaterials.moduleId, moduleId))
      .orderBy(moduleMaterials.displayOrder);
  }

  async createModuleMaterial(material: InsertModuleMaterial): Promise<ModuleMaterial> {
    const [newMaterial] = await db
      .insert(moduleMaterials)
      .values(material)
      .returning();
    return newMaterial;
  }

  async updateModuleMaterial(id: number, updates: Partial<InsertModuleMaterial>): Promise<ModuleMaterial> {
    const [updatedMaterial] = await db
      .update(moduleMaterials)
      .set({ ...updates, updatedAt: new Date() })
      .where(eq(moduleMaterials.id, id))
      .returning();
    return updatedMaterial;
  }

  async deleteModuleMaterial(id: number): Promise<void> {
    await db.delete(moduleMaterials).where(eq(moduleMaterials.id, id));
  }

  // Staff Module Progress methods
  async getAllStaffModuleProgress(): Promise<StaffModuleProgress[]> {
    return await db.select().from(staffModuleProgress);
  }

  async getStaffModuleProgress(staffId: number): Promise<StaffModuleProgress[]> {
    return await db
      .select()
      .from(staffModuleProgress)
      .where(eq(staffModuleProgress.staffId, staffId));
  }

  async getModuleProgressByStaffAndModule(staffId: number, moduleId: number): Promise<StaffModuleProgress | undefined> {
    const [progress] = await db
      .select()
      .from(staffModuleProgress)
      .where(
        and(
          eq(staffModuleProgress.staffId, staffId),
          eq(staffModuleProgress.moduleId, moduleId)
        )
      );
    return progress || undefined;
  }

  async getModuleProgress(moduleId: number): Promise<StaffModuleProgress[]> {
    return await db
      .select()
      .from(staffModuleProgress)
      .where(eq(staffModuleProgress.moduleId, moduleId));
  }

  async createStaffModuleProgress(progress: InsertStaffModuleProgress): Promise<StaffModuleProgress> {
    const [newProgress] = await db
      .insert(staffModuleProgress)
      .values(progress)
      .returning();
    return newProgress;
  }

  async updateStaffModuleProgress(id: number, updates: Partial<InsertStaffModuleProgress>): Promise<StaffModuleProgress> {
    const [updatedProgress] = await db
      .update(staffModuleProgress)
      .set({ ...updates, updatedAt: new Date() })
      .where(eq(staffModuleProgress.id, id))
      .returning();
    return updatedProgress;
  }

  // Complex queries for training matrix - Fixed to read from training_records table
  async getTrainingModuleMatrix(): Promise<Array<{ 
    staffId: number; 
    staffName: string; 
    moduleId: number; 
    moduleName: string; 
    moduleCode: string;
    classificationName: string;
    status: string; 
    competencyLevel: string | null;
    completedDate: Date | null;
    expiryDate: Date | null;
    ableToUse?: boolean;
  }>> {
    const result = await db
      .select({
        staffId: staff.id,
        staffName: staff.name,
        moduleId: trainingModules.id,
        moduleName: trainingModules.name,
        moduleCode: trainingModules.code,
        classificationName: trainingClassifications.name,
        status: trainingRecords.status,
        competencyLevel: trainingRecords.competencyLevel,
        completedDate: trainingRecords.achievedDate,
        expiryDate: trainingRecords.expiryDate,
        ableToUse: trainingRecords.ableToUse,
      })
      .from(staff)
      .leftJoin(
        trainingRecords, 
        and(
          eq(staff.id, trainingRecords.staffId),
          eq(trainingRecords.status, 'Active') // Only get active training records
        )
      )
      .leftJoin(trainingModules, eq(trainingRecords.skillId, trainingModules.id))
      .leftJoin(trainingClassifications, eq(trainingModules.classificationId, trainingClassifications.id))
      .where(eq(staff.isActive, true));
      
    return result.map(row => ({
      staffId: row.staffId,
      staffName: row.staffName,
      moduleId: row.moduleId || 0,
      moduleName: row.moduleName || '',
      moduleCode: row.moduleCode || '',
      classificationName: row.classificationName || '',
      status: row.status || 'Not Started',
      competencyLevel: row.competencyLevel,
      completedDate: row.completedDate,
      expiryDate: row.expiryDate,
      ableToUse: row.ableToUse || false,
    }));
  }

  // Order Items methods
  async getActiveOrderItems(): Promise<OrderItem[]> {
    return await db
      .select()
      .from(orderItems)
      .where(eq(orderItems.status, "active"))
      .orderBy(orderItems.addedAt);
  }

  async createOrderItem(item: InsertOrderItem): Promise<OrderItem> {
    const [created] = await db.insert(orderItems).values(item).returning();
    return created;
  }

  async updateOrderItemStatus(id: number, status: string, orderedBy?: string): Promise<OrderItem> {
    const [updated] = await db
      .update(orderItems)
      .set({
        status,
        orderedAt: status === "ordered" ? new Date() : null,
        orderedBy: orderedBy || null,
      })
      .where(eq(orderItems.id, id))
      .returning();
    return updated;
  }

  async clearActiveOrderItems(): Promise<number> {
    const rows = await db
      .update(orderItems)
      .set({ status: "archived" })
      .where(eq(orderItems.status, "active"))
      .returning({ id: orderItems.id });
    return rows.length;
  }
}

export const storage = new DatabaseStorage();
