CREATE TABLE "action_items" (
	"id" serial PRIMARY KEY NOT NULL,
	"list_type" text NOT NULL,
	"sharepoint_item_id" text NOT NULL,
	"action_priority" text,
	"action_status" text,
	"action_assigned_to" text,
	"action_start_date" text,
	"action_due_date" text,
	"action_notes" text,
	"meeting_notes" text,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "card_ordering" (
	"id" serial PRIMARY KEY NOT NULL,
	"card_id" text NOT NULL,
	"position" integer NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "equipment_authorizations" (
	"id" serial PRIMARY KEY NOT NULL,
	"staff_id" integer NOT NULL,
	"module_id" integer,
	"module_tool_id" integer,
	"equipment_type" text NOT NULL,
	"equipment_model" text,
	"authorized_date" timestamp NOT NULL,
	"expiry_date" timestamp,
	"authorized_by" text NOT NULL,
	"competency_level" text,
	"restrictions" text,
	"status" text DEFAULT 'Active' NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "glove_records" (
	"id" serial PRIMARY KEY NOT NULL,
	"staff_id" integer NOT NULL,
	"glove_type" text NOT NULL,
	"glove_size" text NOT NULL,
	"issue_date" timestamp NOT NULL,
	"condition" text NOT NULL,
	"return_date" timestamp,
	"reason_for_return" text,
	"status" text DEFAULT 'Issued' NOT NULL,
	"notes" text,
	"created_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "induction_records" (
	"id" serial PRIMARY KEY NOT NULL,
	"staff_id" integer NOT NULL,
	"induction_type" text NOT NULL,
	"completion_date" timestamp NOT NULL,
	"location" text,
	"conducted_by" text,
	"expiry_date" timestamp,
	"certificate_issued" boolean DEFAULT false NOT NULL,
	"notes" text,
	"status" text DEFAULT 'Completed' NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "meeting_attendance" (
	"id" serial PRIMARY KEY NOT NULL,
	"meeting_date" text NOT NULL,
	"attendee_name" text NOT NULL,
	"is_present" boolean DEFAULT false NOT NULL,
	"signature_status" text,
	"signature_data" text,
	"signed_at" text
);
--> statement-breakpoint
CREATE TABLE "meeting_locks" (
	"id" serial PRIMARY KEY NOT NULL,
	"meeting_date" text NOT NULL,
	"is_locked" boolean DEFAULT false NOT NULL,
	"is_closed" boolean DEFAULT false NOT NULL,
	"locked_at" timestamp,
	"locked_by" text,
	"closed_at" timestamp,
	"closed_by" text,
	CONSTRAINT "meeting_locks_meeting_date_unique" UNIQUE("meeting_date")
);
--> statement-breakpoint
CREATE TABLE "module_materials" (
	"id" serial PRIMARY KEY NOT NULL,
	"module_id" integer NOT NULL,
	"type" text NOT NULL,
	"title" text NOT NULL,
	"description" text,
	"sharepoint_drive_item_id" text,
	"sharepoint_web_url" text,
	"external_url" text,
	"display_order" integer DEFAULT 0 NOT NULL,
	"is_active" boolean DEFAULT true NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "module_prerequisites" (
	"id" serial PRIMARY KEY NOT NULL,
	"module_id" integer NOT NULL,
	"prerequisite_module_id" integer NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "module_tools" (
	"id" serial PRIMARY KEY NOT NULL,
	"module_id" integer NOT NULL,
	"tool_name" text NOT NULL,
	"equipment_type" text,
	"model" text,
	"description" text,
	"display_order" integer DEFAULT 0 NOT NULL,
	"is_active" boolean DEFAULT true NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "order_items" (
	"id" serial PRIMARY KEY NOT NULL,
	"item_name" text NOT NULL,
	"added_by" text NOT NULL,
	"added_at" timestamp DEFAULT now() NOT NULL,
	"status" text DEFAULT 'active' NOT NULL,
	"ordered_at" timestamp,
	"ordered_by" text
);
--> statement-breakpoint
CREATE TABLE "photo_assets" (
	"id" serial PRIMARY KEY NOT NULL,
	"staff_id" integer NOT NULL,
	"filename" text NOT NULL,
	"sharepoint_drive_item_id" text,
	"sharepoint_web_url" text,
	"thumbnail_url" text,
	"file_size" integer,
	"mime_type" text,
	"uploaded_at" timestamp DEFAULT now() NOT NULL,
	"uploaded_by" text,
	"is_active" boolean DEFAULT true NOT NULL,
	CONSTRAINT "photo_assets_sharepoint_drive_item_id_unique" UNIQUE("sharepoint_drive_item_id")
);
--> statement-breakpoint
CREATE TABLE "ppe_records" (
	"id" serial PRIMARY KEY NOT NULL,
	"staff_id" integer NOT NULL,
	"ppe_type" text NOT NULL,
	"brand" text,
	"size" text,
	"issue_date" timestamp,
	"expiry_date" timestamp,
	"condition" text,
	"location" text,
	"status" text DEFAULT 'Issued' NOT NULL,
	"notes" text,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "skills" (
	"id" serial PRIMARY KEY NOT NULL,
	"category" text NOT NULL,
	"name" text NOT NULL,
	"description" text,
	"requires_certification" boolean DEFAULT false NOT NULL,
	"validity_period_months" integer,
	"is_active" boolean DEFAULT true NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "staff" (
	"id" serial PRIMARY KEY NOT NULL,
	"azure_ad_object_id" text,
	"name" text NOT NULL,
	"email" text NOT NULL,
	"job_title" text,
	"role" text DEFAULT 'field' NOT NULL,
	"role_rank" integer DEFAULT 3 NOT NULL,
	"is_field_staff" boolean DEFAULT true NOT NULL,
	"is_administration_staff" boolean DEFAULT false NOT NULL,
	"start_date" timestamp,
	"is_active" boolean DEFAULT true NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL,
	CONSTRAINT "staff_azure_ad_object_id_unique" UNIQUE("azure_ad_object_id"),
	CONSTRAINT "staff_email_unique" UNIQUE("email")
);
--> statement-breakpoint
CREATE TABLE "staff_module_progress" (
	"id" serial PRIMARY KEY NOT NULL,
	"staff_id" integer NOT NULL,
	"module_id" integer NOT NULL,
	"status" text DEFAULT 'Not Started' NOT NULL,
	"competency_level" text DEFAULT 'Not Trained' NOT NULL,
	"trained_against_sop" boolean DEFAULT false NOT NULL,
	"sop_acknowledged_at" timestamp,
	"sop_version" text,
	"applied_date" timestamp,
	"completed_date" timestamp,
	"authorized_date" timestamp,
	"expiry_date" timestamp,
	"assessor_name" text,
	"assessor_staff_id" integer,
	"evidence_url" text,
	"evidence_filename" text,
	"notes" text,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "training_classifications" (
	"id" serial PRIMARY KEY NOT NULL,
	"key" text NOT NULL,
	"name" text NOT NULL,
	"description" text,
	"audience" text DEFAULT 'both' NOT NULL,
	"display_order" integer DEFAULT 0 NOT NULL,
	"is_active" boolean DEFAULT true NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL,
	CONSTRAINT "training_classifications_key_unique" UNIQUE("key")
);
--> statement-breakpoint
CREATE TABLE "training_modules" (
	"id" serial PRIMARY KEY NOT NULL,
	"classification_id" integer NOT NULL,
	"code" text NOT NULL,
	"name" text NOT NULL,
	"description" text,
	"audience" text DEFAULT 'both' NOT NULL,
	"validity_months" integer,
	"requires_assessment" boolean DEFAULT true NOT NULL,
	"requires_authorization" boolean DEFAULT true NOT NULL,
	"is_safety_critical" boolean DEFAULT false NOT NULL,
	"requires_certification" boolean DEFAULT false NOT NULL,
	"sop_url" text,
	"training_video_url" text,
	"display_order" integer DEFAULT 0 NOT NULL,
	"is_active" boolean DEFAULT true NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL,
	CONSTRAINT "training_modules_code_unique" UNIQUE("code")
);
--> statement-breakpoint
CREATE TABLE "training_records" (
	"id" serial PRIMARY KEY NOT NULL,
	"staff_id" integer NOT NULL,
	"skill_id" integer NOT NULL,
	"competency_level" text NOT NULL,
	"achieved_date" timestamp,
	"applied_date" timestamp,
	"expiry_date" timestamp,
	"assessor_name" text,
	"training_provider" text,
	"certificate_number" text,
	"notes" text,
	"photo_evidence_url" text,
	"photo_evidence_filename" text,
	"able_to_use" boolean DEFAULT false NOT NULL,
	"status" text DEFAULT 'Active' NOT NULL,
	"previous_record_id" integer,
	"version" integer DEFAULT 1 NOT NULL,
	"created_at" timestamp DEFAULT now() NOT NULL,
	"updated_at" timestamp DEFAULT now() NOT NULL
);
--> statement-breakpoint
CREATE TABLE "users" (
	"id" serial PRIMARY KEY NOT NULL,
	"username" text NOT NULL,
	"password" text NOT NULL,
	CONSTRAINT "users_username_unique" UNIQUE("username")
);
--> statement-breakpoint
ALTER TABLE "equipment_authorizations" ADD CONSTRAINT "equipment_authorizations_staff_id_staff_id_fk" FOREIGN KEY ("staff_id") REFERENCES "public"."staff"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "equipment_authorizations" ADD CONSTRAINT "equipment_authorizations_module_id_training_modules_id_fk" FOREIGN KEY ("module_id") REFERENCES "public"."training_modules"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "equipment_authorizations" ADD CONSTRAINT "equipment_authorizations_module_tool_id_module_tools_id_fk" FOREIGN KEY ("module_tool_id") REFERENCES "public"."module_tools"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "glove_records" ADD CONSTRAINT "glove_records_staff_id_staff_id_fk" FOREIGN KEY ("staff_id") REFERENCES "public"."staff"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "induction_records" ADD CONSTRAINT "induction_records_staff_id_staff_id_fk" FOREIGN KEY ("staff_id") REFERENCES "public"."staff"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "module_materials" ADD CONSTRAINT "module_materials_module_id_training_modules_id_fk" FOREIGN KEY ("module_id") REFERENCES "public"."training_modules"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "module_prerequisites" ADD CONSTRAINT "module_prerequisites_module_id_training_modules_id_fk" FOREIGN KEY ("module_id") REFERENCES "public"."training_modules"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "module_prerequisites" ADD CONSTRAINT "module_prerequisites_prerequisite_module_id_training_modules_id_fk" FOREIGN KEY ("prerequisite_module_id") REFERENCES "public"."training_modules"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "module_tools" ADD CONSTRAINT "module_tools_module_id_training_modules_id_fk" FOREIGN KEY ("module_id") REFERENCES "public"."training_modules"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "photo_assets" ADD CONSTRAINT "photo_assets_staff_id_staff_id_fk" FOREIGN KEY ("staff_id") REFERENCES "public"."staff"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "ppe_records" ADD CONSTRAINT "ppe_records_staff_id_staff_id_fk" FOREIGN KEY ("staff_id") REFERENCES "public"."staff"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "staff_module_progress" ADD CONSTRAINT "staff_module_progress_staff_id_staff_id_fk" FOREIGN KEY ("staff_id") REFERENCES "public"."staff"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "staff_module_progress" ADD CONSTRAINT "staff_module_progress_module_id_training_modules_id_fk" FOREIGN KEY ("module_id") REFERENCES "public"."training_modules"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "staff_module_progress" ADD CONSTRAINT "staff_module_progress_assessor_staff_id_staff_id_fk" FOREIGN KEY ("assessor_staff_id") REFERENCES "public"."staff"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "training_modules" ADD CONSTRAINT "training_modules_classification_id_training_classifications_id_fk" FOREIGN KEY ("classification_id") REFERENCES "public"."training_classifications"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "training_records" ADD CONSTRAINT "training_records_staff_id_staff_id_fk" FOREIGN KEY ("staff_id") REFERENCES "public"."staff"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "training_records" ADD CONSTRAINT "training_records_skill_id_training_modules_id_fk" FOREIGN KEY ("skill_id") REFERENCES "public"."training_modules"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "training_records" ADD CONSTRAINT "training_records_previous_record_id_training_records_id_fk" FOREIGN KEY ("previous_record_id") REFERENCES "public"."training_records"("id") ON DELETE set null ON UPDATE no action;--> statement-breakpoint
CREATE UNIQUE INDEX "unique_action_item" ON "action_items" USING btree ("list_type","sharepoint_item_id");--> statement-breakpoint
CREATE UNIQUE INDEX "unique_active_training_record" ON "training_records" USING btree ("staff_id","skill_id") WHERE status = 'Active';