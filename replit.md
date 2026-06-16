# CRANFIELD Health & Safety Idea Management System

## Overview

This is a comprehensive staff engagement system for Cranfield Glass Christchurch, designed to capture, manage, and track business ideas, safety improvements, and near-miss incidents. It offers a modern web interface, enterprise-grade document export, mobile optimization, and AI integration to streamline a Microsoft 365-based workflow involving Microsoft Forms, SharePoint Lists, Power Automate, and Microsoft Teams. The system aims to enhance staff engagement, improve safety, and streamline administrative processes through intelligent automation and robust data management.

## User Preferences

Preferred communication style: Simple, everyday language.
Authentication Needs: User wants a complete boilerplate folder containing all Microsoft Azure authentication code and secrets for reuse in other projects.
Code Reusability: Prefers well-documented, portable authentication components that can be easily copied to new projects.

## System Architecture

### Frontend Architecture
- **Framework**: React 18 with TypeScript.
- **Routing**: Wouter for lightweight navigation.
- **Styling**: Tailwind CSS with custom Microsoft design system variables, shadcn/ui for UI components.
- **State Management**: TanStack Query for server state.
- **Build Tool**: Vite for fast development and optimized builds.
- **Responsive Design**: Mobile-first approach.

### Backend Architecture
- **Runtime**: Node.js with Express.js.
- **Development**: tsx for TypeScript execution.
- **Production Build**: esbuild for optimized server-side bundling.
- **API Design**: RESTful endpoints with centralized error handling.
- **Authentication**: Express sessions with PostgreSQL backing store, Azure AD integration via Microsoft Graph API and MSAL-node. The Microsoft Teams tabs use the OAuth 2.0 On-Behalf-Of (OBO) flow: the browser only holds a Teams SSO token and the server exchanges it for downstream Graph/SharePoint tokens (`server/teams-obo-auth.ts`). The main website still sends ready-made tokens, which pass through unchanged. See `teams-app/README.md` for the full Azure AD setup.
- **Permissions**: Sites.ReadWrite.All, User.Read.All for SharePoint access.

### Data Storage
- **Database**: PostgreSQL via Neon Database (serverless).
- **ORM**: Drizzle ORM with type-safe schema definitions.
- **Session Management**: connect-pg-simple for PostgreSQL session storage.
- **Schema Management**: Drizzle Kit for database migrations.

### UI/UX Decisions
- Clean, professional design with consistent white backgrounds and subtle borders.
- Eliminated gradient backgrounds as per user preference, favoring clean white, gray text, and subtle borders.
- Mobile-first optimization with consistent badge heights, square icons, and streamlined navigation.
- Inline predictive text completion with Tab-to-complete functionality, similar to IDEs.
- Professional export system generates HTML and CSV with Cranfield Glass branding, A4 optimization, structured tables, and typography.
- Unified modal notification system replaced all browser alerts for a professional user experience.

### Technical Implementations & Feature Specifications
- **Microsoft 365 Integration**: Utilizes SharePoint Lists (Business Ideas, Safety Ideas, Near Miss), Power Automate for workflows and notifications, and Microsoft Forms embedded in Teams Viva cards.
- **SharePoint API Integration**: Modular `LIST_CONFIGS` for easy expansion, type-safe interfaces, and consistent processing across lists. Handles complex field mappings and multi-site URL support.
- **SharePoint Document Libraries**: Direct integration with SharePoint document libraries using Microsoft Graph API. Includes adaptive folder discovery, retry logic for service interruptions, and unified document display interface. Supports multiple SharePoint sites with different folder structures.
- **Authentication**: Secure session management with PostgreSQL persistence, Azure AD integration, and robust permissions handling. Authentication boilerplate for reusability in other projects.
- **AI-Powered Features**: OpenAI GPT-4o (and gpt-4o-mini for speed) for intelligent title generation, smart content processing, and meeting notes generation with a natural, conversational tone. Includes live AI auto-suggestions and graceful fallbacks.
- **Professional Export System**: Generates professional HTML meeting minutes (browser-ready, A4 optimized) and structured CSV data exports. Supports dynamic attendance tracking and multi-meeting export.
- **Actions Workflow Integration**: Architecture for integrating an 'Actions' list where items marked "Actioned" are copied for tracking, and marked "Finished" close original items, ensuring data synchronization.
- **Meeting History & Live Stats**: Proactive meeting dashboard with live statistics, showing ongoing items from previous meetings. Includes intelligent item tracking and clean filter interfaces.
- **Attendance Lock System**: Comprehensive lock functionality for meeting attendance with visual indicators and persistent state to prevent accidental changes.
- **Equipment Test & Tag Integration**: Comprehensive integration with EquipmentMaintenance SharePoint site for electrical testing and tagging compliance records. Features adaptive folder discovery and resilient error handling.

## External Dependencies

### Microsoft Services
- **Microsoft Graph API**: Core integration for SharePoint data access, Azure Active Directory.
- **Azure Active Directory**: Authentication and user management.
- **SharePoint Online**: Primary data storage and workflow management.
- **Microsoft Teams**: Communication and notification platform.
- **Power Automate**: Workflow automation and business logic.

### Third-Party Services
- **OpenAI API**: GPT-4o model for content enhancement and title generation.
- **Neon Database**: Serverless PostgreSQL for application data.

### Development Tools
- **Vite**: Fast build tool.
- **TypeScript**: Type safety across frontend and backend.
- **Tailwind CSS**: Utility-first styling framework.
- **Replit Platform**: Development and hosting environment.
- **Mermaid.js**: (Potentially for workflow visualization, mentioned in original but not in implemented features)