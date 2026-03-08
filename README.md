# Cranfield Glass Meeting Management System

A comprehensive staff engagement system for Cranfield Glass Christchurch to capture, manage, and track business ideas, safety improvements, and near miss incidents.

## Current Status & Action Items

For a detailed analysis of the current integration status and required actions, please see:
- **[SHAREPOINT_INTEGRATION_STATUS.md](./SHAREPOINT_INTEGRATION_STATUS.md)** - Complete status report and action plan

### Quick Summary

**✅ Working**: 
- 3 of 4 SharePoint lists fully integrated (Business Ideas, Safety Ideas, Near Miss)
- 117 items loading and displaying correctly
- All core features operational

**⚠️ Needs Attention**:
- Actions list returning 0 items (permissions issue)
- Missing critical fields in Actions list (Completion Date, Outcome Result)
- Some Near Miss fields need sample data to capture internal names

## Features

- **SharePoint Integration**: Direct connection to Microsoft SharePoint Lists for authentic data access
- **Microsoft 365 Authentication**: Secure MSAL 2.0 authentication with popup-based login
- **AI-Powered Enhancement**: OpenAI GPT-4o integration for intelligent title generation
- **Professional Document Export**: Advanced Word document generation with company branding
- **Responsive Design**: Mobile-first approach with progressive enhancement
- **Real-time Data**: Live SharePoint REST API integration with proper timezone handling

## Technology Stack

- **Frontend**: React 18 with TypeScript, Vite, Tailwind CSS
- **Backend**: Node.js with Express.js
- **Authentication**: Azure MSAL 2.0 for SharePoint access
- **Database**: PostgreSQL via Neon Database
- **AI Integration**: OpenAI GPT-4o for content enhancement
- **Document Generation**: Advanced Word templating engine

## Environment Setup

Required environment variables:

```
# Azure/SharePoint Configuration
AZURE_TENANT_ID=your_tenant_id
AZURE_CLIENT_ID=your_client_id
AZURE_CLIENT_SECRET=your_client_secret
AZURE_REDIRECT_URI=your_redirect_uri

# SharePoint List IDs
SAFETY_IDEAS_SITE_ID=your_site_id
SAFETY_IDEAS_LIST_ID=your_list_id
BUSINESS_IDEAS_LIST_ID=your_list_id
ACCIDENT_REPORTS_SITE_ID=your_site_id
ACCIDENT_REPORTS_LIST_ID=your_list_id

# OpenAI Integration (Optional)
OPENAI_API_KEY=your_openai_key

# Database
DATABASE_URL=your_postgresql_connection_string
```

## Development

```bash
# Install dependencies
npm install

# Start development server
npm run dev

# Build for production
npm run build

# Start production server
npm start
```

## Production Deployment

1. Set all required environment variables
2. Run `npm run build` to create optimized production build
3. Deploy with `npm start` for production server

## Architecture

The system follows modern web application patterns with:
- Clean separation between frontend and backend
- Type-safe SharePoint API integration
- Secure authentication flow
- Responsive mobile-first design
- Professional document generation
- Real-time data synchronization

For detailed technical documentation, see the project's replit.md file.