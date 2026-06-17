import { useState } from "react";
import { Dialog, DialogContent, DialogDescription, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from "@/components/ui/accordion";
import { ScrollArea } from "@/components/ui/scroll-area";
import { FileText, Users, Phone, Mail, CheckCircle, AlertCircle, Building, Shield } from "lucide-react";
import { CSCWorkflow, CSC_WORKFLOW_MODULES } from "@shared/schema";

interface CSCWorkflowPopupProps {
  workflowCode: keyof typeof CSC_WORKFLOW_MODULES;
  workflowData?: CSCWorkflow;
  triggerButton?: React.ReactNode;
}

// Default workflow data - this would normally come from your database or API
const getDefaultWorkflowData = (workflowCode: keyof typeof CSC_WORKFLOW_MODULES): CSCWorkflow => {
  const module = CSC_WORKFLOW_MODULES[workflowCode];
  
  const commonSteps = [
    {
      stepNumber: 1,
      title: "Initial Assessment",
      description: "Assess the job requirements and scope of work",
      requirements: ["Site inspection", "Client consultation", "Risk assessment"]
    },
    {
      stepNumber: 2,
      title: "Documentation Review", 
      description: "Review all relevant documentation and requirements",
      requirements: ["Check permits", "Review specifications", "Verify insurance"]
    }
  ];

  const commonContacts = [
    { role: "Operations Manager", name: "TBD", phone: "+64 XX XXX XXXX", email: "operations@company.com" },
    { role: "Health & Safety Officer", name: "TBD", phone: "+64 XX XXX XXXX", email: "safety@company.com" }
  ];

  const commonDocs = [
    { title: "Risk Assessment Form", type: "Form" as const },
    { title: "Safety Checklist", type: "Checklist" as const },
    { title: "Quality Control Checklist", type: "Checklist" as const }
  ];

  switch (workflowCode) {
    case "CSC_WF_001":
      return {
        code: module.code,
        name: "Direct Client",
        description: "Simple direct client workflow",
        content: `**Client Setup:**
• Client: Client (party being billed)
• Location: Only for multiple addresses (moved home, rental properties)
• Location Contact: When site contact differs from client

**Key Points:**
• Straightforward billing structure
• Client is both the contact and billing party
• Use locations only when necessary for multiple addresses
• Minimal complexity compared to other workflow types`
      };

    case "CSC_WF_002":
      return {
        code: module.code,
        name: "Insurance – Master Glaziers",
        description: "Standard insurance claim workflow",
        content: `**Client Setup:**
• Parent Client: Insurer (main biller)
• Child Client: Insured (Homeowner) - excess / upgrade billing
• Location Contact: When site contact differs from Child Client (tenant or onsite person)

**Key Requirements:**
• Master Glazier certification required
• Two-tier billing structure (insurer and homeowner)
• Separate tracking for excess and upgrade charges
• Insurance claim documentation and compliance`
      };

    case "CSC_WF_003":
      return {
        code: module.code,
        name: "Government & Builders",
        description: "Council housing, housing organisations, prison, government and builder contracts",
        content: `**Client Setup:**
• Client: Council/Housing Org/Government Agency (billing entity)
• Locations: Under Client (site address)
• Location Contact: Tenant, Homeowner or Facility Manager (main site contact)
• Reported By: Site Foreman/Overseer/Site Coordinator (job requester)

**Key Features:**
• Multi-site management under single client
• Clear separation of roles (billing vs. contact vs. requester)
• Compliance with government contracting requirements
• Structured hierarchy for large-scale operations`
      };

    case "CSC_WF_004":
      return {
        code: module.code,
        name: "Property Management",
        description: "Residential & commercial property management, some social housing",
        content: `**Client Setup:**
• Client: Rental Agency (only one client per agency)
• Location: Property Address (under Rental Agency)
• Location Contact: Tenant (on-site contact)
• Reported By: Property Manager (job requester)
• Property Owner: Custom Field: CareOfName (for "C/-" invoice line)

**Invoice Template & Email Setup:**

**Invoice Template:**
{{Client.Name}}
C/- {{Task.CustomFields.CareOfName}}
{{Location.Street}}, {{Location.Suburb}}, {{Location.Postcode}}

**Email Template:**
To: {{Client.Email}}
Cc: {{Task.ReportedBy.Email}}
Subject: Invoice for {{Location.Street}}

**Example Output:**
Ray White Property Management
C/- Jane and Peter Collins
7/132 Wainoni Road, Christchurch 8061

**Summary:**
• Invoice To: Rental Agency
• Invoice CC: Property Manager via Reported By
• Care Of: Property Owner for billing transparency`
      };

    default:
      return {
        code: module.code,
        name: module.name,
        description: module.description,
        steps: commonSteps,
        documentation: commonDocs,
        contacts: commonContacts
      };
  }
};

const getWorkflowIcon = (workflowCode: keyof typeof CSC_WORKFLOW_MODULES) => {
  switch (workflowCode) {
    case "CSC_WF_001": return Users;
    case "CSC_WF_002": return Shield;
    case "CSC_WF_003": return Building;
    case "CSC_WF_004": return Building;
    default: return FileText;
  }
};

const getDocumentIcon = (type: string) => {
  switch (type) {
    case "SOP": return FileText;
    case "Form": return FileText;
    case "Checklist": return CheckCircle;
    case "Reference": return AlertCircle;
    default: return FileText;
  }
};

export function CSCWorkflowPopup({ workflowCode, workflowData, triggerButton }: CSCWorkflowPopupProps) {
  const [isOpen, setIsOpen] = useState(false);
  const workflow = workflowData || getDefaultWorkflowData(workflowCode);
  const WorkflowIcon = getWorkflowIcon(workflowCode);
  
  const defaultTrigger = (
    <Button 
      variant="outline" 
      size="sm"
      data-testid={`button-view-workflow-${workflowCode.toLowerCase().replace('_', '-')}`}
    >
      <WorkflowIcon className="h-4 w-4 mr-2" />
      View {workflow.code} SOP
    </Button>
  );

  return (
    <Dialog open={isOpen} onOpenChange={setIsOpen}>
      <DialogTrigger asChild>
        {triggerButton || defaultTrigger}
      </DialogTrigger>
      <DialogContent className="sm:max-w-3xl w-[min(768px,100vw-16px)] p-6 max-h-[90vh] overflow-y-auto [&>button]:focus:outline-none [&>button]:focus:ring-0 [&>button]:focus:ring-offset-0 [&>button]:focus:border-none [&>button]:hover:bg-transparent [&>button]:border-0" data-testid={`dialog-workflow-${workflowCode.toLowerCase().replace('_', '-')}`} aria-describedby="workflow-description">
        <DialogHeader className="pb-2">
          <DialogTitle className="sr-only">CSC Workflow SOP</DialogTitle>
          <DialogDescription id="workflow-description" className="sr-only">
            Standard Operating Procedure for AroFlo workflow
          </DialogDescription>
        </DialogHeader>
        
        <ScrollArea className="max-h-[70vh]">
          <div className="space-y-4">
            {/* Clean Module Information Card */}
            <div className="bg-white border border-gray-200 rounded-lg p-3 space-y-3">
              {/* Title Section */}
              <div className="flex items-start space-x-2.5">
                <WorkflowIcon className="w-4 h-4 text-purple-600 mt-0.5 flex-shrink-0" />
                <div className="flex-1 min-w-0">
                  <h2 className="text-base font-medium text-gray-900 leading-tight">{workflow.name}</h2>
                  <div className="flex items-center space-x-1.5 mt-1.5">
                    <Badge variant="outline" className="text-xs px-1.5 py-0.5 h-5">{workflow.code}</Badge>
                    <Badge className="bg-purple-100 text-purple-800 border-purple-200 text-xs px-1.5 py-0.5 h-5">
                      <FileText className="w-2.5 h-2.5 mr-1" />
                      SOP
                    </Badge>
                  </div>
                </div>
              </div>
              
              {/* Description */}
              <div className="text-sm text-gray-700 leading-relaxed">
                {workflow.description}
              </div>
              
              {/* Software Information */}
              <div className="border-t border-gray-100 pt-2">
                <div className="flex items-center space-x-1.5 text-xs">
                  <span className="text-gray-500 font-medium">Software:</span>
                  <span className="text-gray-800 font-medium">AroFlo Field Service Management</span>
                </div>
              </div>
            </div>

            {/* SOP Content */}
            <Card>
              <CardContent className="pt-4">
                <div 
                  className="max-w-none text-sm leading-relaxed text-gray-800"
                  dangerouslySetInnerHTML={{
                    __html: workflow.content?.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>').replace(/\n/g, '<br>') || ''
                  }}
                />
              </CardContent>
            </Card>

            {/* Last Updated at bottom */}
            <div className="text-center pt-3 border-t border-gray-200">
              <div className="text-xs text-gray-500">
                Last Updated: <span className="font-medium text-gray-700">25 Sep 2025</span>
              </div>
            </div>
          </div>
        </ScrollArea>
      </DialogContent>
    </Dialog>
  );
}