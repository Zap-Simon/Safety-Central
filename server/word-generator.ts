import { AdvancedWordTemplateEngine } from './template-engine';

interface MeetingItem {
  id: string;
  title: string;
  description: string;
  type: string;
  status: string;
  meetingDate: string;
  meetingNotes: string;
  submittedBy: string;
  submittedDate: string;
  ideaType?: string;
  assignedTo?: string;
}

interface AttendanceData {
  [meetingDate: string]: string[];
}

export async function generateMeetingWordDoc(data: any): Promise<Buffer> {
  // Use the advanced templating engine for professional document generation
  return await AdvancedWordTemplateEngine.generateMeetingDocument(data);
}

// Status color helper for legacy compatibility
function getStatusColor(status: string): string {
  switch (status?.toLowerCase()) {
    case 'submitted': return '3B82F6';
    case 'in discussion': return 'F59E0B';
    case 'actions': return '10B981';
    case 'closed': return '6B7280';
    default: return '4B5563';
  }
}