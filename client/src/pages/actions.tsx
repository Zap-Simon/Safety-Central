import { useState, useRef, useEffect } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { authService } from "@/auth/authService";
import MeetingHeader from "@/components/meeting-header";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Calendar, FileText, AlertTriangle, Lightbulb, Shield, CheckCircle, Download, Search, Clock, User, Target } from "lucide-react";
import { format } from "date-fns";

interface ActionableItem {
  id: string;
  title: string;
  description: string;
  type: 'Safety Ideas' | 'Business Ideas' | 'Near Miss';
  status: string;
  meetingDate: string;
  meetingNotes: string;
  submittedBy: string;
  submittedDate: string;
  assignedTo?: string;
  secondaryDescription?: string;
  ideaType?: string;
  actionAssignedTo?: string;
  actionStatus?: string;
  actionPriority?: string;
  actionStartDate?: string;
  actionDueDate?: string;
  actionNotes?: string;
}

interface ActivityEntry {
  id: number;
  listType: string;
  sharePointItemId: string;
  entryType: string;
  content: string;
  author: string | null;
  createdAt: string;
}

export default function Actions() {
  const queryClient = useQueryClient();
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [filterStatus, setFilterStatus] = useState<string>('open');
  const [filterPriority, setFilterPriority] = useState<string>('all');
  const [filterType, setFilterType] = useState<string>('all');
  const [filterAssignedTo, setFilterAssignedTo] = useState<string>('all');
  const [isUpdatingAction, setIsUpdatingAction] = useState<string>('');
  const [recentlySavedActions, setRecentlySavedActions] = useState<Set<string>>(new Set());
  const [isEnhancingActionNotes, setIsEnhancingActionNotes] = useState<string>('');
  const [localActionEdits, setLocalActionEdits] = useState<Record<string, { actionNotes?: string; actionAssignedTo?: string }>>({});
  const debounceTimers = useRef<Record<string, NodeJS.Timeout>>({});
  const [showExportModal, setShowExportModal] = useState(false);
  const [selectedItem, setSelectedItem] = useState<ActionableItem | null>(null);
  const [activityLogs, setActivityLogs] = useState<Record<string, ActivityEntry[]>>({});
  const [loadingActivity, setLoadingActivity] = useState<string>('');

  // Use the SharePoint-scoped token (same as meeting-history.tsx). The
  // /api/meeting-history endpoint talks to SharePoint, so it needs a SharePoint
  // token — NOT the login/Graph token from getAccessToken(). Using the wrong
  // token type is what caused the 401s. getSharePointToken acquires silently
  // (it only falls back to a popup on a genuine InteractionRequiredAuthError),
  // so the normal data-load path won't hit the Cross-Origin-Opener-Policy block.
  const authenticatedFetch = async (url: string, options: RequestInit = {}) => {
    try {
      const token = await authService.getSharePointToken();
      return fetch(url, {
        ...options,
        headers: {
          ...options.headers,
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });
    } catch (error) {
      console.error('Error getting SharePoint token:', error);
      throw error;
    }
  };

  const getListType = (item: ActionableItem) =>
    item.type === 'Near Miss' ? 'NearMiss' : item.type.replace(' ', '');

  const fetchActivityLog = async (item: ActionableItem) => {
    setLoadingActivity(item.id);
    try {
      const listType = getListType(item);
      const res = await authenticatedFetch(`/api/action-activity/${listType}/${encodeURIComponent(item.id)}`);
      if (res.ok) {
        const data = await res.json();
        setActivityLogs(prev => ({ ...prev, [item.id]: data.data || [] }));
      }
    } catch { /* silent */ } finally {
      setLoadingActivity('');
    }
  };

  const postActivityEntry = async (item: ActionableItem, entryType: string, content: string) => {
    try {
      await authenticatedFetch('/api/action-activity', {
        method: 'POST',
        body: JSON.stringify({ listType: getListType(item), sharePointItemId: item.id, entryType, content, author: null }),
      }).then(async res => {
        if (res.ok) {
          const data = await res.json();
          setActivityLogs(prev => ({ ...prev, [item.id]: [...(prev[item.id] || []), data.data] }));
        }
      });
    } catch { /* silent */ }
  };

  useEffect(() => {
    if (selectedItem) fetchActivityLog(selectedItem);
  }, [selectedItem?.id]);

  const { data: apiResponse, isLoading, isError, refetch } = useQuery({
    queryKey: ['/api/meeting-history'],
    queryFn: () => authenticatedFetch('/api/meeting-history').then(async res => {
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return res.json();
    }),
    staleTime: 5 * 60 * 1000,
    // Retry transient faults (e.g. token not ready for a beat right after a deploy
    // reload) so the request self-heals. Do NOT retry when the user genuinely needs
    // to sign in again — that would just re-trigger sign-in popups over and over.
    retry: (failureCount, error) => {
      const msg = error instanceof Error ? error.message : String(error);
      if (/interaction|popup|cancel|No authenticated accounts/i.test(msg)) return false;
      return failureCount < 3;
    },
    retryDelay: (attempt) => Math.min(1000 * 2 ** attempt, 8000),
  });

  const meetingItems: ActionableItem[] = (apiResponse as any)?.data || [];

  const uniqueNames: string[] = Array.from(new Set(
    meetingItems.flatMap(i => [i.submittedBy, i.actionAssignedTo].filter(Boolean) as string[])
  )).sort();

  const actionItems = meetingItems.filter(item => {
    const hasActionData = !!(
      item.actionPriority || 
      item.actionStatus || 
      item.actionAssignedTo || 
      item.actionStartDate || 
      item.actionDueDate || 
      item.actionNotes ||
      item.status === 'Actioned'
    );
    return hasActionData;
  });

  // A finished/archived action is one that's either fully Completed or sitting in
  // "Ready to Close" awaiting formal closure. These are kept out of the live
  // workload numbers so the dashboard reflects only actions still to manage.
  const isArchivedStatus = (status?: string) =>
    status === 'Completed' || status === 'Ready to Close';

  const filteredItems = actionItems.filter(item => {
    if (searchQuery) {
      const query = searchQuery.toLowerCase();
      const matchesSearch = 
        item.title?.toLowerCase().includes(query) ||
        item.description?.toLowerCase().includes(query) ||
        item.actionNotes?.toLowerCase().includes(query) ||
        item.actionAssignedTo?.toLowerCase().includes(query) ||
        item.submittedBy?.toLowerCase().includes(query);
      if (!matchesSearch) return false;
    }

    if (filterStatus !== 'all') {
      const itemStatus = item.actionStatus || 'Not Started';
      const isArchived = isArchivedStatus(itemStatus);
      // "open" = live actions you still need to manage (hides anything finished).
      if (filterStatus === 'open' && isArchived) return false;
      // "archived" = finished work (Completed or Ready to Close) kept out of the way.
      if (filterStatus === 'archived' && !isArchived) return false;
      if (filterStatus === 'Completed' && itemStatus !== 'Completed') return false;
      if (filterStatus !== 'open' && filterStatus !== 'archived' && filterStatus !== 'Completed' && itemStatus !== filterStatus) return false;
    }

    if (filterPriority !== 'all' && item.actionPriority !== filterPriority) return false;
    if (filterType !== 'all' && item.type !== filterType) return false;
    if (filterAssignedTo !== 'all' && (item.actionAssignedTo || '') !== filterAssignedTo) return false;

    return true;
  });

  const sortedItems = [...filteredItems].sort((a, b) => {
    const priorityOrder: Record<string, number> = { 'High': 1, 'Medium': 2, 'Low': 3 };
    const aPriority = priorityOrder[a.actionPriority || ''] || 4;
    const bPriority = priorityOrder[b.actionPriority || ''] || 4;
    if (aPriority !== bPriority) return aPriority - bPriority;
    
    const aDate = a.actionDueDate ? new Date(a.actionDueDate).getTime() : Infinity;
    const bDate = b.actionDueDate ? new Date(b.actionDueDate).getTime() : Infinity;
    return aDate - bDate;
  });

  const getActionFieldValue = (item: ActionableItem, field: 'actionNotes' | 'actionAssignedTo') => {
    if (localActionEdits[item.id]?.[field] !== undefined) {
      return localActionEdits[item.id][field] || '';
    }
    return item[field] ?? '';
  };

  const handleActionTextChange = (item: ActionableItem, field: 'actionNotes' | 'actionAssignedTo', value: string) => {
    setLocalActionEdits(prev => ({
      ...prev,
      [item.id]: { ...prev[item.id], [field]: value }
    }));

    // actionNotes uses a manual Save button — no auto-save
    if (field === 'actionNotes') return;

    if (debounceTimers.current[`${item.id}-${field}`]) {
      clearTimeout(debounceTimers.current[`${item.id}-${field}`]);
    }

    debounceTimers.current[`${item.id}-${field}`] = setTimeout(() => {
      updateActionFields(item, { [field]: value }, field);
    }, 800);
  };

  const updateActionFields = async (
    item: ActionableItem,
    updates: {
      actionPriority?: string;
      actionStatus?: string;
      actionAssignedTo?: string;
      actionStartDate?: string;
      actionDueDate?: string;
      actionNotes?: string;
    },
    clearLocalEditField?: 'actionNotes' | 'actionAssignedTo'
  ) => {
    setIsUpdatingAction(item.id);
    try {
      const payload = {
        listType: item.type === 'Near Miss' ? 'NearMiss' : item.type.replace(' ', ''),
        sharePointItemId: item.id,
        actionPriority: updates.actionPriority ?? item.actionPriority ?? '',
        actionStatus: updates.actionStatus ?? item.actionStatus ?? '',
        actionAssignedTo: localActionEdits[item.id]?.actionAssignedTo ?? updates.actionAssignedTo ?? item.actionAssignedTo ?? '',
        actionStartDate: updates.actionStartDate ?? item.actionStartDate ?? '',
        actionDueDate: updates.actionDueDate ?? item.actionDueDate ?? '',
        actionNotes: localActionEdits[item.id]?.actionNotes ?? updates.actionNotes ?? item.actionNotes ?? ''
      };

      const response = await authenticatedFetch('/api/action-items', {
        method: 'POST',
        body: JSON.stringify(payload)
      });

      if (response.ok) {
        // Update the cache directly so changes show immediately without waiting for a re-fetch
        queryClient.setQueryData(['/api/meeting-history'], (old: any) => {
          if (!old?.data) return old;
          return {
            ...old,
            data: old.data.map((i: any) => i.id === item.id ? { ...i, ...payload } : i)
          };
        });

        if (clearLocalEditField) {
          setLocalActionEdits(prev => {
            const updated = { ...prev };
            if (updated[item.id]) {
              delete updated[item.id][clearLocalEditField];
              if (Object.keys(updated[item.id]).length === 0) {
                delete updated[item.id];
              }
            }
            return updated;
          });
        }
        
        setRecentlySavedActions(prev => new Set(Array.from(prev).concat([item.id])));
        setTimeout(() => {
          setRecentlySavedActions(prev => {
            const newSet = new Set(prev);
            newSet.delete(item.id);
            return newSet;
          });
        }, 2000);

        // Background re-fetch to keep in sync (non-blocking)
        queryClient.invalidateQueries({ queryKey: ['/api/meeting-history'] });
      }
    } catch (error) {
      console.error('Failed to update action fields:', error);
    } finally {
      setIsUpdatingAction('');
    }
  };


  const formatDate = (dateString: string | undefined) => {
    if (!dateString) return 'Not set';
    try {
      return format(new Date(dateString), 'dd MMM yyyy');
    } catch {
      return dateString;
    }
  };

  const getPriorityColor = (priority: string | undefined) => {
    switch (priority) {
      case 'High': return 'bg-red-100 text-red-800 border-red-300';
      case 'Medium': return 'bg-amber-100 text-amber-800 border-amber-300';
      case 'Low': return 'bg-green-100 text-green-800 border-green-300';
      default: return 'bg-gray-100 text-gray-600 border-gray-300';
    }
  };

  const getStatusColor = (status: string | undefined) => {
    switch (status) {
      case 'Completed': return 'bg-green-100 text-green-800 border-green-300';
      case 'In Progress': return 'bg-blue-100 text-blue-800 border-blue-300';
      case 'On Hold': return 'bg-orange-100 text-orange-800 border-orange-300';
      case 'Ready to Close': return 'bg-emerald-100 text-emerald-800 border-emerald-300';
      default: return 'bg-gray-100 text-gray-600 border-gray-300';
    }
  };

  const getTypeIcon = (type: string) => {
    switch (type) {
      case 'Safety Ideas': return <Shield className="h-4 w-4 text-green-600" />;
      case 'Business Ideas': return <Lightbulb className="h-4 w-4 text-blue-600" />;
      case 'Near Miss': return <AlertTriangle className="h-4 w-4 text-orange-600" />;
      default: return <FileText className="h-4 w-4 text-gray-600" />;
    }
  };

  const getTypeBadge = (type: string) => {
    const colorClass =
      type === 'Safety Ideas'   ? 'bg-red-100    text-red-800    border-red-200'    :
      type === 'Business Ideas' ? 'bg-blue-100   text-blue-800   border-blue-200'   :
      type === 'Near Miss'      ? 'bg-orange-100 text-orange-800 border-orange-200' :
                                  'bg-gray-100   text-gray-800   border-gray-200';
    return (
      <Badge variant="outline" className={`text-[10px] font-medium px-1.5 py-0 h-5 leading-none rounded-full ${colorClass}`}>
        {type || 'Unknown'}
      </Badge>
    );
  };

  const getDueDateStatus = (dueDate: string | undefined) => {
    if (!dueDate) return null;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const due = new Date(dueDate);
    due.setHours(0, 0, 0, 0);
    const diffDays = Math.ceil((due.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
    
    if (diffDays < 0) return { label: 'Overdue', color: 'text-red-600 bg-red-50' };
    if (diffDays === 0) return { label: 'Due Today', color: 'text-orange-600 bg-orange-50' };
    if (diffDays <= 3) return { label: `${diffDays} days`, color: 'text-amber-600 bg-amber-50' };
    if (diffDays <= 7) return { label: `${diffDays} days`, color: 'text-blue-600 bg-blue-50' };
    return { label: `${diffDays} days`, color: 'text-gray-600 bg-gray-50' };
  };

  const stats = {
    total: actionItems.length,
    open: actionItems.filter(i => !isArchivedStatus(i.actionStatus)).length,
    completed: actionItems.filter(i => i.actionStatus === 'Completed').length,
    overdue: actionItems.filter(i => {
      if (!i.actionDueDate || isArchivedStatus(i.actionStatus)) return false;
      return new Date(i.actionDueDate) < new Date();
    }).length,
    highPriority: actionItems.filter(i => i.actionPriority === 'High' && !isArchivedStatus(i.actionStatus)).length
  };

  const exportToCSV = async () => {
    const headers = ['Title', 'Type', 'Priority', 'Status', 'Assigned To', 'Due Date', 'Meeting Date', 'Action Notes', 'Submitted By'];
    const rows = sortedItems.map(item => [
      item.title || '',
      item.type,
      item.actionPriority || 'Not Set',
      item.actionStatus || 'Not Started',
      item.actionAssignedTo || 'Unassigned',
      item.actionDueDate ? formatDate(item.actionDueDate) : 'Not Set',
      formatDate(item.meetingDate),
      (item.actionNotes || '').replace(/"/g, '""'),
      item.submittedBy || ''
    ]);

    const csvContent = [
      headers.join(','),
      ...rows.map(row => row.map(cell => `"${cell}"`).join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Actions_Export_${format(new Date(), 'yyyy-MM-dd')}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportToHTML = async () => {
    const htmlContent = `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Actions Report - Cranfield Glass Christchurch</title>
  <style>
    body { font-family: Arial, sans-serif; max-width: 1200px; margin: 0 auto; padding: 20px; background: #f9fafb; }
    .header { background: linear-gradient(135deg, #1e3a5f 0%, #2563eb 100%); color: white; padding: 30px; border-radius: 12px; margin-bottom: 30px; }
    .header h1 { margin: 0 0 10px 0; font-size: 28px; }
    .header p { margin: 0; opacity: 0.9; }
    .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin-bottom: 30px; }
    .stat-card { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; }
    .stat-value { font-size: 32px; font-weight: bold; color: #1e3a5f; }
    .stat-label { color: #6b7280; font-size: 14px; margin-top: 5px; }
    .action-card { background: white; border-radius: 10px; padding: 20px; margin-bottom: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); border-left: 4px solid #f59e0b; }
    .action-card.high { border-left-color: #ef4444; }
    .action-card.medium { border-left-color: #f59e0b; }
    .action-card.low { border-left-color: #22c55e; }
    .action-card.completed { border-left-color: #10b981; opacity: 0.8; }
    .action-title { font-size: 18px; font-weight: bold; color: #111827; margin-bottom: 10px; }
    .action-meta { display: flex; flex-wrap: wrap; gap: 15px; margin-bottom: 15px; font-size: 14px; color: #6b7280; }
    .badge { display: inline-block; padding: 4px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .badge-high { background: #fee2e2; color: #dc2626; }
    .badge-medium { background: #fef3c7; color: #d97706; }
    .badge-low { background: #dcfce7; color: #16a34a; }
    .badge-completed { background: #d1fae5; color: #059669; }
    .badge-inprogress { background: #dbeafe; color: #2563eb; }
    .notes { background: #f9fafb; padding: 15px; border-radius: 8px; margin-top: 10px; font-size: 14px; color: #374151; }
    .footer { text-align: center; padding: 30px; color: #6b7280; font-size: 12px; border-top: 1px solid #e5e7eb; margin-top: 30px; }
    @media print { body { background: white; } .action-card { break-inside: avoid; } }
  </style>
</head>
<body>
  <div class="header">
    <h1>Actions Report</h1>
    <p>Cranfield Glass Christchurch - Health & Safety Management</p>
    <p>Generated: ${format(new Date(), 'dd MMMM yyyy, HH:mm')}</p>
  </div>

  <div class="stats">
    <div class="stat-card">
      <div class="stat-value">${stats.total}</div>
      <div class="stat-label">Total Actions</div>
    </div>
    <div class="stat-card">
      <div class="stat-value" style="color: #2563eb">${stats.open}</div>
      <div class="stat-label">Open Actions</div>
    </div>
    <div class="stat-card">
      <div class="stat-value" style="color: #10b981">${stats.completed}</div>
      <div class="stat-label">Completed</div>
    </div>
    <div class="stat-card">
      <div class="stat-value" style="color: #ef4444">${stats.overdue}</div>
      <div class="stat-label">Overdue</div>
    </div>
    <div class="stat-card">
      <div class="stat-value" style="color: #f59e0b">${stats.highPriority}</div>
      <div class="stat-label">High Priority</div>
    </div>
  </div>

  ${sortedItems.map(item => `
    <div class="action-card ${(item.actionPriority || '').toLowerCase()} ${item.actionStatus === 'Completed' ? 'completed' : ''}">
      <div class="action-title">${item.title || 'Untitled Action'}</div>
      <div class="action-meta">
        <span><strong>Type:</strong> ${item.type}</span>
        <span><strong>Priority:</strong> <span class="badge badge-${(item.actionPriority || 'medium').toLowerCase()}">${item.actionPriority || 'Not Set'}</span></span>
        <span><strong>Status:</strong> <span class="badge ${item.actionStatus === 'Completed' ? 'badge-completed' : item.actionStatus === 'In Progress' ? 'badge-inprogress' : ''}">${item.actionStatus || 'Not Started'}</span></span>
        <span><strong>Assigned:</strong> ${item.actionAssignedTo || 'Unassigned'}</span>
        <span><strong>Due:</strong> ${item.actionDueDate ? formatDate(item.actionDueDate) : 'Not Set'}</span>
      </div>
      ${item.actionNotes ? `<div class="notes"><strong>Action Notes:</strong> ${item.actionNotes}</div>` : ''}
      ${item.meetingNotes ? `<div class="notes" style="margin-top: 10px;"><strong>Meeting Discussion:</strong> ${item.meetingNotes}</div>` : ''}
    </div>
  `).join('')}

  <div class="footer">
    <p>Cranfield Glass Christchurch - Health & Safety Compliance Records</p>
    <p>Document ID: CG-ACT-${format(new Date(), 'yyyyMMdd-HHmm')}</p>
  </div>
</body>
</html>`;

    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Actions_Report_${format(new Date(), 'yyyy-MM-dd')}.html`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  if (isLoading && !apiResponse) {
    return (
      <div className="min-h-screen bg-gray-50">
        <MeetingHeader />
        <div className="container mx-auto px-4 py-8">
          <div className="flex items-center justify-center h-64">
            <div className="text-center">
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-amber-600 mx-auto mb-4"></div>
              <p className="text-gray-600">Loading actions...</p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (isError && !apiResponse) {
    return (
      <div className="min-h-screen bg-gray-50">
        <MeetingHeader />
        <div className="container mx-auto px-4 py-8">
          <div className="flex items-center justify-center h-64">
            <div className="text-center">
              <AlertTriangle className="h-12 w-12 text-amber-500 mx-auto mb-4" />
              <p className="text-gray-700 font-medium mb-2">Couldn't load actions</p>
              <p className="text-gray-500 text-sm mb-4">Your session may have refreshed. Try again below.</p>
              <Button onClick={() => refetch()} className="bg-amber-600 hover:bg-amber-700 text-white">
                Retry
              </Button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <MeetingHeader />
      
      <div className="pt-16 sm:pt-20 bg-gray-50 min-h-screen">
        <div className="container mx-auto px-4 py-6">

        <div className="mb-6">
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 mb-6">
            <div>
              <h1 className="text-2xl sm:text-3xl font-bold text-gray-900 flex items-center gap-3">
                <div className="w-10 h-10 bg-gradient-to-br from-amber-400 to-orange-500 rounded-lg flex items-center justify-center shadow-lg">
                  <Target className="h-6 w-6 text-white" />
                </div>
                Action Management
              </h1>
              <p className="text-gray-600 mt-1">Track, assign, and manage all health & safety actions</p>
            </div>
            <div className="flex gap-2">
              <Button
                variant="outline"
                onClick={() => setShowExportModal(true)}
                className="flex items-center gap-2"
                data-testid="button-export-actions"
              >
                <Download className="h-4 w-4" />
                <span>Export</span>
              </Button>
            </div>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 mb-6">
            <div className="bg-white rounded-lg p-4 shadow-sm border border-gray-100">
              <div className="text-2xl font-bold text-gray-900">{stats.total}</div>
              <div className="text-xs text-gray-500">Total Actions</div>
            </div>
            <div className="bg-white rounded-lg p-4 shadow-sm border border-blue-100">
              <div className="text-2xl font-bold text-blue-600">{stats.open}</div>
              <div className="text-xs text-gray-500">Open</div>
            </div>
            <div className="bg-white rounded-lg p-4 shadow-sm border border-green-100">
              <div className="text-2xl font-bold text-green-600">{stats.completed}</div>
              <div className="text-xs text-gray-500">Completed</div>
            </div>
            <div className="bg-white rounded-lg p-4 shadow-sm border border-red-100">
              <div className="text-2xl font-bold text-red-600">{stats.overdue}</div>
              <div className="text-xs text-gray-500">Overdue</div>
            </div>
            <div className="bg-white rounded-lg p-4 shadow-sm border border-amber-100 col-span-2 sm:col-span-1">
              <div className="text-2xl font-bold text-amber-600">{stats.highPriority}</div>
              <div className="text-xs text-gray-500">High Priority</div>
            </div>
          </div>

          <div className="bg-white rounded-lg shadow-sm border border-gray-100 p-4 mb-6">
            <div className="flex flex-col sm:flex-row gap-4">
              <div className="flex-1">
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-gray-400" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Search actions..."
                    className="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-amber-500"
                    data-testid="input-search-actions"
                  />
                </div>
              </div>
              <div className="grid grid-cols-3 gap-2 sm:flex sm:flex-wrap sm:gap-2">
                <select
                  value={filterStatus}
                  onChange={(e) => setFilterStatus(e.target.value)}
                  className="w-full sm:w-auto px-2 sm:px-3 py-2 border border-gray-200 rounded-lg text-xs sm:text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                  data-testid="select-filter-status"
                >
                  <option value="open">Open (to manage)</option>
                  <option value="all">All Status</option>
                  <option value="Not Started">Not Started</option>
                  <option value="In Progress">In Progress</option>
                  <option value="On Hold">On Hold</option>
                  <option value="Completed">Completed</option>
                  <option value="Ready to Close">Ready to Close</option>
                  <option value="archived">Archived (finished)</option>
                </select>
                <select
                  value={filterPriority}
                  onChange={(e) => setFilterPriority(e.target.value)}
                  className="w-full sm:w-auto px-2 sm:px-3 py-2 border border-gray-200 rounded-lg text-xs sm:text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                  data-testid="select-filter-priority"
                >
                  <option value="all">All Priority</option>
                  <option value="High">High</option>
                  <option value="Medium">Medium</option>
                  <option value="Low">Low</option>
                </select>
                <select
                  value={filterType}
                  onChange={(e) => setFilterType(e.target.value)}
                  className="w-full sm:w-auto px-2 sm:px-3 py-2 border border-gray-200 rounded-lg text-xs sm:text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                  data-testid="select-filter-type"
                >
                  <option value="all">All Types</option>
                  <option value="Safety Ideas">Safety Ideas</option>
                  <option value="Business Ideas">Business Ideas</option>
                  <option value="Near Miss">Near Miss</option>
                </select>
                <select
                  value={filterAssignedTo}
                  onChange={(e) => setFilterAssignedTo(e.target.value)}
                  className="w-full sm:w-auto px-2 sm:px-3 py-2 border border-gray-200 rounded-lg text-xs sm:text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                  data-testid="select-filter-assigned-to"
                >
                  <option value="all">All People</option>
                  <option value="">Unassigned</option>
                  {uniqueNames.map(name => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        </div>

        {sortedItems.length === 0 ? (
          <div className="bg-white rounded-lg shadow-sm border border-gray-100 p-12 text-center">
            <Target className="h-16 w-16 text-gray-300 mx-auto mb-4" />
            <h3 className="text-lg font-medium text-gray-900 mb-2">No Actions Found</h3>
            <p className="text-gray-500 mb-4">
              {actionItems.length > 0
                ? 'Try adjusting your filters to find more actions.'
                : 'Actions will appear here when items are set to "Actioned" status in Meeting History.'}
            </p>
            {actionItems.length > 0 && (
              <Button
                variant="outline"
                onClick={() => {
                  setSearchQuery('');
                  setFilterStatus('open');
                  setFilterPriority('all');
                  setFilterType('all');
                  setFilterAssignedTo('all');
                }}
              >
                Clear Filters
              </Button>
            )}
          </div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3">
            {sortedItems.map((item) => {
              const dueDateStatus = getDueDateStatus(item.actionDueDate);
              const isCompleted = item.actionStatus === 'Completed' || item.actionStatus === 'Ready to Close';

              // Full class names written out so Tailwind JIT includes them
              const borderColor =
                isCompleted           ? 'border-l-green-500' :
                item.actionPriority === 'High'   ? 'border-l-red-500'   :
                item.actionPriority === 'Medium' ? 'border-l-amber-400' :
                item.actionPriority === 'Low'    ? 'border-l-green-400' : 'border-l-gray-300';

              return (
                <button
                  key={item.id}
                  onClick={() => setSelectedItem(item)}
                  className={`text-left bg-white rounded-lg shadow-sm border border-gray-100 border-l-4 ${borderColor} hover:shadow-md transition-all focus:outline-none focus:ring-2 focus:ring-amber-400 ${isCompleted ? 'opacity-60' : ''} h-36 flex flex-col`}
                  data-testid={`card-action-${item.id}`}
                >
                  <div className="p-3 flex flex-col flex-1 min-h-0">
                    {/* Category + status row — always at top */}
                    <div className="flex items-center justify-between gap-1 mb-2">
                      {getTypeBadge(item.type)}
                      <Badge variant="outline" className={`h-5 px-1.5 py-0 leading-none rounded-full text-[10px] font-medium ${getStatusColor(item.actionStatus)}`}>
                        {item.actionStatus || 'Not Started'}
                      </Badge>
                    </div>

                    {/* Title — takes up middle space */}
                    <h3 className={`font-semibold text-sm leading-snug flex-1 overflow-hidden line-clamp-2 ${isCompleted ? 'text-green-800 line-through' : 'text-gray-900'}`}>
                      {item.title || 'Untitled Action'}
                    </h3>

                    {/* Assigned + due date — always pinned to bottom */}
                    <div className="flex flex-col gap-0.5 pt-2 border-t border-gray-100 mt-2">
                      <span className="inline-flex items-center gap-1 text-[11px] text-gray-500 h-4">
                        {item.actionAssignedTo ? (
                          <>
                            <User className="h-3 w-3 shrink-0 text-gray-400" />
                            <span className="truncate">{item.actionAssignedTo}</span>
                          </>
                        ) : (
                          <span className="text-gray-300">Unassigned</span>
                        )}
                      </span>
                      <span className={`inline-flex items-center gap-1 text-[11px] h-4 ${dueDateStatus && !isCompleted ? `font-medium ${dueDateStatus.color}` : 'text-gray-400'}`}>
                        <Clock className="h-3 w-3 shrink-0" />
                        {dueDateStatus && !isCompleted
                          ? dueDateStatus.label
                          : item.actionDueDate
                            ? formatDate(item.actionDueDate)
                            : <span className="text-gray-300">No due date</span>}
                      </span>
                    </div>
                  </div>
                </button>
              );
            })}
          </div>
        )}

      {/* Action detail modal */}
      {selectedItem && (() => {
        const item = selectedItem;
        return (
          <Dialog open={!!selectedItem} onOpenChange={(open) => { if (!open) setSelectedItem(null); }}>
            <DialogContent className="sm:max-w-2xl max-h-[90vh] overflow-y-auto">
              <DialogHeader>
                <div className="mb-1">{getTypeBadge(item.type)}</div>
                <DialogTitle className="leading-snug pr-6">
                  {item.title || 'Untitled Action'}
                </DialogTitle>
              </DialogHeader>

              <div className="space-y-4 py-2">
                {item.description && (
                  <p className="text-sm text-gray-600">{item.description}</p>
                )}

                {/* Quick-complete banner */}
                {item.actionStatus !== 'Completed' && (
                  <Button
                    className="w-full bg-green-600 hover:bg-green-700 text-white flex items-center justify-center gap-2"
                    onClick={() => {
                      updateActionFields(item, { actionStatus: 'Completed' });
                      setSelectedItem({ ...item, actionStatus: 'Completed' });
                    }}
                    disabled={isUpdatingAction === item.id}
                  >
                    <CheckCircle className="h-4 w-4" />
                    Mark as Completed
                  </Button>
                )}

                {item.actionStatus === 'Completed' && (
                  <div className="flex items-center gap-2 px-3 py-2 bg-green-50 border border-green-200 rounded-lg text-sm text-green-700">
                    <CheckCircle className="h-4 w-4" />Completed
                  </div>
                )}

                {/* Editable fields */}
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-[10px] font-medium text-gray-500 uppercase tracking-wide">Priority</label>
                    <select
                      value={item.actionPriority || ''}
                      onChange={(e) => {
                        const priority = e.target.value;
                        let dueDate = item.actionDueDate || '';
                        if (priority && !item.actionDueDate) {
                          const today = new Date();
                          today.setDate(today.getDate() + (priority === 'High' ? 7 : priority === 'Medium' ? 14 : 30));
                          dueDate = today.toISOString().split('T')[0];
                        }
                        updateActionFields(item, { actionPriority: priority, actionDueDate: dueDate });
                        setSelectedItem({ ...item, actionPriority: priority, actionDueDate: dueDate });
                        if (priority) postActivityEntry(item, 'priority', `Priority set to ${priority}`);
                      }}
                      className="w-full mt-0.5 text-sm border border-gray-200 rounded px-2 h-9 bg-white focus:border-amber-400 focus:outline-none"
                    >
                      <option value="">No Priority</option>
                      <option value="High">High</option>
                      <option value="Medium">Medium</option>
                      <option value="Low">Low</option>
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-medium text-gray-500 uppercase tracking-wide">Status</label>
                    <select
                      value={item.actionStatus || ''}
                      onChange={(e) => {
                        const newStatus = e.target.value;
                        updateActionFields(item, { actionStatus: newStatus });
                        setSelectedItem({ ...item, actionStatus: newStatus });
                        if (newStatus) postActivityEntry(item, 'status', `Status changed to ${newStatus || 'Not Started'}`);
                      }}
                      className="w-full mt-0.5 text-sm border border-gray-200 rounded px-2 h-9 bg-white focus:border-amber-400 focus:outline-none"
                    >
                      <option value="">Not Started</option>
                      <option value="In Progress">In Progress</option>
                      <option value="On Hold">On Hold</option>
                      <option value="Completed">Completed</option>
                      <option value="Ready to Close">Ready to Close</option>
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-medium text-gray-500 uppercase tracking-wide">Assigned To</label>
                    <select
                      value={getActionFieldValue(item, 'actionAssignedTo')}
                      onChange={(e) => {
                        handleActionTextChange(item, 'actionAssignedTo', e.target.value);
                        setSelectedItem({ ...item, actionAssignedTo: e.target.value });
                      }}
                      className="w-full mt-0.5 text-sm border border-gray-200 rounded px-2 h-9 bg-white focus:border-amber-400 focus:outline-none"
                    >
                      <option value="">Unassigned</option>
                      {uniqueNames.map(name => (
                        <option key={name} value={name}>{name}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-medium text-gray-500 uppercase tracking-wide">Due Date</label>
                    <input
                      type="date"
                      value={item.actionDueDate ? item.actionDueDate.split('T')[0] : ''}
                      onChange={(e) => {
                        const newDate = e.target.value;
                        updateActionFields(item, { actionDueDate: newDate });
                        setSelectedItem({ ...item, actionDueDate: newDate });
                        if (newDate) {
                          try { postActivityEntry(item, 'due_date', `Due date set to ${format(new Date(newDate), 'dd MMM yyyy')}`); } catch { /* silent */ }
                        }
                      }}
                      className="w-full mt-0.5 text-sm border border-gray-200 rounded px-2 h-9 bg-white focus:border-amber-400 focus:outline-none"
                    />
                  </div>
                </div>

                {/* Action Notes */}
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-[10px] font-medium text-gray-500 uppercase tracking-wide flex items-center gap-1">
                      Action Notes <i className="fas fa-robot text-green-500 opacity-70"></i>
                    </label>
                    <div className="flex items-center gap-2">
                      {localActionEdits[item.id]?.actionNotes !== undefined && (
                        <span className="text-xs text-amber-600 italic">Editing…</span>
                      )}
                      {recentlySavedActions.has(item.id) && !localActionEdits[item.id]?.actionNotes && (
                        <span className="text-xs text-green-600 flex items-center gap-1"><i className="fas fa-check text-[10px]"></i>Saved</span>
                      )}
                      <button
                        type="button"
                        onClick={async () => {
                          const currentNotes = getActionFieldValue(item, 'actionNotes');
                          if (!currentNotes.trim()) return;
                          setIsEnhancingActionNotes(item.id);
                          try {
                            const response = await fetch('/api/ai-enhance-notes', {
                              method: 'POST',
                              headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${await authService.getAccessToken()}` },
                              body: JSON.stringify({ content: currentNotes, itemType: 'Action Notes' })
                            });
                            if (response.ok) {
                              const data = await response.json();
                              setLocalActionEdits(prev => ({ ...prev, [item.id]: { ...prev[item.id], actionNotes: data.enhancedContent } }));
                            }
                          } catch (e) {
                            console.error('AI enhancement failed:', e);
                          } finally {
                            setIsEnhancingActionNotes('');
                          }
                        }}
                        disabled={!getActionFieldValue(item, 'actionNotes').trim() || isEnhancingActionNotes === item.id}
                        className="text-xs bg-green-600 hover:bg-green-700 disabled:bg-gray-200 disabled:text-gray-400 text-white px-2 py-0.5 rounded flex items-center gap-1"
                      >
                        {isEnhancingActionNotes === item.id
                          ? <><i className="fas fa-spinner fa-spin text-[10px]"></i>Enhancing...</>
                          : <><i className="fas fa-wand-magic-sparkles text-[10px]"></i>Finish Text</>}
                      </button>
                    </div>
                  </div>
                  <textarea
                    value={getActionFieldValue(item, 'actionNotes')}
                    onChange={(e) => handleActionTextChange(item, 'actionNotes', e.target.value)}
                    placeholder="Add action notes..."
                    rows={3}
                    className="w-full text-sm border border-gray-200 rounded px-3 py-2 bg-gray-50 focus:bg-white focus:border-amber-400 focus:outline-none resize-none"
                  />
                  <div className="flex justify-end mt-1">
                    <button
                      type="button"
                      disabled={localActionEdits[item.id]?.actionNotes === undefined || isUpdatingAction === item.id}
                      onClick={() => {
                        const notes = localActionEdits[item.id]?.actionNotes ?? '';
                        updateActionFields(item, { actionNotes: notes }, 'actionNotes');
                        if (notes.trim()) postActivityEntry(item, 'note', notes.trim());
                      }}
                      className="text-xs bg-amber-600 hover:bg-amber-700 disabled:bg-gray-200 disabled:text-gray-400 text-white px-3 py-1 rounded flex items-center gap-1"
                    >
                      {isUpdatingAction === item.id
                        ? <><i className="fas fa-spinner fa-spin text-[10px]"></i>Saving…</>
                        : <><i className="fas fa-save text-[10px]"></i>Save</>}
                    </button>
                  </div>
                </div>

                {/* Meeting discussion */}
                {item.meetingNotes && item.meetingNotes.trim() && (
                  <div className="bg-blue-50 rounded-lg p-3">
                    <p className="text-[10px] font-medium text-blue-600 uppercase tracking-wide mb-1">Meeting Discussion</p>
                    <p className="text-sm text-gray-700">{item.meetingNotes}</p>
                  </div>
                )}

                {/* Activity timeline */}
                <div>
                  <p className="text-[10px] font-medium text-gray-500 uppercase tracking-wide mb-2 flex items-center gap-1">
                    <i className="fas fa-history text-[10px]"></i> Activity
                  </p>
                  {loadingActivity === item.id ? (
                    <p className="text-xs text-gray-400 italic">Loading activity…</p>
                  ) : (activityLogs[item.id] || []).length === 0 ? (
                    <p className="text-xs text-gray-400 italic">No activity yet. Changes and notes you save will appear here.</p>
                  ) : (
                    <ol className="relative border-l border-gray-200 ml-2 space-y-0">
                      {[...(activityLogs[item.id] || [])].reverse().map((entry) => {
                        const icons: Record<string, string> = {
                          note: 'fas fa-comment text-amber-500',
                          status: 'fas fa-arrow-right-arrow-left text-blue-500',
                          priority: 'fas fa-flag text-red-500',
                          due_date: 'fas fa-calendar-day text-purple-500',
                          assigned: 'fas fa-user text-green-500',
                          start_date: 'fas fa-play text-gray-500',
                        };
                        const iconClass = icons[entry.entryType] || 'fas fa-circle text-gray-400';
                        const when = (() => {
                          try {
                            const d = new Date(entry.createdAt);
                            const diff = Date.now() - d.getTime();
                            if (diff < 60000) return 'just now';
                            if (diff < 3600000) return `${Math.floor(diff / 60000)}m ago`;
                            if (diff < 86400000) return `${Math.floor(diff / 3600000)}h ago`;
                            if (diff < 604800000) return `${Math.floor(diff / 86400000)}d ago`;
                            return format(d, 'dd MMM yyyy');
                          } catch { return ''; }
                        })();
                        const fullDate = (() => {
                          try { return format(new Date(entry.createdAt), 'dd MMM yyyy, h:mm a'); } catch { return ''; }
                        })();
                        return (
                          <li key={entry.id} className="ml-4 py-2 border-b border-gray-50 last:border-0">
                            <span className="absolute -left-[9px] flex items-center justify-center w-4 h-4 bg-white border border-gray-200 rounded-full mt-0.5">
                              <i className={`${iconClass} text-[8px]`}></i>
                            </span>
                            <div className="flex items-start justify-between gap-2">
                              <p className={`text-sm text-gray-800 ${entry.entryType === 'note' ? 'whitespace-pre-wrap' : 'italic text-gray-600'}`}>
                                {entry.content}
                              </p>
                              <span title={fullDate} className="text-[10px] text-gray-400 whitespace-nowrap shrink-0 mt-0.5 cursor-default">
                                {when}
                              </span>
                            </div>
                            {entry.author && (
                              <p className="text-[10px] text-gray-400 mt-0.5">{entry.author}</p>
                            )}
                          </li>
                        );
                      })}
                    </ol>
                  )}
                </div>

                {/* Metadata footer */}
                <div className="flex flex-wrap gap-3 text-xs text-gray-400 pt-1 border-t border-gray-100">
                  <span className="flex items-center gap-1"><User className="h-3 w-3" />Submitted by {item.submittedBy}</span>
                  <span className="flex items-center gap-1"><Calendar className="h-3 w-3" />Submitted {formatDate(item.submittedDate)}</span>
                  <span className="flex items-center gap-1"><Calendar className="h-3 w-3" />Meeting {formatDate(item.meetingDate)}</span>
                </div>
              </div>
            </DialogContent>
          </Dialog>
        );
      })()}
        </div>
      </div>

      <Dialog open={showExportModal} onOpenChange={setShowExportModal}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <Download className="h-5 w-5 text-amber-600" />
              Export Actions
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4 py-4">
            <p className="text-sm text-gray-600">
              Export {sortedItems.length} action{sortedItems.length !== 1 ? 's' : ''} to your preferred format.
            </p>
            <div className="grid grid-cols-2 gap-3">
              <Button
                variant="outline"
                onClick={() => {
                  exportToHTML();
                  setShowExportModal(false);
                }}
                className="flex flex-col items-center gap-2 h-auto py-4"
              >
                <FileText className="h-8 w-8 text-blue-600" />
                <span className="text-sm font-medium">HTML Report</span>
                <span className="text-xs text-gray-500">Professional format</span>
              </Button>
              <Button
                variant="outline"
                onClick={() => {
                  exportToCSV();
                  setShowExportModal(false);
                }}
                className="flex flex-col items-center gap-2 h-auto py-4"
              >
                <FileText className="h-8 w-8 text-green-600" />
                <span className="text-sm font-medium">CSV Spreadsheet</span>
                <span className="text-xs text-gray-500">Excel compatible</span>
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </div>
  );
}
