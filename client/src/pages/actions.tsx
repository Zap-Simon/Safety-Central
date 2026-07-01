import { useState, useRef, useEffect } from "react";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { authService } from "@/auth/authService";
import MeetingHeader from "@/components/meeting-header";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Calendar, FileText, AlertTriangle, Lightbulb, Shield, CheckCircle, Download, Search, Clock, User, Target, ClipboardList, Loader2 } from "lucide-react";
import { format } from "date-fns";
import NearMissInvestigationModal from "@/components/near-miss/NearMissInvestigationModal";
import ActionStatusWorkflow from "@/components/ActionStatusWorkflow";
import { getDateGroupKey } from "@shared/dateUtils";

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
  reconsiderDate?: string;
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
  const [showRegisterModal, setShowRegisterModal] = useState(false);
  const [isExportingRegister, setIsExportingRegister] = useState<string>('');
  const [registerDateFrom, setRegisterDateFrom] = useState<string>('');
  const [registerDateTo, setRegisterDateTo] = useState<string>('');
  const [selectedItem, setSelectedItem] = useState<ActionableItem | null>(null);
  const [activityLogs, setActivityLogs] = useState<Record<string, ActivityEntry[]>>({});
  const [loadingActivity, setLoadingActivity] = useState<string>('');
  const [investigationItem, setInvestigationItem] = useState<ActionableItem | null>(null);
  const [highlightedId, setHighlightedId] = useState<string | null>(null);
  const [focusedItemId, setFocusedItemId] = useState<string | null>(null);
  const cardRefs = useRef<Record<string, HTMLButtonElement | null>>({});
  const deepLinkHandled = useRef(false);

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

  const FORMER_EMPLOYEES = ['Wayne', 'Teresa'];

  const uniqueNames: string[] = Array.from(new Set(
    meetingItems.flatMap(i => [i.submittedBy, i.actionAssignedTo].filter(Boolean) as string[])
  ))
    .filter(name => !FORMER_EMPLOYEES.some(ex => name.toLowerCase().includes(ex.toLowerCase())))
    .sort();

  // True when an item has any action-management data attached (assigned, prioritised,
  // given a status / notes / dates, or explicitly marked "Actioned" in SharePoint).
  const hasActionData = (item: ActionableItem) => !!(
    item.actionPriority ||
    item.actionStatus ||
    item.actionAssignedTo ||
    item.actionStartDate ||
    item.actionDueDate ||
    item.actionNotes ||
    item.status === 'Actioned'
  );

  // Every Near Miss stays on this page as a permanent safety record, even when no
  // action was ever raised. An "old" near miss with no action taken is shown as a
  // closed record — visible for the register, but kept out of the live workload.
  const isNearMissRecord = (item: ActionableItem) =>
    item.type === 'Near Miss' && !hasActionData(item);

  // Actionable items PLUS every near miss (so the full near-miss register is always
  // present on the page, not just the ones that became tracked actions).
  const actionItems = meetingItems.filter(item =>
    item.type === 'Near Miss' || hasActionData(item)
  );

  // A finished/archived action is one that's either fully Completed or sitting in
  // "Ready to Close" awaiting formal closure. These are kept out of the live
  // workload numbers so the dashboard reflects only actions still to manage.
  const isArchivedStatus = (status?: string) =>
    status === 'Completed' || status === 'Ready to Close';

  // An item whose OVERALL status is "Closed" was finished from Meeting Minutes
  // (the "bypass actions" path). Closing can leave action data behind (e.g. the
  // closure-notes prompt writes actionNotes), so we must key off the item's own
  // status — not the presence of action data — to keep it out of the live list
  // and workload counts. It stays in `actionItems` so the focused deep-link from
  // Meeting Minutes can still surface the single closed item for the record.
  const isClosedItem = (item: ActionableItem) => item.status === 'Closed';

  // A Near Miss drops out of the live "open (to manage)" workload once it's been
  // formally closed from Meeting Minutes, its action is finished (Completed /
  // Ready to Close), or it never became a tracked action (record-only). It still
  // shows under the "All" filter as a permanent safety record. This mirrors how
  // the dashboard "open" count already treats near misses (see stats.open).
  const isNearMissOutOfWorkload = (item: ActionableItem) =>
    item.type === 'Near Miss' &&
    (isClosedItem(item) || isArchivedStatus(item.actionStatus) || isNearMissRecord(item));

  // Live actions still to manage — closed/finished/record items are excluded so
  // the dashboard numbers reflect only open work, consistent with how
  // Completed / Ready-to-Close are kept out of the live workload.
  const liveActionItems = actionItems.filter(item =>
    !isClosedItem(item) && !isNearMissOutOfWorkload(item)
  );

  const filteredItems = actionItems.filter(item => {
    // Deep link from Minutes focuses a single action — show only that card.
    // This intentionally runs before the closed-item filter so a closed item
    // opened from Meeting Minutes still surfaces as its single focused card.
    if (focusedItemId) return item.id === focusedItemId;

    // Outside the focused deep-link, closed items never show as live/open actions —
    // except Near Misses, which always stay visible as a permanent safety record.
    if (isClosedItem(item) && item.type !== 'Near Miss') return false;

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

    // The status toggle governs the view. Only the "All" filter shows the full
    // register — every other filter (including the default "Open (to manage)")
    // hides near misses that are closed, finished, or record-only so they don't
    // clutter the live workload. They remain a permanent record under "All".
    if (filterStatus !== 'all') {
      const itemStatus = item.actionStatus || 'Not Started';
      // For near misses, "finished" also covers formal closure and record-only
      // items, not just an archived action status.
      const isArchived = item.type === 'Near Miss'
        ? isNearMissOutOfWorkload(item)
        : isArchivedStatus(itemStatus);
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

  // Deep link from Minutes (Meeting History): a per-item "Actions" button opens
  // /actions?itemId=…&type=… in a new tab. Focus that exact action — clear the
  // default filters so it can't be hidden, scroll to it, highlight it, and open
  // its detail (the investigation modal for Near Miss, the standard modal otherwise).
  useEffect(() => {
    if (deepLinkHandled.current) return;
    if (meetingItems.length === 0) return;

    const params = new URLSearchParams(window.location.search);
    const itemId = params.get('itemId');
    const itemType = params.get('type');
    if (!itemId) {
      deepLinkHandled.current = true;
      return;
    }

    const target = actionItems.find(
      (i) => i.id === itemId && (!itemType || i.type === itemType)
    );
    // Data may still be settling — wait for the next data change before giving up.
    if (!target) return;

    deepLinkHandled.current = true;

    // Show ONLY the targeted action by filtering the list down to it. This avoids
    // opening a detail modal (which triggers an MSAL prompt) and the awkward
    // scroll-into-view. Clear other filters/search so they can't interfere.
    setFilterStatus('all');
    setFilterPriority('all');
    setFilterType('all');
    setFilterAssignedTo('all');
    setSearchQuery('');
    setFocusedItemId(target.id);
  }, [meetingItems]);

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
      reconsiderDate?: string;
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
        reconsiderDate: updates.reconsiderDate ?? item.reconsiderDate ?? '',
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
    total: filteredItems.length,
    open: filteredItems.filter(i => !isArchivedStatus(i.actionStatus) && !isNearMissRecord(i)).length,
    completed: filteredItems.filter(i => i.actionStatus === 'Completed').length,
    overdue: filteredItems.filter(i => {
      if (!i.actionDueDate || isArchivedStatus(i.actionStatus)) return false;
      return new Date(i.actionDueDate) < new Date();
    }).length,
    highPriority: filteredItems.filter(i => i.actionPriority === 'High' && !isArchivedStatus(i.actionStatus)).length
  };

  // Exports now run on the shared, professional server export engine (the same
  // one the meeting minutes use) so the Actions report is Cranfield-branded,
  // A4 print-ready with "Page X of Y" footers, and carries the rich action data
  // (lifecycle, due-date analytics, investigation details, activity history).
  // The current filters/sort are honoured by sending `sortedItems` — what the
  // user sees is exactly what gets exported.
  const [isExporting, setIsExporting] = useState<string>('');

  const downloadBlob = (blob: Blob, filename: string) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportActions = async (format: 'html' | 'csv' | 'markdown' | 'word') => {
    setIsExporting(format);
    try {
      const response = await fetch(`/api/generate-actions-${format}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items: sortedItems, stats }),
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);

      if (format === 'html') {
        const result = await response.json();
        if (!result.success) throw new Error(result.error || 'Failed to generate report');
        downloadBlob(new Blob([result.htmlContent], { type: 'text/html' }), result.filename || 'Actions-Report.html');
      } else {
        const blob = await response.blob();
        const disposition = response.headers.get('content-disposition');
        const match = disposition?.match(/filename="(.+)"/);
        const fallback: Record<string, string> = {
          csv: 'Actions.csv', markdown: 'Actions.md', word: 'Actions.docx',
        };
        downloadBlob(blob, match ? match[1] : fallback[format]);
      }
      setShowExportModal(false);
    } catch (error) {
      console.error(`Failed to export actions (${format}):`, error);
    } finally {
      setIsExporting('');
    }
  };

  // Near Miss Register: EVERY near miss card (not just the ones that became
  // tracked actions like the Actions report shows) turned into a branded
  // register, enriched server-side with its full investigation. Built on the
  // same export engine so it matches the meeting minutes / Actions report look.
  const nearMissItems = meetingItems.filter(item => item.type === 'Near Miss');

  // Near misses falling inside the optional From/To date band (by the date the near
  // miss was submitted/reported). Empty band = every near miss. Both the inputs and
  // getDateGroupKey produce a NZ-aligned yyyy-mm-dd key, so a plain string compare is
  // inclusive of both ends and immune to the viewer's local timezone.
  const registerItems = nearMissItems
    .filter(item => {
      if (!registerDateFrom && !registerDateTo) return true;
      const key = getDateGroupKey(item.submittedDate);
      if (key === 'unknown-meeting') return false;
      if (registerDateFrom && key < registerDateFrom) return false;
      if (registerDateTo && key > registerDateTo) return false;
      return true;
    })
    // Most recently submitted near miss first. The server preserves the order it
    // receives, so this client-side sort is what drives the register's row order.
    .sort((a, b) => {
      const aTime = a.submittedDate ? new Date(a.submittedDate).getTime() : 0;
      const bTime = b.submittedDate ? new Date(b.submittedDate).getTime() : 0;
      return bTime - aTime;
    });

  const exportRegister = async (format: 'html' | 'csv' | 'markdown' | 'word') => {
    setIsExportingRegister(format);
    try {
      const response = await fetch(`/api/generate-near-miss-register-${format}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items: registerItems, dateFrom: registerDateFrom || undefined, dateTo: registerDateTo || undefined }),
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);

      if (format === 'html') {
        const result = await response.json();
        if (!result.success) throw new Error(result.error || 'Failed to generate register');
        downloadBlob(new Blob([result.htmlContent], { type: 'text/html' }), result.filename || 'Near-Miss-Register.html');
      } else {
        const blob = await response.blob();
        const disposition = response.headers.get('content-disposition');
        const match = disposition?.match(/filename="(.+)"/);
        const fallback: Record<string, string> = {
          csv: 'Near-Miss-Register.csv', markdown: 'Near-Miss-Register.md', word: 'Near-Miss-Register.docx',
        };
        downloadBlob(blob, match ? match[1] : fallback[format]);
      }
      setShowRegisterModal(false);
    } catch (error) {
      console.error(`Failed to export near miss register (${format}):`, error);
    } finally {
      setIsExportingRegister('');
    }
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
                onClick={() => setShowRegisterModal(true)}
                className="flex items-center gap-2 border-orange-200 text-orange-700 hover:bg-orange-50 hover:text-orange-800"
                data-testid="button-near-miss-register"
              >
                <AlertTriangle className="h-4 w-4" />
                <span>Near Miss Register</span>
              </Button>
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

        {focusedItemId && (
          <div className="bg-amber-50 border border-amber-200 rounded-lg px-4 py-3 mb-4 flex items-center justify-between gap-3">
            <div className="flex items-center gap-2 text-sm text-amber-800 min-w-0">
              <Target className="h-4 w-4 flex-shrink-0 text-amber-600" />
              <span className="truncate">Showing a single action opened from Meeting Minutes.</span>
            </div>
            <Button
              variant="outline"
              size="sm"
              className="flex-shrink-0 bg-white"
              onClick={() => setFocusedItemId(null)}
              data-testid="button-show-all-actions"
            >
              Show all actions
            </Button>
          </div>
        )}

        {sortedItems.length === 0 ? (
          <div className="bg-white rounded-lg shadow-sm border border-gray-100 p-12 text-center">
            <Target className="h-16 w-16 text-gray-300 mx-auto mb-4" />
            <h3 className="text-lg font-medium text-gray-900 mb-2">No Actions Found</h3>
            <p className="text-gray-500 mb-4">
              {liveActionItems.length > 0
                ? 'Try adjusting your filters to find more actions.'
                : 'Actions will appear here when items are set to "Actioned" status in Meeting History.'}
            </p>
            {liveActionItems.length > 0 && (
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
                  ref={(el) => { cardRefs.current[item.id] = el; }}
                  onClick={() => item.type === 'Near Miss' ? setInvestigationItem(item) : setSelectedItem(item)}
                  className={`text-left bg-white rounded-lg shadow-sm border border-gray-100 border-l-4 ${borderColor} hover:shadow-md transition-all focus:outline-none focus:ring-2 focus:ring-amber-400 ${isCompleted ? 'opacity-60' : ''} ${highlightedId === item.id ? 'ring-2 ring-amber-500 ring-offset-2 shadow-lg' : ''} h-36 flex flex-col`}
                  data-testid={`card-action-${item.id}`}
                >
                  <div className="p-3 flex flex-col flex-1 min-h-0">
                    {/* Category + status row — always at top */}
                    <div className="flex items-center justify-between gap-1 mb-2">
                      {getTypeBadge(item.type)}
                      <div className="flex items-center gap-1">
                        {item.type === 'Near Miss' && !isNearMissRecord(item) && (
                          <Badge className="h-5 px-1.5 py-0 leading-none rounded-full text-[10px] font-medium bg-orange-600 text-white flex items-center gap-0.5">
                            <ClipboardList className="h-2.5 w-2.5" />Investigate
                          </Badge>
                        )}
                        {isNearMissRecord(item) ? (
                          <Badge variant="outline" className="h-5 px-1.5 py-0 leading-none rounded-full text-[10px] font-medium bg-gray-100 text-gray-500 border-gray-200">
                            Closed
                          </Badge>
                        ) : (
                          <Badge variant="outline" className={`h-5 px-1.5 py-0 leading-none rounded-full text-[10px] font-medium ${getStatusColor(item.actionStatus)}`}>
                            {item.actionStatus || 'Not Started'}
                          </Badge>
                        )}
                      </div>
                    </div>

                    {/* Title — takes up middle space */}
                    <h3 className={`font-semibold text-sm leading-snug flex-1 overflow-hidden line-clamp-2 ${isCompleted ? 'text-green-800' : 'text-gray-900'} ${item.actionStatus === 'Completed' ? 'line-through' : ''}`}>
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

                {/* Status workflow */}
                <div>
                  <label className="text-[10px] font-medium text-gray-500 uppercase tracking-wide mb-1.5 block">Status</label>
                  <ActionStatusWorkflow
                    status={item.actionStatus || ''}
                    reconsiderDate={item.reconsiderDate}
                    allowComplete={item.type !== 'Near Miss'}
                    onChange={(updates) => {
                      updateActionFields(item, updates);
                      setSelectedItem({ ...item, ...updates });
                      if (updates.actionStatus !== undefined) {
                        postActivityEntry(item, 'status', `Status changed to ${updates.actionStatus || 'Not Started'}`);
                      }
                      if (updates.reconsiderDate) {
                        try { postActivityEntry(item, 'status', `On hold — reconsider on ${format(new Date(updates.reconsiderDate), 'dd MMM yyyy')}`); } catch { /* silent */ }
                      }
                    }}
                  />
                </div>

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
                          : <><i className="fas fa-wand-magic-sparkles text-[10px]"></i>Improve Writing</>}
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
              Reflects your current filters and sort order.
            </p>
            <div className="grid grid-cols-2 gap-3">
              <Button
                variant="outline"
                disabled={!!isExporting}
                onClick={() => exportActions('html')}
                className="flex flex-col items-center gap-2 h-auto py-4"
              >
                {isExporting === 'html'
                  ? <Loader2 className="h-8 w-8 text-blue-600 animate-spin" />
                  : <FileText className="h-8 w-8 text-blue-600" />}
                <span className="text-sm font-medium">HTML / PDF</span>
                <span className="text-xs text-gray-500">Print-ready report</span>
              </Button>
              <Button
                variant="outline"
                disabled={!!isExporting}
                onClick={() => exportActions('word')}
                className="flex flex-col items-center gap-2 h-auto py-4"
              >
                {isExporting === 'word'
                  ? <Loader2 className="h-8 w-8 text-indigo-600 animate-spin" />
                  : <FileText className="h-8 w-8 text-indigo-600" />}
                <span className="text-sm font-medium">Word Document</span>
                <span className="text-xs text-gray-500">Editable .docx</span>
              </Button>
              <Button
                variant="outline"
                disabled={!!isExporting}
                onClick={() => exportActions('csv')}
                className="flex flex-col items-center gap-2 h-auto py-4"
              >
                {isExporting === 'csv'
                  ? <Loader2 className="h-8 w-8 text-green-600 animate-spin" />
                  : <FileText className="h-8 w-8 text-green-600" />}
                <span className="text-sm font-medium">CSV Spreadsheet</span>
                <span className="text-xs text-gray-500">Excel compatible</span>
              </Button>
              <Button
                variant="outline"
                disabled={!!isExporting}
                onClick={() => exportActions('markdown')}
                className="flex flex-col items-center gap-2 h-auto py-4"
              >
                {isExporting === 'markdown'
                  ? <Loader2 className="h-8 w-8 text-gray-600 animate-spin" />
                  : <FileText className="h-8 w-8 text-gray-600" />}
                <span className="text-sm font-medium">Markdown</span>
                <span className="text-xs text-gray-500">Plain text format</span>
              </Button>
            </div>
            <p className="text-xs text-gray-400 text-center">
              Tip: open the HTML report and use your browser's Print to save as PDF.
            </p>
          </div>
        </DialogContent>
      </Dialog>

      <Dialog open={showRegisterModal} onOpenChange={setShowRegisterModal}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <AlertTriangle className="h-5 w-5 text-orange-600" />
              Near Miss Register
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4 py-4">
            <p className="text-sm text-gray-600">
              Export a register of your near misses, each with its full investigation,
              risk assessment and sign-off.
            </p>

            <div className="rounded-md border border-gray-200 p-3 space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-medium text-gray-700">Date range (optional)</span>
                {(registerDateFrom || registerDateTo) && (
                  <button
                    type="button"
                    onClick={() => { setRegisterDateFrom(''); setRegisterDateTo(''); }}
                    className="text-xs text-orange-600 hover:underline"
                    data-testid="button-clear-register-dates"
                  >
                    Clear
                  </button>
                )}
              </div>
              <div className="grid grid-cols-2 gap-2">
                <label className="flex flex-col gap-1 text-xs text-gray-500">
                  From
                  <input
                    type="date"
                    value={registerDateFrom}
                    max={registerDateTo || undefined}
                    onChange={(e) => setRegisterDateFrom(e.target.value)}
                    className="rounded-md border border-gray-300 px-2 py-1.5 text-sm text-gray-900 focus:border-orange-500 focus:outline-none focus:ring-1 focus:ring-orange-500"
                    data-testid="input-register-date-from"
                  />
                </label>
                <label className="flex flex-col gap-1 text-xs text-gray-500">
                  To
                  <input
                    type="date"
                    value={registerDateTo}
                    min={registerDateFrom || undefined}
                    onChange={(e) => setRegisterDateTo(e.target.value)}
                    className="rounded-md border border-gray-300 px-2 py-1.5 text-sm text-gray-900 focus:border-orange-500 focus:outline-none focus:ring-1 focus:ring-orange-500"
                    data-testid="input-register-date-to"
                  />
                </label>
              </div>
              <p className="text-xs text-gray-500" data-testid="text-register-count">
                {registerItems.length} near miss{registerItems.length !== 1 ? 'es' : ''}
                {(registerDateFrom || registerDateTo) ? ' in this range' : ' (all dates)'} will be included.
              </p>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <Button
                variant="outline"
                disabled={!!isExportingRegister || registerItems.length === 0}
                onClick={() => exportRegister('html')}
                className="flex flex-col items-center gap-2 h-auto py-4"
              >
                {isExportingRegister === 'html'
                  ? <Loader2 className="h-8 w-8 text-orange-600 animate-spin" />
                  : <FileText className="h-8 w-8 text-orange-600" />}
                <span className="text-sm font-medium">HTML / PDF</span>
                <span className="text-xs text-gray-500">Print-ready register</span>
              </Button>
              <Button
                variant="outline"
                disabled={!!isExportingRegister || registerItems.length === 0}
                onClick={() => exportRegister('word')}
                className="flex flex-col items-center gap-2 h-auto py-4"
              >
                {isExportingRegister === 'word'
                  ? <Loader2 className="h-8 w-8 text-indigo-600 animate-spin" />
                  : <FileText className="h-8 w-8 text-indigo-600" />}
                <span className="text-sm font-medium">Word Document</span>
                <span className="text-xs text-gray-500">Editable .docx</span>
              </Button>
              <Button
                variant="outline"
                disabled={!!isExportingRegister || registerItems.length === 0}
                onClick={() => exportRegister('csv')}
                className="flex flex-col items-center gap-2 h-auto py-4"
              >
                {isExportingRegister === 'csv'
                  ? <Loader2 className="h-8 w-8 text-green-600 animate-spin" />
                  : <FileText className="h-8 w-8 text-green-600" />}
                <span className="text-sm font-medium">CSV Spreadsheet</span>
                <span className="text-xs text-gray-500">Excel compatible</span>
              </Button>
              <Button
                variant="outline"
                disabled={!!isExportingRegister || registerItems.length === 0}
                onClick={() => exportRegister('markdown')}
                className="flex flex-col items-center gap-2 h-auto py-4"
              >
                {isExportingRegister === 'markdown'
                  ? <Loader2 className="h-8 w-8 text-gray-600 animate-spin" />
                  : <FileText className="h-8 w-8 text-gray-600" />}
                <span className="text-sm font-medium">Markdown</span>
                <span className="text-xs text-gray-500">Plain text format</span>
              </Button>
            </div>
            <p className="text-xs text-gray-400 text-center">
              Includes every near miss, even those without a tracked action.
            </p>
          </div>
        </DialogContent>
      </Dialog>

      {/* Near Miss Investigation Modal */}
      {investigationItem && (
        <NearMissInvestigationModal
          open={!!investigationItem}
          item={{
            id: investigationItem.id,
            title: investigationItem.title,
            description: investigationItem.description || "",
            secondaryDescription: investigationItem.secondaryDescription,
            submittedBy: investigationItem.submittedBy || "",
            meetingDate: investigationItem.meetingDate || "",
            meetingNotes: investigationItem.meetingNotes,
            actionNotes: investigationItem.actionNotes,
            ideaType: investigationItem.ideaType,
          }}
          onClose={() => setInvestigationItem(null)}
        />
      )}
    </div>
  );
}
