import { useState, useEffect, useRef } from "react";
import { useLocation } from "wouter";
import { useQuery, useQueryClient, useMutation } from "@tanstack/react-query";
import { authService } from "@/auth/authService";
import MeetingHeader from "@/components/meeting-header";
import DashboardStats from "@/components/DashboardStats";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from "@/components/ui/collapsible";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Progress } from "@/components/ui/progress";
import { ChevronDown, ChevronRight, Calendar, Users, FileText, AlertTriangle, Lightbulb, Shield, Bot, Loader2, LogIn, UserCheck, ExternalLink, ArrowRight, CalendarX, CheckCircle, CheckCircle2, Plus, Lock, Unlock, PenLine } from "lucide-react";
import SignatureCarousel from "@/components/SignatureCarousel";
import { parseSharePointDate, formatDisplayDate, getDateGroupKey, isSameDay, getMeetingStatus } from "@shared/dateUtils";
import { predictiveText } from "@/lib/predictiveText";
import { InlineTextarea } from "@/components/ui/inline-textarea";
import { format } from "date-fns";

// Sample data structure based on your SharePoint lists
interface MeetingItem {
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
  secondaryDescription?: string; // For Near Miss "How It Happened" field

  hasActualMeetingDate?: boolean;
  ideaType?: string; // Choice column values like "Waste Reduce", "Safety Improvement", etc.
  
  // Action-related properties
  actionAssignedTo?: string;
  actionStatus?: string;
  actionPriority?: string;
  actionStartDate?: string;
  actionDueDate?: string;
  actionNotes?: string;
}

// Actions are separate items that link to existing records
interface ActionItem {
  id: string;
  title: string;
  description: string;
  type: 'Actions';
  status: string;
  meetingDate: string;
  meetingNotes: string;
  submittedBy: string;
  submittedDate: string;
  assignedTo?: string;
  
  // Actions-specific fields
  priority?: string; // High, Medium, Low
  category?: string; // Safety Idea, Business Idea, etc.
  listOrigin?: string; // Source list identifier
  sourceLink?: string; // Link back to original item
  actionStartDate?: string; // When action work began
  actionCompletionDate?: string; // When action was completed (field not yet available)
  outcomeResult?: string; // Final outcome and results (field not yet available)
}

// Meeting attendees data structure
const meetingAttendees = {
  management: [
    { name: 'Hoani Hunt', role: 'Company Director' },
    { name: 'Simon Hubbard', role: 'Health & Safety Coordinator' },
    { name: 'James Waites', role: 'Glazing Supervisor' },
    { name: 'Emma White', role: 'Administrator' }
  ],
  glaziers: [
    { name: 'Kevin Young', role: 'Glazier' },
    { name: 'Ryan Newman', role: 'Glazier' },
    { name: 'Dan Conlan', role: 'Glazier' },
    { name: 'Struan O\'Donnell', role: 'Glazier' },
    { name: 'Sam Chang', role: 'Glazier' }
  ]
};

export default function MeetingHistory() {
  const queryClient = useQueryClient();
  const [, setLocation] = useLocation();
  const [selectedMeetingsForExport, setSelectedMeetingsForExport] = useState<Set<string>>(new Set());
  const [showExportSelector, setShowExportSelector] = useState<boolean>(false);
  const [expandedItems, setExpandedItems] = useState<Set<string>>(new Set());
  const [expandedCategories, setExpandedCategories] = useState<Set<string>>(new Set());
  const [expandedActionEditors, setExpandedActionEditors] = useState<Set<string>>(new Set());
  const [expandedMeetings, setExpandedMeetings] = useState<Set<string>>(new Set());
  const [meetingAttendance, setMeetingAttendance] = useState<Record<string, string[]>>({});
  const [showAttendanceSelector, setShowAttendanceSelector] = useState<string>('');
  const [lockedAttendance, setLockedAttendance] = useState<Record<string, boolean>>({});
  const [meetingSignatures, setMeetingSignatures] = useState<Record<string, Record<string, { status: 'signed' | 'remote' | 'absent'; signatureData: string | null; signedAt: string }>>>({});
  const [showSignatureCarousel, setShowSignatureCarousel] = useState<string>('');
  const [selectedPeriod, setSelectedPeriod] = useState<string>('relevant');
  const [selectedMeetingForStats, setSelectedMeetingForStats] = useState<string>('all');
  const [isGeneratingTitles, setIsGeneratingTitles] = useState<boolean>(false);
  const [aiGenerationProgress, setAiGenerationProgress] = useState({ current: 0, total: 0, currentItem: '' });
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [isMergingMeeting, setIsMergingMeeting] = useState<boolean>(false);
  const [mergeProgress, setMergeProgress] = useState({ current: 0, total: 0, currentItem: '' });
  const [showMergeConfirmation, setShowMergeConfirmation] = useState(false);
  const [mergeConfirmationData, setMergeConfirmationData] = useState<{
    currentMeetingDate: string;
    nextMeetingDate: Date;
    itemCount: number;
    movableItemCount: number;
    formattedNextDate: string;
  } | null>(null);
  const [selectedDestinationDate, setSelectedDestinationDate] = useState<string>('');
  const [showRescheduleMeetingModal, setShowRescheduleMeetingModal] = useState(false);
  const [rescheduleMeetingSourceDate, setRescheduleMeetingSourceDate] = useState<string>('');
  const [rescheduleTargetDate, setRescheduleTargetDate] = useState<string>('');
  const [isRescheduling, setIsRescheduling] = useState(false);

  const [aiTitleStatus, setAiTitleStatus] = useState({ needsAI: false, emptyTitles: 0, totalItems: 0 });
  
  // New tab state management
  const [activeTab, setActiveTab] = useState<'meeting-history' | 'actions'>('meeting-history');
  const [showAddIdeaModal, setShowAddIdeaModal] = useState(false);
  const [selectedMeetingDate, setSelectedMeetingDate] = useState<string>('');
  const [showNotesModal, setShowNotesModal] = useState(false);
  const [editingMeetingNotes, setEditingMeetingNotes] = useState<string>('');
  const [editingNotesItem, setEditingNotesItem] = useState<MeetingItem | null>(null);
  const [showSaveDropdown, setShowSaveDropdown] = useState(false);
  const [isEnhancingNotes, setIsEnhancingNotes] = useState<boolean>(false);
  
  const [newIdeaForm, setNewIdeaForm] = useState({
    type: 'Business Ideas' as 'Business Ideas' | 'Safety Ideas',
    description: '',
    submittedBy: '',
    submittedById: '',
    assignedTo: '',
    ideaType: '',
    meetingDate: ''
  });
  const [ideaTypeOptions, setIdeaTypeOptions] = useState<string[]>([]);
  const [sharepointUsers, setSharepointUsers] = useState<Array<{id: number, title: string, email: string}>>([]);

  // Status update states
  const [isUpdatingStatus, setIsUpdatingStatus] = useState<string>(''); // item ID being updated

  // Status options from SharePoint (including 'Closed' for Business Ideas and Safety Ideas)
  const [statusOptions, setStatusOptions] = useState<string[]>(['Submitted', 'In Discussion', 'Actioned', 'Closed']);
  const [isCreatingItem, setIsCreatingItem] = useState(false);
  const [pendingItems, setPendingItems] = useState<Array<{
    id: string;
    type: string;
    description: string;
    submittedBy: string;
    meetingDate: string;
  }>>([]);
  
  // Note: Old dropdown suggestion states removed - now using InlineTextarea component

  // Floating add button and note feedback states
  const [showFloatingAdd, setShowFloatingAdd] = useState(false);
  const [recentlySavedNotes, setRecentlySavedNotes] = useState<Set<string>>(new Set());
  const [floatingMeetingDate, setFloatingMeetingDate] = useState<string>('');
  
  // Flip card state for action items
  const [flippedCards, setFlippedCards] = useState<Set<string>>(new Set());
  const [isUpdatingAction, setIsUpdatingAction] = useState<string>(''); // item ID being updated
  const [recentlySavedActions, setRecentlySavedActions] = useState<Set<string>>(new Set());
  const [isEnhancingActionNotes, setIsEnhancingActionNotes] = useState<string>(''); // item ID being enhanced
  
  // Local edit state for text fields to prevent save-on-every-keystroke
  const [localActionEdits, setLocalActionEdits] = useState<Record<string, { 
    actionNotes?: string; 
    actionAssignedTo?: string;
  }>>({});
  const debounceTimers = useRef<Record<string, NodeJS.Timeout>>({});
  
  // Toggle flip card
  const toggleFlipCard = (itemId: string) => {
    setFlippedCards(prev => {
      const newSet = new Set(prev);
      if (newSet.has(itemId)) {
        newSet.delete(itemId);
      } else {
        newSet.add(itemId);
      }
      return newSet;
    });
  };

  // Functions for SharePoint item management
  const createSharePointItem = async (itemData: any) => {
    // Add an optimistic pending item immediately so something appears on screen
    const pendingId = `pending-${Date.now()}`;
    setPendingItems(prev => [...prev, {
      id: pendingId,
      type: itemData.type,
      description: itemData.description,
      submittedBy: itemData.submittedBy,
      meetingDate: itemData.meetingDate
    }]);
    setIsCreatingItem(true);
    try {
      const response = await authenticatedFetch('/api/sharepoint/create-item', {
        method: 'POST',
        body: JSON.stringify({
          listType: itemData.type,
          itemData: {
            title: itemData.title,
            description: itemData.description,
            submittedBy: itemData.submittedBy,
            ideaType: itemData.ideaType,
            status: 'Submitted'
          }
        })
      });

      const result = await response.json();
      if (result.success) {
        // Refresh data without page reload
        await queryClient.invalidateQueries({ queryKey: ['/api/meeting-history'] });
        showSuccess('Item Added', `Your ${itemData.type === 'Safety Ideas' ? 'safety idea' : 'business idea'} has been saved to SharePoint.`);
      } else {
        showError('Creation Failed', result.error || 'Failed to create item');
      }
    } catch (error) {
      console.error('Error creating item:', error);
      showError('Creation Failed', 'Failed to create item');
    } finally {
      setIsCreatingItem(false);
      // Remove the pending placeholder once we have the real result
      setPendingItems(prev => prev.filter(p => p.id !== pendingId));
    }
  };

  const updateItemStatus = async (item: MeetingItem, newStatus: string) => {
    setIsUpdatingStatus(item.id);
    try {
      const response = await authenticatedFetch('/api/sharepoint/update-item', {
        method: 'POST',
        body: JSON.stringify({
          itemId: item.id,
          listType: item.type,
          updates: { status: newStatus }
        })
      });

      const result = await response.json();
      if (result.success) {
        showSuccess('Success', `Status updated to ${newStatus}`);
        // Refresh data without page reload
        await queryClient.invalidateQueries({ queryKey: ['/api/meeting-history'] });
      } else {
        showError('Update Failed', result.error || 'Failed to update status');
      }
    } catch (error) {
      console.error('Error updating status:', error);
      showError('Update Failed', 'Failed to update status');
    } finally {
      setIsUpdatingStatus('');
    }
  };

  const updateMeetingNotes = async (item: MeetingItem, notes: string) => {
    try {
      // Save locally first (instant, reliable backup)
      const localSave = await fetch('/api/action-items', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          listType: item.type === 'Near Miss' ? 'NearMiss' : item.type.replace(' ', ''),
          sharePointItemId: item.id,
          meetingNotes: notes
        })
      });
      const localResult = await localSave.json();
      if (!localResult.success) {
        console.warn('Local meeting notes save failed:', localResult.error);
      }

      // Also save to SharePoint (background sync)
      try {
        const response = await authenticatedFetch('/api/sharepoint/update-item', {
          method: 'POST',
          body: JSON.stringify({
            itemId: item.id,
            listType: item.type,
            updates: { meetingNotes: notes }
          })
        });
        const result = await response.json();
        if (!result.success) {
          console.warn('SharePoint meeting notes sync failed:', result.error);
        }
      } catch (spError) {
        console.warn('SharePoint sync error (notes saved locally):', spError);
      }

      showSuccess('Success', 'Meeting notes saved successfully!');
      // Refresh data
      await queryClient.invalidateQueries({ queryKey: ['/api/meeting-history'] });
    } catch (error) {
      console.error('Error updating meeting notes:', error);
      showError('Update Failed', 'Failed to update meeting notes');
    }
  };

  // Helper to get current value for text fields (local edit or saved value)
  const getActionFieldValue = (item: MeetingItem, field: 'actionNotes' | 'actionAssignedTo') => {
    const localEdit = localActionEdits[item.id];
    if (localEdit && localEdit[field] !== undefined) {
      return localEdit[field];
    }
    return item[field] || '';
  };
  
  // Handle text field changes with debouncing
  const handleActionTextChange = (item: MeetingItem, field: 'actionNotes' | 'actionAssignedTo', value: string) => {
    // Update local state immediately for responsive typing
    setLocalActionEdits(prev => ({
      ...prev,
      [item.id]: { ...prev[item.id], [field]: value }
    }));
    
    // Clear existing timer for this item
    const timerKey = `${item.id}-${field}`;
    if (debounceTimers.current[timerKey]) {
      clearTimeout(debounceTimers.current[timerKey]);
    }
    
    // Set new debounce timer (1 second delay)
    debounceTimers.current[timerKey] = setTimeout(async () => {
      await updateActionFields(item, { [field]: value }, item.id, field);
    }, 1000);
  };

  // Update action fields for an item - stored in local database (not SharePoint)
  const updateActionFields = async (
    item: MeetingItem, 
    actionUpdates: {
      actionPriority?: string;
      actionStatus?: string;
      actionAssignedTo?: string;
      actionStartDate?: string;
      actionDueDate?: string;
      actionNotes?: string;
    },
    clearLocalEditItemId?: string,
    clearLocalEditField?: 'actionNotes' | 'actionAssignedTo'
  ) => {
    setIsUpdatingAction(item.id);
    try {
      // Get current action data - prefer local edits over item values for text fields
      const currentActionData = {
        actionPriority: item.actionPriority || '',
        actionStatus: item.actionStatus || '',
        actionAssignedTo: localActionEdits[item.id]?.actionAssignedTo ?? item.actionAssignedTo ?? '',
        actionStartDate: item.actionStartDate || '',
        actionDueDate: item.actionDueDate || '',
        actionNotes: localActionEdits[item.id]?.actionNotes ?? item.actionNotes ?? ''
      };
      
      const mergedData = { ...currentActionData, ...actionUpdates };
      
      // Save to local database (not SharePoint)
      const response = await fetch('/api/action-items', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          listType: item.type === 'Near Miss' ? 'NearMiss' : item.type.replace(' ', ''),
          sharePointItemId: item.id,
          ...mergedData
        })
      });

      const result = await response.json();
      if (result.success) {
        // Update cache first, then clear local edits
        queryClient.setQueryData(['/api/meeting-history'], (oldData: any) => {
          if (!oldData?.data) return oldData;
          return {
            ...oldData,
            data: oldData.data.map((i: MeetingItem) => 
              i.id === item.id ? { ...i, ...mergedData } : i
            )
          };
        });
        
        // Clear local edit after cache is updated (only if specified)
        if (clearLocalEditItemId && clearLocalEditField) {
          setLocalActionEdits(prev => {
            const newEdits = { ...prev };
            if (newEdits[clearLocalEditItemId]) {
              delete newEdits[clearLocalEditItemId][clearLocalEditField];
              if (Object.keys(newEdits[clearLocalEditItemId]).length === 0) {
                delete newEdits[clearLocalEditItemId];
              }
            }
            return newEdits;
          });
        }
        
        // Show "Saved" indicator for 2 seconds
        setRecentlySavedActions(prev => new Set(prev).add(item.id));
        setTimeout(() => {
          setRecentlySavedActions(prev => {
            const newSet = new Set(prev);
            newSet.delete(item.id);
            return newSet;
          });
        }, 2000);
      } else {
        showError('Update Failed', result.error || 'Failed to update action details');
      }
    } catch (error) {
      console.error('Error updating action fields:', error);
      showError('Update Failed', 'Failed to update action details');
    } finally {
      setIsUpdatingAction('');
    }
  };

  const openAddModal = (meetingDate: string) => {
    setSelectedMeetingDate(meetingDate);
    setNewIdeaForm({
      type: 'Business Ideas',
      description: '',
      submittedBy: '',
      submittedById: '',
      assignedTo: '',
      ideaType: '',
      meetingDate: meetingDate
    });
    setShowAddIdeaModal(true);
    fetchSharePointFormData();
  };

  const openNotesModal = (item: MeetingItem) => {
    setEditingNotesItem(item);
    setEditingMeetingNotes(item.meetingNotes || '');
    setShowNotesModal(true);
  };
  
  // Error and success modal states
  const [showErrorModal, setShowErrorModal] = useState(false);
  const [showSuccessModal, setShowSuccessModal] = useState(false);
  const [showShareableModal, setShowShareableModal] = useState(false);
  const [modalMessage, setModalMessage] = useState('');
  const [modalTitle, setModalTitle] = useState('');
  const [shareableUrl, setShareableUrl] = useState('');
  const [shareableFilename, setShareableFilename] = useState('');

  // Helper functions for modal notifications
  const showError = (title: string, message: string) => {
    setModalTitle(title);
    setModalMessage(message);
    setShowErrorModal(true);
  };

  const showSuccess = (title: string, message: string) => {
    setModalTitle(title);
    setModalMessage(message);
    setShowSuccessModal(true);
    // Auto-dismiss after 3 seconds
    setTimeout(() => setShowSuccessModal(false), 3000);
  };

  const showShareableUrlSuccess = (url: string, filename: string) => {
    setShareableUrl(url);
    setShareableFilename(filename);
    setShowShareableModal(true);
  };

  // Fetch SharePoint data for form
  const fetchSharePointFormData = async () => {
    try {
      const token = await authService.getSharePointToken();
      
      // Fetch idea type options based on selected list type
      const ideaTypeResponse = await fetch(`/api/sharepoint/choice-options/${newIdeaForm.type}/Idea Type`, {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      
      if (ideaTypeResponse.ok) {
        const ideaTypeData = await ideaTypeResponse.json();
        if (ideaTypeData.success) {
          setIdeaTypeOptions(ideaTypeData.choices || []);
        }
      }
      
      // Fetch SharePoint users (only once as it doesn't depend on form type)
      if (sharepointUsers.length === 0) {
        const usersResponse = await fetch('/api/sharepoint/users', {
          headers: { 'Authorization': `Bearer ${token}` }
        });
        
        if (usersResponse.ok) {
          const usersData = await usersResponse.json();
          if (usersData.success) {
            setSharepointUsers(usersData.users || []);
          }
        }
      }
    } catch (error) {
      console.error('Error fetching SharePoint form data:', error);
    }
  };

  // Refetch idea type options when form type changes
  const handleFormTypeChange = async (newType: 'Business Ideas' | 'Safety Ideas') => {
    setNewIdeaForm(prev => ({ 
      ...prev, 
      type: newType,
      ideaType: '' // Reset idea type when changing list type
    }));
    
    // Fetch new idea type options for this list type
    try {
      const token = await authService.getSharePointToken();
      const ideaTypeResponse = await fetch(`/api/sharepoint/choice-options/${newType}/Idea Type`, {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      
      if (ideaTypeResponse.ok) {
        const ideaTypeData = await ideaTypeResponse.json();
        if (ideaTypeData.success) {
          setIdeaTypeOptions(ideaTypeData.choices || []);
        }
      }
    } catch (error) {
      console.error('Error fetching idea type options:', error);
    }
  };



  // Enhanced update meeting notes with visual feedback
  const updateMeetingNotesWithFeedback = async (item: MeetingItem, notes: string) => {
    try {
      // Show saving state
      setRecentlySavedNotes(prev => {
        const newSet = new Set(prev);
        newSet.delete(item.id); // Remove any existing state
        return newSet;
      });

      // Call the actual update function
      await updateMeetingNotes(item, notes);
      
      // Show success feedback
      setRecentlySavedNotes(prev => {
        const newSet = new Set(prev);
        newSet.add(item.id);
        return newSet;
      });
      
      // Remove feedback after 3 seconds
      setTimeout(() => {
        setRecentlySavedNotes(prev => {
          const newSet = new Set(prev);
          newSet.delete(item.id);
          return newSet;
        });
      }, 3000);
      
    } catch (error) {
      console.error('Error updating meeting notes:', error);
      showError('Update Failed', 'Failed to update meeting notes');
    }
  };

  // Handle adding new ideas
  const handleAddIdea = async () => {
    try {
      if (!newIdeaForm.description.trim() || !newIdeaForm.submittedBy.trim()) {
        showError('Validation Error', 'Please fill in description and select a person');
        return;
      }

      const token = await authService.getSharePointToken();
      const response = await fetch('/api/sharepoint/create-item', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          listType: newIdeaForm.type,
          itemData: {
            title: '', // Will be generated by AI or left empty
            description: newIdeaForm.description,
            submittedBy: newIdeaForm.submittedBy || 'Current User',
            ideaType: newIdeaForm.ideaType,
            status: 'Submitted',
            meetingDate: newIdeaForm.meetingDate
          }
        }),
      });

      if (!response.ok) {
        throw new Error(`Failed to add idea: ${response.status}`);
      }

      const result = await response.json();
      if (result.success) {
        showSuccess('Idea Added', 'Your idea has been successfully submitted to SharePoint');
        setShowAddIdeaModal(false);
        setNewIdeaForm({
          type: 'Business Ideas',
          description: '',
          submittedBy: '',
          submittedById: '',
          assignedTo: '',
          ideaType: '',
          meetingDate: ''
        });
        // Refresh the data without page reload
        await queryClient.invalidateQueries({ queryKey: ['/api/meeting-history'] });
      } else {
        showError('Submission Failed', result.error || 'Failed to add idea');
      }
    } catch (error) {
      console.error('Error adding idea:', error);
      showError('Error', 'Failed to submit idea. Please try again.');
    }
  };

  // URL state management for search persistence
  useEffect(() => {
    const urlParams = new URLSearchParams(window.location.search);
    const savedSearch = urlParams.get('search');
    if (savedSearch) {
      setSearchQuery(savedSearch);
    }
  }, []);

  // Update URL when search changes
  useEffect(() => {
    const url = new URL(window.location.href);
    if (searchQuery) {
      url.searchParams.set('search', searchQuery);
    } else {
      url.searchParams.delete('search');
    }
    window.history.replaceState({}, '', url);
  }, [searchQuery]);


  // Create authenticated fetch function
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

  // Fetch meeting history from backend API
  const { data: apiResponse, isLoading } = useQuery({
    queryKey: ['/api/meeting-history'],
    queryFn: () => authenticatedFetch('/api/meeting-history').then(res => res.json()),
  });

  // Meeting attendance query
  const { data: attendanceResponse } = useQuery({
    queryKey: ['/api/meeting-attendance'],
    queryFn: () => fetch('/api/meeting-attendance').then(res => res.json()),
  });

  // Meeting signatures query
  const { data: signaturesResponse } = useQuery({
    queryKey: ['/api/meeting-signatures'],
    queryFn: () => fetch('/api/meeting-signatures').then(res => res.json()),
  });

  // Meeting locks will be fetched individually when needed
  // For now, we'll use local state combined with mutations to the API

  // Update meeting signatures from API response
  useEffect(() => {
    if (signaturesResponse?.success && signaturesResponse.signatures) {
      setMeetingSignatures(signaturesResponse.signatures);
    }
  }, [signaturesResponse]);

  // Update meeting attendance from API response
  useEffect(() => {
    if (attendanceResponse?.success && attendanceResponse.attendance) {
      setMeetingAttendance(attendanceResponse.attendance);
    }
  }, [attendanceResponse]);

  // Normalise any date to YYYY-MM-DD for consistent lock key matching
  const normaliseLockDate = (raw: string): string => {
    try { return new Date(raw).toISOString().split('T')[0]; } catch { return raw; }
  };

  // Single query for ALL meeting locks — avoids N per-meeting fetch calls and race conditions
  const { data: locksResponse } = useQuery({
    queryKey: ['/api/meeting-locks'],
    queryFn: () => fetch('/api/meeting-locks').then(res => res.json()),
    refetchOnWindowFocus: true,
    staleTime: 0,
  });

  // Sync locks from query into state — keyed by normalised YYYY-MM-DD
  useEffect(() => {
    if (locksResponse?.success && locksResponse.locks) {
      setLockedAttendance(locksResponse.locks);
    }
  }, [locksResponse]);

  // Mutations for updating meeting locks and attendance
  const updateMeetingLockMutation = useMutation({
    mutationFn: async ({ meetingDate, isLocked }: { meetingDate: string, isLocked: boolean }) => {
      const normDate = normaliseLockDate(meetingDate);
      const response = await fetch('/api/meeting-locks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          meetingDate: normDate,
          isLocked,
          lockedBy: 'current-user'
        })
      });
      return response.json();
    },
    onSuccess: (data, variables) => {
      if (data.success) {
        const normDate = normaliseLockDate(variables.meetingDate);
        setLockedAttendance(prev => ({ ...prev, [normDate]: variables.isLocked }));
        queryClient.invalidateQueries({ queryKey: ['/api/meeting-locks'] });
      }
    },
    onError: (error) => {
      console.error('Error updating meeting lock:', error);
      showError('Lock Update Failed', 'Failed to update meeting lock status');
    }
  });

  const updateMeetingAttendanceMutation = useMutation({
    mutationFn: async ({ meetingDate, attendeeName, isPresent }: { meetingDate: string, attendeeName: string, isPresent: boolean }) => {
      const response = await fetch('/api/meeting-attendance', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          meetingDate,
          attendeeName,
          isPresent
        })
      });
      return response.json();
    },
    onSuccess: (data, variables) => {
      if (data.success) {
        // Update local state immediately for responsive UI
        setMeetingAttendance(prev => {
          const current = prev[variables.meetingDate] || [];
          if (variables.isPresent) {
            // Add attendee if not already present
            if (!current.includes(variables.attendeeName)) {
              return {
                ...prev,
                [variables.meetingDate]: [...current, variables.attendeeName]
              };
            }
          } else {
            // Remove attendee
            return {
              ...prev,
              [variables.meetingDate]: current.filter(name => name !== variables.attendeeName)
            };
          }
          return prev;
        });
        // Invalidate queries to refresh data
        queryClient.invalidateQueries({ queryKey: ['/api/meeting-attendance'] });
      }
    },
    onError: (error) => {
      console.error('Error updating meeting attendance:', error);
      showError('Attendance Update Failed', 'Failed to update meeting attendance');
    }
  });

  const meetingItems: MeetingItem[] = (apiResponse as any)?.data || [];
  const actionItems: ActionItem[] = (apiResponse as any)?.actions || [];
  const isConfigured = (apiResponse as any)?.configured || false;
  
  // Actions loaded successfully

  // Helper function to find Actions related to a specific item
  const getRelatedActions = (item: MeetingItem): ActionItem[] => {
    if (!actionItems || actionItems.length === 0) return [];
    
    // Actions link to original items via the Link field (Hyperlink type)
    return actionItems.filter(action => {
      // The Link field contains SharePoint URLs like:
      // .../Lists/Safety%20Ideas/DispForm.aspx?ID=123
      // .../Lists/Business%20Ideas/DispForm.aspx?ID=456
      if (action.sourceLink) {
        // Extract ID from the URL
        const idMatch = action.sourceLink.match(/ID=(\d+)/i);
        if (idMatch && idMatch[1] === item.id) {
          return true;
        }
        
        // Also check if the link contains the list type and ID
        const listNameMatch = 
          (item.type === 'Business Ideas' && action.sourceLink.includes('Business%20Ideas')) ||
          (item.type === 'Safety Ideas' && action.sourceLink.includes('Safety%20Ideas')) ||
          (item.type === 'Near Miss' && action.sourceLink.includes('Near%20Miss'));
        
        return listNameMatch && action.sourceLink.includes(`ID=${item.id}`);
      }
      
      // Fallback: match by category and details if no link is present
      if (action.category) {
        const categoryMatch = 
          (item.type === 'Business Ideas' && action.category.includes('Business')) ||
          (item.type === 'Safety Ideas' && action.category.includes('Safety')) ||
          (item.type === 'Near Miss' && (action.category.includes('Near Miss') || action.category.includes('Incident')));
        
        // Match submitter and timeframe
        const submitterMatch = action.submittedBy === item.submittedBy;
        const sameTimeframe = Math.abs(
          new Date(action.submittedDate).getTime() - new Date(item.submittedDate).getTime()
        ) < (7 * 24 * 60 * 60 * 1000); // Within 7 days
        
        return categoryMatch && submitterMatch && sameTimeframe;
      }
      
      return false;
    });
  };
  


  // Check AI title status when needed
  const checkAIStatus = async () => {
    // Authentication is handled at app level
    
    try {
      const token = await authService.getSharePointToken();
      const response = await fetch('/api/ai-title-status', {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });
      const result = await response.json();
      
      // The backend returns data directly, not wrapped in success
      if (response.ok) {
        setAiTitleStatus({
          needsAI: result.needsAI,
          emptyTitles: result.emptyTitles,
          totalItems: result.totalItems,
        });
      }
    } catch (error) {
      console.error('Error checking AI status:', error);
    }
  };

  // Check AI status on mount and when configured
  useEffect(() => {
    if (isConfigured) {
      checkAIStatus();
    }
  }, [isConfigured]);
  const statusMessage = (apiResponse as any)?.message || '';

  // Generate AI titles with enhanced progress tracking
  const generateAITitles = async () => {
    if (!isConfigured || isGeneratingTitles) return;
    
    setIsGeneratingTitles(true);
    
    try {
      // Get initial count of items needing AI titles
      await checkAIStatus();
      const totalItems = aiTitleStatus.emptyTitles;
      
      setAiGenerationProgress({ 
        current: 0, 
        total: totalItems, 
        currentItem: `Initializing OpenAI GPT-4o for ${totalItems} items...` 
      });
      
      // Simulate progress updates during processing
      const progressInterval = setInterval(() => {
        setAiGenerationProgress(prev => {
          if (prev.current < prev.total) {
            const newCurrent = prev.current + 1;
            return {
              ...prev,
              current: newCurrent,
              currentItem: `Processing item ${newCurrent} of ${prev.total} with AI...`
            };
          }
          return prev;
        });
      }, 800); // Update every 800ms for realistic timing
      
      const token = await authService.getSharePointToken();
      const response = await fetch('/api/run-ai-titles', {
        method: 'POST',
        headers: { 
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        }
      });
      
      const result = await response.json();
      
      clearInterval(progressInterval);
      
      if (result.success) {
        setAiGenerationProgress({ 
          current: totalItems, 
          total: totalItems, 
          currentItem: 'AI title generation completed successfully!' 
        });
        
        // Wait to show completion, then refresh without page reload
        setTimeout(async () => {
          await queryClient.invalidateQueries({ queryKey: ['/api/meeting-history'] });
          setIsGeneratingTitles(false);
        }, 1500);
      } else {
        throw new Error(result.error || 'AI generation failed');
      }
    } catch (error) {
      console.error('AI generation failed:', error);
      setAiGenerationProgress({ 
        current: 0, 
        total: 0, 
        currentItem: 'AI generation failed. Please try again.' 
      });
      setTimeout(() => setIsGeneratingTitles(false), 2000);
    }
  };

  // Mock data for demonstration - replace with actual SharePoint data
  const mockMeetingItems: MeetingItem[] = [
    {
      id: '1',
      title: 'Improve safety signage in warehouse',
      description: 'Current safety signs are faded and need updating',
      type: 'Safety Ideas',
      status: 'Closed',
      meetingDate: '2025-01-15',
      meetingNotes: 'Approved for immediate implementation. Estimated cost £500. Signs to be replaced within 2 weeks.',
      submittedBy: 'John Smith',
      submittedDate: '2025-01-08',
      assignedTo: 'Facilities Team'
    },
    {
      id: '2',
      title: 'Digital inventory tracking system',
      description: 'Replace manual inventory with barcode scanning system',
      type: 'Business Ideas',
      status: 'Closed',
      meetingDate: '2025-01-15',
      meetingNotes: 'Discussed implementation timeline. Requires budget approval. Pilot program approved for main warehouse.',
      submittedBy: 'Sarah Johnson',
      submittedDate: '2025-01-10',
      assignedTo: 'IT Department'
    },
    {
      id: '3',
      title: 'Slip hazard near loading dock',
      description: 'Water accumulation causing slip hazard during rain',
      type: 'Near Miss',
      status: 'Closed',
      meetingDate: '2025-01-01',
      meetingNotes: 'Immediate action required. Drainage issue identified. Temporary matting to be installed immediately.',
      submittedBy: 'Mike Wilson',
      submittedDate: '2024-12-28',
      assignedTo: 'Maintenance Team'
    }
  ];

  // Smart filtering logic for forward-looking planning
  const getFilteredData = () => {
    const allData: MeetingItem[] = (meetingItems && meetingItems.length > 0) ? meetingItems : mockMeetingItems;
    
    if (selectedPeriod === 'all') return allData;
    
    // Open Actions filter - items that are Actioned OR have action tracking data and aren't closed
    if (selectedPeriod === 'actions') {
      return allData.filter(item => {
        const isActioned = item.status === 'Actioned';
        const hasActionFields = !!(item.actionPriority || 
          item.actionStatus || 
          item.actionNotes || 
          item.actionAssignedTo ||
          item.actionDueDate);
        const isNotClosed = item.status?.toLowerCase() !== 'closed';
        return isActioned || (hasActionFields && isNotClosed);
      });
    }
    
    const now = new Date();
    const currentYear = now.getFullYear();
    
    // For quarter-based filtering, use intelligent logic
    if (selectedPeriod.startsWith('q')) {
      const quarter = selectedPeriod;
      
      // Get the date range for the selected quarter
      let quarterStart: Date;
      let quarterEnd: Date;
      
      switch (quarter) {
        case 'q1':
          quarterStart = new Date(currentYear, 0, 1); // January 1
          quarterEnd = new Date(currentYear, 2, 31); // March 31
          break;
        case 'q2':
          quarterStart = new Date(currentYear, 3, 1); // April 1
          quarterEnd = new Date(currentYear, 5, 30); // June 30
          break;
        case 'q3':
          quarterStart = new Date(currentYear, 6, 1); // July 1
          quarterEnd = new Date(currentYear, 8, 30); // September 30
          break;
        case 'q4':
          quarterStart = new Date(currentYear, 9, 1); // October 1
          quarterEnd = new Date(currentYear, 11, 31); // December 31
          break;
        default:
          return allData;
      }
      
      return allData.filter(item => {
        const itemDate = new Date(item.meetingDate);
        const isInQuarter = itemDate >= quarterStart && itemDate <= quarterEnd;
        
        // Smart filtering: Show items if they...
        // 1. Have meeting dates in the selected quarter
        // 2. Are active (not closed) and submitted before/during the quarter
        // 3. Are recent items that might need attention in this quarter
        
        if (isInQuarter) return true;
        
        // For current quarter or future quarters, also show active items
        if (quarterStart <= now || quarterEnd >= now) {
          const isActive = item.status !== 'Closed';
          const submittedDate = new Date(item.submittedDate);
          const isRecentlySubmitted = submittedDate >= quarterStart;
          
          return isActive && isRecentlySubmitted;
        }
        
        return false;
      });
    }
    
    // For "relevant" filter - show items with meeting dates within 3 months before/after today
    if (selectedPeriod === 'relevant') {
      const threeMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 3, 1);
      const threeMonthsAhead = new Date(now.getFullYear(), now.getMonth() + 3, 30);
      
      return allData.filter(item => {
        const itemDate = new Date(item.meetingDate);
        
        // Simple date-based filter: show items with meeting dates in the 6-month window
        return itemDate >= threeMonthsAgo && itemDate <= threeMonthsAhead;
      });
    }
    
    // Legacy "last6months" kept for backward compatibility
    if (selectedPeriod === 'last6months') {
      const sixMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 6, 1);
      return allData.filter(item => {
        const itemDate = new Date(item.meetingDate);
        return itemDate >= sixMonthsAgo && itemDate <= now;
      });
    }
    
    return allData;
  };
  
  const dataToUse = getFilteredData();
  
  // Search filter function - searches all text fields
  const searchItems = (items: MeetingItem[], query: string): MeetingItem[] => {
    if (!query.trim()) return items;
    
    const lowerQuery = query.toLowerCase();
    return items.filter(item => {
      // Search in all text fields
      const searchableFields = [
        item.title,
        item.description,
        item.meetingNotes,
        item.submittedBy,
        item.status,
        item.assignedTo,
        item.ideaType,
        item.type
      ];
      
      return searchableFields.some(field => 
        field && field.toLowerCase().includes(lowerQuery)
      );
    });
  };

  const filteredItems = searchQuery ? searchItems(dataToUse, searchQuery) : dataToUse;
  
  // Create unique meeting dates by grouping items using centralized date logic
  const meetingDateGroups = new Map<string, string>();
  filteredItems.forEach((item: MeetingItem) => {
    const dateKey = getDateGroupKey(item.meetingDate);
    if (dateKey === 'unknown-meeting') return;
    
    // Store the first ISO string we find for each date group
    if (!meetingDateGroups.has(dateKey)) {
      meetingDateGroups.set(dateKey, item.meetingDate);
    }
  });
  
  // Sort by actual date values
  const meetingDates: string[] = Array.from(meetingDateGroups.values())
    .sort((a, b) => new Date(b).getTime() - new Date(a).getTime());
  
  const toggleMeetingForExport = (meetingDate: string) => {
    setSelectedMeetingsForExport(prev => {
      const newSet = new Set(prev);
      if (newSet.has(meetingDate)) {
        newSet.delete(meetingDate);
      } else {
        // Limit to 3 meetings maximum
        if (newSet.size >= 3) {
          showError('Export Limit', 'Maximum 3 meetings can be exported at once to prevent browser overload');
          return prev;
        }
        newSet.add(meetingDate);
      }
      return newSet;
    });
  };
  
  const isMultipleMeetings = selectedMeetingsForExport.size > 1;
  
  // Helper functions for meeting status and styling - using centralized utilities

  const getMeetingHeaderStyle = (dateString: string, isFullyClosed: boolean = false) => {
    const status = getMeetingStatus(dateString);
    
    if (status.isUnknown) {
      return 'bg-gradient-to-r from-gray-600 to-gray-700 text-white'; // Gray for unknown
    } else {
      // Get current date in NZ timezone
      const now = new Date();
      const nzDate = new Date(now.toLocaleString("en-US", {timeZone: "Pacific/Auckland"}));
      
      // Parse meeting date and get the date part only
      const meetingDate = new Date(dateString);
      
      // Set both dates to start of day in NZ timezone for comparison
      const currentDateNZ = new Date(nzDate.getFullYear(), nzDate.getMonth(), nzDate.getDate());
      const meetingDateOnly = new Date(meetingDate.getFullYear(), meetingDate.getMonth(), meetingDate.getDate());
      
      if (currentDateNZ <= meetingDateOnly) {
        return 'bg-gradient-to-r from-green-600 to-green-700 text-white'; // Green = upcoming
      } else if (isFullyClosed) {
        return 'bg-gradient-to-r from-gray-500 to-gray-600 text-white'; // Gray = past & all items closed
      } else {
        return 'bg-gradient-to-r from-purple-600 to-purple-700 text-white'; // Purple = past with open items
      }
    }
  };

  const getMeetingTitle = (dateString: string) => {
    const status = getMeetingStatus(dateString);
    
    if (status.isUnknown) {
      return 'Unscheduled Items';
    } else {
      return 'Meeting';
    }
  };

  const getStatusSummary = (categorized: Record<string, MeetingItem[]>) => {
    const allItems = Object.values(categorized).flat();
    const statusCounts: Record<string, number> = {};
    
    allItems.forEach(item => {
      const status = item.status || 'Unknown';
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    });
    
    // Create a readable summary of statuses
    const statusParts = Object.entries(statusCounts).map(([status, count]) => {
      return `${count} ${status}`;
    });
    
    return statusParts.join(', ');
  };

  // Highlight search terms in text
  const highlightSearchTerm = (text: string, searchTerm: string): JSX.Element => {
    if (!searchTerm || !text) return <span>{text}</span>;
    
    const regex = new RegExp(`(${searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
    const parts = text.split(regex);
    
    return (
      <span>
        {parts.map((part, index) => 
          regex.test(part) ? (
            <mark key={index} className="bg-yellow-200 text-gray-900 px-1 rounded">
              {part}
            </mark>
          ) : (
            <span key={index}>{part}</span>
          )
        )}
      </span>
    );
  };

  // Helper function to find the most recent past meeting (last purple meeting)
  const getMostRecentPastMeeting = (groupedData: Array<{meetingDate: string; categorized: any; totalItems: number}>) => {
    if (!groupedData || groupedData.length === 0) {
      return null;
    }

    // Filter past meetings and sort by date (most recent first)
    const pastMeetings = groupedData
      .filter(({ meetingDate }) => getMeetingStatus(meetingDate).isPast)
      .sort((a, b) => {
        const dateA = parseSharePointDate(a.meetingDate);
        const dateB = parseSharePointDate(b.meetingDate);
        if (!dateA || !dateB) return 0;
        return dateB.getTime() - dateA.getTime(); // Most recent first
      });

    return pastMeetings.length > 0 ? pastMeetings[0].meetingDate : null;
  };

  // Helper function to check if a meeting has movable items (Submitted status only)
  const getMeetingMovableStatus = (categorized: Record<string, MeetingItem[]>) => {
    const allItems = Object.values(categorized).flat();
    const movableItems = allItems.filter(item => item.status === 'Submitted');
    const hasMovableItems = movableItems.length > 0;
    const allItemsClosed = allItems.length > 0 && movableItems.length === 0;
    return {
      hasMovableItems,
      movableCount: movableItems.length,
      totalCount: allItems.length,
      allItemsClosed
    };
  };

  const groupedByMeeting = filteredItems.reduce((acc: Record<string, MeetingItem[]>, item: MeetingItem) => {
    // Use centralized date grouping logic
    const groupKey = getDateGroupKey(item.meetingDate);
    
    if (!acc[groupKey]) {
      acc[groupKey] = [];
    }
    acc[groupKey].push(item);
    return acc;
  }, {} as Record<string, MeetingItem[]>);

  // Group items by category within each meeting and sort by submission date (newest first)
  const groupedByMeetingAndCategory = Object.entries(groupedByMeeting)
    .map(([meetingDate, items]) => {
      const itemsTyped = items as MeetingItem[];
      
      // Sort items within each category by status priority, then by submission date (newest first)
      const sortBySubmissionDate = (items: MeetingItem[]) => 
        items.sort((a, b) => {
          // Priority order: Submitted first, Actioned second, Closed last
          const statusPriority = {
            'Submitted': 1,
            'In Discussion': 2,
            'Actioned': 3,
            'Closed': 4
          };
          
          const priorityA = statusPriority[a.status as keyof typeof statusPriority] || 5;
          const priorityB = statusPriority[b.status as keyof typeof statusPriority] || 5;
          
          // Sort by status priority first
          if (priorityA !== priorityB) {
            return priorityA - priorityB;
          }
          
          // If same status, sort by submission date (newest first)
          const dateA = new Date(a.submittedDate || a.id);
          const dateB = new Date(b.submittedDate || b.id);
          
          // Debug logging for timezone issues
          if (window.location.search.includes('debug')) {
          }
          
          return dateB.getTime() - dateA.getTime(); // Newest first
        });
      
      const categorized = {
        'Safety Ideas': sortBySubmissionDate(itemsTyped.filter((item: MeetingItem) => item.type === 'Safety Ideas')),
        'Business Ideas': sortBySubmissionDate(itemsTyped.filter((item: MeetingItem) => item.type === 'Business Ideas')),
        'Near Miss': sortBySubmissionDate(itemsTyped.filter((item: MeetingItem) => item.type === 'Near Miss')),
      };
      return { meetingDate, categorized, totalItems: itemsTyped.length };
    })
    // Filter out meetings with no matching items when searching
    .filter(({ categorized, totalItems }) => {
      if (!searchQuery) return totalItems > 0; // Show all meetings with items when not searching
      // When searching, only show meetings that have matching items
      return Object.values(categorized).some(categoryItems => categoryItems.length > 0);
    })
    .sort((a, b) => {
      // Sort by priority: Unknown first, then Upcoming, then Past meetings (newest first)
      const statusA = getMeetingStatus(a.meetingDate);
      const statusB = getMeetingStatus(b.meetingDate);
      
      // Unknown meetings first
      if (statusA.isUnknown && !statusB.isUnknown) return -1;
      if (!statusA.isUnknown && statusB.isUnknown) return 1;
      
      // If both unknown, sort alphabetically
      if (statusA.isUnknown && statusB.isUnknown) return 0;
      
      // Upcoming meetings second (chronological order - earliest first)
      if (statusA.isUpcoming && statusB.isPast) return -1;
      if (statusA.isPast && statusB.isUpcoming) return 1;
      
      // Within same status type, sort by date
      const dateA = new Date(a.meetingDate);
      const dateB = new Date(b.meetingDate);
      
      if (statusA.isUpcoming && statusB.isUpcoming) {
        return dateB.getTime() - dateA.getTime(); // Newest upcoming first (most recent dates at top)
      } else {
        return dateB.getTime() - dateA.getTime(); // Latest past first
      }
    });



  const toggleItemExpansion = (itemId: string) => {
    const newExpanded = new Set(expandedItems);
    if (newExpanded.has(itemId)) {
      newExpanded.delete(itemId);
    } else {
      newExpanded.add(itemId);
    }
    setExpandedItems(newExpanded);
  };

  // Smart meeting expansion - auto-collapse others when one opens
  const toggleMeetingExpansion = (meetingDate: string) => {
    const newExpanded = new Set(expandedMeetings);
    const isCurrentlyExpanded = newExpanded.has(meetingDate);
    
    if (isCurrentlyExpanded) {
      // Simply collapse this meeting
      newExpanded.delete(meetingDate);
      // Also collapse all categories when collapsing a meeting
      setExpandedCategories(new Set());
      setExpandedItems(new Set());
    } else {
      // Smart behavior: collapse all others and expand this one
      newExpanded.clear();
      newExpanded.add(meetingDate);
      // Clear all categories when switching to a different meeting
      setExpandedCategories(new Set());
      setExpandedItems(new Set());
    }
    setExpandedMeetings(newExpanded);
  };

  // Smart category expansion within a meeting
  const toggleCategoryExpansion = (categoryKey: string, meetingDate: string) => {
    const newExpanded = new Set(expandedCategories);
    const isCurrentlyExpanded = newExpanded.has(categoryKey);
    
    if (isCurrentlyExpanded) {
      newExpanded.delete(categoryKey);
    } else {
      // Smart behavior: within this meeting, collapse other categories
      const meetingCategories = Array.from(expandedCategories).filter(key => 
        key.startsWith(meetingDate)
      );
      meetingCategories.forEach(key => newExpanded.delete(key));
      newExpanded.add(categoryKey);
    }
    setExpandedCategories(newExpanded);
  };

  // Master controls for all meetings
  const expandAllMeetings = () => {
    const allMeetingDates = groupedByMeetingAndCategory.map(({ meetingDate }) => meetingDate);
    setExpandedMeetings(new Set(allMeetingDates));
  };

  const collapseAllMeetings = () => {
    setExpandedMeetings(new Set());
    setExpandedCategories(new Set());
    setExpandedItems(new Set());
  };

  // Scroll detection for floating add button (after data processing)
  useEffect(() => {
    const handleScroll = () => {
      const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
      
      // Show floating button when scrolled down more than 200px
      const shouldShow = scrollTop > 200;
      setShowFloatingAdd(shouldShow);
      
      // Find the nearest upcoming meeting when showing floating button
      if (shouldShow && groupedByMeetingAndCategory) {
        const upcomingMeetings = groupedByMeetingAndCategory.filter(({ meetingDate }) => {
          const meetingStatus = getMeetingStatus(meetingDate);
          return meetingStatus.isUpcoming;
        });
        
        if (upcomingMeetings.length > 0) {
          setFloatingMeetingDate(upcomingMeetings[0].meetingDate);
        }
      }
    };

    window.addEventListener('scroll', handleScroll, { passive: true });
    return () => window.removeEventListener('scroll', handleScroll);
  }, [groupedByMeetingAndCategory]);

  // Master controls for categories within expanded meetings
  const expandAllCategories = () => {
    const allCategoryKeys: string[] = [];
    groupedByMeetingAndCategory.forEach(({ meetingDate, categorized }) => {
      if (expandedMeetings.has(meetingDate)) {
        (['Safety Ideas', 'Business Ideas', 'Near Miss'] as const).forEach(category => {
          if (categorized[category].length > 0) {
            allCategoryKeys.push(`${meetingDate}-${category}`);
          }
        });
      }
    });
    setExpandedCategories(new Set(allCategoryKeys));
  };

  const collapseAllCategories = () => {
    setExpandedCategories(new Set());
    setExpandedItems(new Set());
  };

  // Helper function to collapse a specific category
  const collapseCategory = (categoryKey: string) => {
    const newExpanded = new Set(expandedCategories);
    newExpanded.delete(categoryKey);
    setExpandedCategories(newExpanded);
  };

  const getTypeColor = (type: string) => {
    switch (type) {
      case 'Safety Ideas': return 'bg-red-100 text-red-800 border-red-200';
      case 'Business Ideas': return 'bg-blue-100 text-blue-800 border-blue-200';
      case 'Near Miss': return 'bg-orange-100 text-orange-800 border-orange-200';
      case 'Actions': return 'bg-purple-100 text-purple-800 border-purple-200';
      default: return 'bg-gray-100 text-gray-800 border-gray-200';
    }
  };

  const getCategoryIcon = (type: string) => {
    switch (type) {
      case 'Safety Ideas': return Shield;
      case 'Business Ideas': return Lightbulb;
      case 'Near Miss': return AlertTriangle;
      case 'Actions': return Users;
      default: return FileText;
    }
  };

  const getCategoryColor = (type: string) => {
    switch (type) {
      case 'Safety Ideas': return 'text-red-600 bg-red-50 border-red-200';
      case 'Business Ideas': return 'text-blue-600 bg-blue-50 border-blue-200';
      case 'Near Miss': return 'text-orange-600 bg-orange-50 border-orange-200';
      case 'Actions': return 'text-purple-600 bg-purple-50 border-purple-200';
      default: return 'text-gray-600 bg-gray-50 border-gray-200';
    }
  };

  const getStatusBadgeColor = (status: string, type: string) => {
    // SharePoint choice field colors - exact color matching
    switch (status) {
      case 'Closed':
        return 'bg-emerald-100 text-emerald-800 border-emerald-300 hover:bg-emerald-200'; // Mint green
      case 'Actions':
        return 'bg-sky-100 text-sky-800 border-sky-300 hover:bg-sky-200'; // Light blue
      case 'In Discussion':
        return 'bg-yellow-100 text-yellow-800 border-yellow-300 hover:bg-yellow-200'; // Gold
      case 'Submitted':
        return 'bg-blue-100 text-blue-800 border-blue-300 hover:bg-blue-200'; // Blue
      case 'Open':
      case 'In Progress':
        return 'bg-orange-100 text-orange-800 border-orange-300 hover:bg-orange-200'; // Orange for open items
      default:
        return 'bg-gray-100 text-gray-800 border-gray-300 hover:bg-gray-200'; // Default gray
    }
  };

  const formatDate = (dateString: string) => {
    return formatDisplayDate(dateString, 'meeting');
  };

  const toggleAttendee = (meetingDate: string, attendeeName: string) => {
    // Don't allow changes if attendance is locked
    if (lockedAttendance[normaliseLockDate(meetingDate)]) {
      return;
    }
    
    const currentAttendees = meetingAttendance[meetingDate] || [];
    const isCurrentlyAttending = currentAttendees.includes(attendeeName);
    
    // Update database using mutation
    updateMeetingAttendanceMutation.mutate({
      meetingDate,
      attendeeName,
      isPresent: !isCurrentlyAttending
    });
  };

  const saveSignatureMutation = useMutation({
    mutationFn: async ({ meetingDate, attendeeName, status, signatureData, signedAt }: {
      meetingDate: string; attendeeName: string; status: string; signatureData: string | null; signedAt: string;
    }) => {
      const response = await fetch('/api/meeting-signatures', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ meetingDate, attendeeName, status, signatureData, signedAt })
      });
      const data = await response.json();
      if (!response.ok || data.success === false) {
        throw new Error(data.error ?? `Server error ${response.status}`);
      }
      return data;
    },
    onSuccess: (data, variables) => {
      if (data.success) {
        setMeetingSignatures(prev => ({
          ...prev,
          [variables.meetingDate]: {
            ...(prev[variables.meetingDate] ?? {}),
            [variables.attendeeName]: {
              status: variables.status as 'signed' | 'remote' | 'absent',
              signatureData: variables.signatureData,
              signedAt: variables.signedAt
            }
          }
        }));
        queryClient.invalidateQueries({ queryKey: ['/api/meeting-signatures'] });
      }
    }
  });

  const toggleAttendanceLock = (meetingDate: string) => {
    const isCurrentlyLocked = lockedAttendance[normaliseLockDate(meetingDate)] || false;
    
    // Update database using mutation
    updateMeetingLockMutation.mutate({
      meetingDate,
      isLocked: !isCurrentlyLocked
    });
  };


  const isAttending = (meetingDate: string, attendeeName: string) => {
    // If no attendance data exists for this meeting, default to everyone present
    if (!meetingAttendance[meetingDate]) {
      return true;
    }
    return meetingAttendance[meetingDate]?.includes(attendeeName) || false;
  };

  // Initialize attendance with everyone present when first opening
  const initializeAttendance = (meetingDate: string) => {
    if (!meetingAttendance[meetingDate]) {
      const allAttendees = [
        ...meetingAttendees.management.map(a => a.name),
        ...meetingAttendees.glaziers.map(a => a.name)
      ];
      setMeetingAttendance(prev => ({
        ...prev,
        [meetingDate]: allAttendees
      }));
    }
  };

  const exportToCSV = async () => {
    try {
      // Show loading state
      const button = document.querySelector('.export-csv-btn') as HTMLButtonElement;
      if (button) {
        button.disabled = true;
        if (isMultipleMeetings) {
          button.innerHTML = '<i class="fas fa-spinner fa-spin text-xs"></i> <span>Generating ' + selectedMeetingsForExport.size + ' CSV Files...</span>';
        } else {
          button.innerHTML = '<i class="fas fa-spinner fa-spin text-xs"></i> <span>Generating CSV...</span>';
        }
      }

      if (selectedMeetingsForExport.size === 0) {
        showError('Export Required', 'Please select at least one meeting to export');
        return;
      }

      const meetingsToExport = Array.from(selectedMeetingsForExport);
      
      if (isMultipleMeetings) {
        // Generate separate CSV file for each selected meeting
        for (let i = 0; i < meetingsToExport.length; i++) {
          const meetingDate = meetingsToExport[i];
          
          // Update button to show current progress
          if (button) {
            button.innerHTML = `<i class="fas fa-spinner fa-spin text-xs"></i> <span>CSV ${i + 1} of ${meetingsToExport.length}...</span>`;
          }
          
          const response = await fetch('/api/generate-meeting-csv', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({
              meetingData: dataToUse,
              selectedMeeting: meetingDate,
              selectedType: 'all',
            }),
          });

          if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
          }

          // Download the CSV file
          const blob = await response.blob();
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.style.display = 'none';
          a.href = url;
          
          // Get filename from response headers or generate one
          const contentDisposition = response.headers.get('content-disposition');
          let filename = 'meeting-data.csv';
          if (contentDisposition) {
            const filenameMatch = contentDisposition.match(/filename="(.+)"/);
            if (filenameMatch) {
              filename = filenameMatch[1];
            }
          }
          
          a.download = filename;
          document.body.appendChild(a);
          a.click();
          window.URL.revokeObjectURL(url);
          document.body.removeChild(a);
          
          // Small delay between downloads to prevent browser blocking
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      } else {
        // Single CSV file generation
        const meetingDate = meetingsToExport[0];
        
        const response = await fetch('/api/generate-meeting-csv', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            meetingData: dataToUse,
            selectedMeeting: meetingDate,
            selectedType: 'all',
          }),
        });

        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }

        // Download the CSV file
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        
        // Get filename from response headers or generate one
        const contentDisposition = response.headers.get('content-disposition');
        let filename = 'meeting-data.csv';
        if (contentDisposition) {
          const filenameMatch = contentDisposition.match(/filename="(.+)"/);
          if (filenameMatch) {
            filename = filenameMatch[1];
          }
        }
        
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
      }

    } catch (error) {
      console.error('Error generating CSV:', error);
      showError('Export Failed', 'Failed to generate CSV. Please try again.');
    } finally {
      // Reset button state
      const button = document.querySelector('.export-csv-btn') as HTMLButtonElement;
      if (button) {
        button.disabled = false;
        button.innerHTML = '<i class="fas fa-file-csv text-xs"></i> <span>Export CSV</span>';
      }
    }
  };

  const exportToHTML = async () => {
    try {
      // Show loading state
      const button = document.querySelector('.export-html-btn') as HTMLButtonElement;
      if (button) {
        button.disabled = true;
        if (isMultipleMeetings) {
          button.innerHTML = '<i class="fas fa-spinner fa-spin text-xs"></i> <span>Generating ' + selectedMeetingsForExport.size + ' HTML Files...</span>';
        } else {
          button.innerHTML = '<i class="fas fa-spinner fa-spin text-xs"></i> <span>Generating HTML...</span>';
        }
      }

      if (selectedMeetingsForExport.size === 0) {
        showError('Export Required', 'Please select at least one meeting to export');
        return;
      }

      const meetingsToExport = Array.from(selectedMeetingsForExport);
      
      if (isMultipleMeetings) {
        // Generate separate HTML file for each selected meeting
        for (let i = 0; i < meetingsToExport.length; i++) {
          const meetingDate = meetingsToExport[i];
          
          // Update button to show current progress
          if (button) {
            button.innerHTML = `<i class="fas fa-spinner fa-spin text-xs"></i> <span>HTML ${i + 1} of ${meetingsToExport.length}...</span>`;
          }
          
          const response = await fetch('/api/generate-meeting-html', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({
              meetingData: dataToUse,
              selectedMeeting: meetingDate,
              selectedType: 'all',
              meetingAttendance: meetingAttendance,
              meetingSignatures: meetingSignatures,
            }),
          });

          if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
          }

          // Get the JSON response with shareable URL
          const result = await response.json();
          
          if (!result.success) {
            throw new Error(result.error || 'Failed to generate HTML');
          }

          // Create downloadable blob from HTML content
          const blob = new Blob([result.htmlContent], { type: 'text/html' });
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.style.display = 'none';
          a.href = url;
          a.download = result.filename || 'meeting-minutes.html';
          document.body.appendChild(a);
          a.click();
          window.URL.revokeObjectURL(url);
          document.body.removeChild(a);

          // Show shareable URL in success modal
          showShareableUrlSuccess(result.shareUrl, result.filename);
          
          // Small delay between downloads to prevent browser blocking
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      } else {
        // Single HTML file generation
        const meetingDate = meetingsToExport[0];
        
        const response = await fetch('/api/generate-meeting-html', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            meetingData: dataToUse,
            selectedMeeting: meetingDate,
            selectedType: 'all',
            meetingAttendance: meetingAttendance,
            meetingSignatures: meetingSignatures,
          }),
        });

        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }

        // Get the JSON response with shareable URL
        const result = await response.json();
        
        if (!result.success) {
          throw new Error(result.error || 'Failed to generate HTML');
        }

        // Create downloadable blob from HTML content
        const blob = new Blob([result.htmlContent], { type: 'text/html' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = result.filename || 'meeting-minutes.html';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        // Show shareable URL in success modal
        showShareableUrlSuccess(result.shareUrl, result.filename);
      }

      // Success message handled by individual shareable URL modals
      if (isMultipleMeetings) {
        showSuccess('Multiple HTML Export Complete', 
          `Successfully exported ${selectedMeetingsForExport.size} HTML meeting minutes files with shareable URLs`
        );
      }

    } catch (error) {
      console.error('Error generating HTML:', error);
      showError('Export Failed', 'Failed to generate HTML. Please try again.');
    } finally {
      // Reset button state
      const button = document.querySelector('.export-html-btn') as HTMLButtonElement;
      if (button) {
        button.disabled = false;
        button.innerHTML = '<i class="fas fa-file-code text-xs"></i> <span>Export HTML</span>';
      }
    }
  };

  const handleMergeToNextMeeting = async (currentMeetingDate: string, categorized: Record<string, MeetingItem[]>) => {
    try {
      // Find the next upcoming (green) meeting from the existing meeting list
      // meetingDates is sorted descending, so filter upcoming ones and take the last (earliest future date)
      const upcomingMeetings = meetingDates.filter(d => getMeetingStatus(d).isUpcoming);
      const nextGreenMeeting = upcomingMeetings.at(-1);

      if (!nextGreenMeeting) {
        showError('No Upcoming Meeting', 'There is no upcoming meeting to move items to. Please create a scheduled meeting first.');
        return;
      }

      const nextMeetingDate = new Date(nextGreenMeeting);
      const formattedNextDate = nextMeetingDate.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
      
      // Count movable items (not closed) vs total items
      const movableStatus = getMeetingMovableStatus(categorized);

      // Set destination date to the next green meeting
      setSelectedDestinationDate(nextGreenMeeting.split('T')[0]);
      
      // Show modal with actual count
      setMergeConfirmationData({
        currentMeetingDate,
        nextMeetingDate,
        itemCount: movableStatus.totalCount,
        movableItemCount: movableStatus.movableCount,
        formattedNextDate
      });
      setShowMergeConfirmation(true);
      
    } catch (error) {
      console.error('Error preparing merge modal:', error);
      showError('Rescheduling Error', 'Failed to prepare move. Please try again.');
    }
  };

  // Actual merge function called when user confirms in modal
  const executeMerge = async () => {
    if (!mergeConfirmationData || !selectedDestinationDate) return;
    
    const { currentMeetingDate } = mergeConfirmationData;
    const destinationDate = new Date(selectedDestinationDate);
    
    try {
      setShowMergeConfirmation(false);
      setIsMergingMeeting(true);
      setMergeProgress({ current: 0, total: 0, currentItem: 'Initializing merge...' });

      // Get the SharePoint token
      const token = await authService.getSharePointToken();
      
      const response = await fetch('/api/merge-to-next-meeting', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          currentMeetingDate,
          nextMeetingDate: destinationDate.toISOString()
        })
      });

      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error || 'Failed to merge meeting');
      }

      // Show final success message with actual counts from backend
      setMergeProgress({ current: result.updated, total: result.total, currentItem: 'Merge completed!' });
      
      // Wait a moment to show completion, then show success modal
      const formattedDestDate = destinationDate.toLocaleDateString('en-GB');
      setTimeout(() => {
        setIsMergingMeeting(false);
        showSuccess(
          'Move Complete', 
          `Successfully moved ${result.updated}/${result.total} items to meeting on ${formattedDestDate}`
        );
      }, 1000);

    } catch (error) {
      console.error('Error moving agenda items:', error);
      setIsMergingMeeting(false);
      showError('Move Failed', 'Failed to move agenda items. Please try again.');
    }
  };

  const handleRescheduleMeeting = (meetingDate: string) => {
    setRescheduleMeetingSourceDate(meetingDate);
    setRescheduleTargetDate(new Date(meetingDate).toISOString().split('T')[0]);
    setShowRescheduleMeetingModal(true);
  };

  const executeRescheduleMeeting = async () => {
    if (!rescheduleMeetingSourceDate || !rescheduleTargetDate) return;

    const targetDate = new Date(rescheduleTargetDate);

    try {
      setShowRescheduleMeetingModal(false);
      setIsRescheduling(true);

      const token = await authService.getSharePointToken();

      const response = await fetch('/api/merge-to-next-meeting', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({
          currentMeetingDate: rescheduleMeetingSourceDate,
          nextMeetingDate: targetDate.toISOString()
        })
      });

      const result = await response.json();
      setIsRescheduling(false);

      if (!response.ok) {
        throw new Error(result.error || 'Failed to reschedule meeting');
      }

      const formattedDate = targetDate.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
      showSuccess('Meeting Rescheduled', `Successfully updated ${result.updated} item${result.updated !== 1 ? 's' : ''} to ${formattedDate} in SharePoint.`);

      // Optimistically update the local cache immediately so the UI reflects the change
      // SharePoint has a propagation delay so we can't rely on an immediate refetch
      const sourceDateStr = new Date(rescheduleMeetingSourceDate).toISOString().split('T')[0];
      const newMeetingDateISO = `${rescheduleTargetDate}T10:00:00.000Z`;

      const currentCacheData = queryClient.getQueryData(['/api/meeting-history']) as any;
      if (currentCacheData?.data) {
        const updatedData = {
          ...currentCacheData,
          data: currentCacheData.data.map((item: any) => {
            const itemDateStr = new Date(item.meetingDate).toISOString().split('T')[0];
            if (itemDateStr === sourceDateStr) {
              return { ...item, meetingDate: newMeetingDateISO };
            }
            return item;
          })
        };
        queryClient.setQueryData(['/api/meeting-history'], updatedData);
      }

      // Also schedule a real refetch after 5 seconds to sync with SharePoint
      setTimeout(() => {
        queryClient.invalidateQueries({ queryKey: ['/api/meeting-history'] });
      }, 5000);

    } catch (error) {
      console.error('Error rescheduling meeting:', error);
      setIsRescheduling(false);
      showError('Reschedule Failed', 'Failed to update meeting date in SharePoint. Please try again.');
    }
  };

  // Loading state for data
  if (isLoading) {
    return (
      <div className="min-h-screen bg-gray-50">
        <MeetingHeader />
        <div className="pt-16 sm:pt-20 bg-gray-50 min-h-screen">
          <div className="max-w-6xl mx-auto px-4 py-6 sm:py-8">
            <div className="text-center py-12">
              <div className="inline-block animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600 mb-4"></div>
              <p className="text-gray-600">Loading meeting records...</p>
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
        <div className="max-w-4xl mx-auto px-3 sm:px-6 py-4 sm:py-8">

          {/* Page Header - Mobile Optimized */}
          <div className="bg-gradient-to-br from-white to-gray-50 rounded-xl shadow-lg p-3 sm:p-6 mb-4 sm:mb-8 border border-gray-200">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2 sm:space-x-3">
                <div className="w-8 h-8 sm:w-10 sm:h-10 bg-gradient-to-r from-blue-600 to-blue-700 rounded-lg sm:rounded-xl flex items-center justify-center shadow-lg flex-shrink-0">
                  <FileText className="text-white" size={16} />
                </div>
                <div className="min-w-0 flex-1">
                  <h1 className="text-lg sm:text-2xl font-bold text-gray-900 leading-tight">Meeting History</h1>
                  <p className="text-xs text-gray-600 hidden sm:block">View and export meeting records for compliance documentation</p>
                </div>
              </div>
              {!isConfigured && (
                <div className="hidden sm:inline-flex items-center space-x-2 bg-amber-100 text-amber-800 px-3 py-2 rounded-lg text-xs font-medium flex-shrink-0 ml-4">
                  <i className="fas fa-exclamation-triangle"></i>
                  <span>SharePoint credentials required</span>
                </div>
              )}
            </div>
            {!isConfigured && (
              <div className="sm:hidden mt-3 inline-flex items-center space-x-2 bg-amber-100 text-amber-800 px-2 py-1.5 rounded-lg text-xs font-medium">
                <i className="fas fa-exclamation-triangle"></i>
                <span>SharePoint credentials required</span>
              </div>
            )}
          </div>

          {/* Meeting History Content */}
          <div>
              {/* Period Filter & Overview */}
              <div className="bg-white rounded-lg shadow-sm p-4 mb-6 border border-gray-200">
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 mb-4">
              <h2 className="text-lg font-semibold text-gray-900">
                <span className="inline sm:hidden lg:inline">Meeting Records</span>
                <span className="hidden sm:inline lg:hidden">Meeting History</span>
              </h2>
              
              {/* Clean Filter Controls */}
              <div className="flex flex-col sm:flex-row items-start sm:items-center gap-3">
                <select
                  value={selectedPeriod === 'actions' ? 'all' : selectedPeriod}
                  onChange={(e) => setSelectedPeriod(e.target.value)}
                  className="text-sm border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 w-full sm:w-auto"
                >
                  <option value="all">📅 All Time</option>
                  <option value="relevant">⚡ Relevant Items (Active & Upcoming)</option>
                  <option value="q1">🗓️ Q1 2025 (Jan-Mar)</option>
                  <option value="q2">🗓️ Q2 2025 (Apr-Jun)</option>
                  <option value="q3">🗓️ Q3 2025 (Jul-Sep)</option>
                  <option value="q4">🗓️ Q4 2025 (Oct-Dec)</option>
                  <option value="last6months">📊 Last 6 Months</option>
                </select>
                
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => setLocation('/actions')}
                  className="flex items-center gap-2 w-full sm:w-auto hover:bg-amber-50 hover:border-amber-300"
                >
                  <i className="fas fa-tasks text-sm"></i>
                  Open Actions
                </Button>
                
                <Button
                  variant={selectedMeetingForStats === 'all' ? 'outline' : 'default'}
                  size="sm"
                  onClick={() => {
                    if (selectedMeetingForStats === 'all') {
                      // Switch to live stats - find the next upcoming green meeting
                      const now = new Date();
                      const today = now.toISOString().split('T')[0]; // Get YYYY-MM-DD format
                      
                      // First try to find today's meeting
                      const todaysMeeting = Object.keys(groupedByMeeting).find(meetingDate => {
                        const meetingDateStr = new Date(meetingDate).toISOString().split('T')[0];
                        return meetingDateStr === today;
                      });
                      
                      if (todaysMeeting) {
                        // Show today's meeting for live stats
                        setSelectedMeetingForStats(todaysMeeting);
                      } else {
                        // Find the next upcoming meeting (green header) instead of closest meeting
                        const upcomingMeetings = Object.keys(groupedByMeeting)
                          .filter(meetingDate => {
                            const status = getMeetingStatus(meetingDate);
                            return status.isUpcoming; // Only upcoming meetings (green headers)
                          })
                          .sort((a, b) => new Date(a).getTime() - new Date(b).getTime()); // Earliest upcoming first
                        
                        if (upcomingMeetings.length > 0) {
                          setSelectedMeetingForStats(upcomingMeetings[0]); // Select the next upcoming meeting
                        } else {
                          // Fallback to most recent past meeting if no upcoming meetings
                          const pastMeetings = Object.keys(groupedByMeeting)
                            .filter(meetingDate => {
                              const status = getMeetingStatus(meetingDate);
                              return status.isPast;
                            })
                            .sort((a, b) => new Date(b).getTime() - new Date(a).getTime()); // Most recent past first
                          
                          if (pastMeetings.length > 0) {
                            setSelectedMeetingForStats(pastMeetings[0]);
                          }
                        }
                      }
                    } else {
                      // Switch back to all meetings
                      setSelectedMeetingForStats('all');
                    }
                  }}
                  className={`flex items-center gap-2 w-full sm:w-auto ${
                    selectedMeetingForStats !== 'all' 
                      ? 'bg-green-600 hover:bg-green-700 text-white' 
                      : 'hover:bg-gray-50'
                  }`}
                >
                  {selectedMeetingForStats === 'all' ? (
                    <>
                      <div className="w-2 h-2 bg-green-500 rounded-full animate-pulse"></div>
                      Live Stats
                    </>
                  ) : (
                    <>
                      <div className="w-2 h-2 bg-white rounded-full animate-pulse"></div>
                      Live: {format(new Date(selectedMeetingForStats), 'MMM d')}
                    </>
                  )}
                </Button>
              </div>
            </div>

            {/* Analytics Dashboard Section */}
            <div className="mb-8">
              <h3 className="text-sm font-medium text-gray-700 mb-3">
                {selectedMeetingForStats === 'all' ? 'Meeting History Analytics' : 'Live Meeting Statistics'}
              </h3>
              <DashboardStats 
                meetings={groupedByMeetingAndCategory}
                items={filteredItems}
                period={selectedPeriod}
                selectedMeeting={selectedMeetingForStats}
              />
            </div>

            {/* Search Section - Separate Card */}
            <div className="bg-white rounded-lg border border-gray-200 p-4 mb-4">
              <div className="flex items-center gap-2">
                <div className="flex-1 relative">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                    <i className="fas fa-search text-gray-400"></i>
                  </div>
                  <input
                    type="text"
                    placeholder="Search ideas, meeting notes, titles, staff names, or status..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="block w-full pl-10 pr-3 py-2 border border-gray-300 rounded-md leading-5 bg-white placeholder-gray-500 focus:outline-none focus:placeholder-gray-400 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm"
                  />
                </div>
                {searchQuery && (
                  <button
                    onClick={() => setSearchQuery('')}
                    className="px-3 py-2 border border-gray-300 text-sm rounded-md text-gray-600 hover:bg-gray-50"
                    title="Clear"
                  >
                    <i className="fas fa-times"></i>
                  </button>
                )}
              </div>
              {searchQuery && (
                <div className="mt-2 text-xs text-gray-600">
                  {filteredItems.length} result{filteredItems.length !== 1 ? 's' : ''} found
                </div>
              )}
            </div>

            {/* Action Controls - Sticky Card */}
            <div className="sticky top-0 z-10 bg-white rounded-lg border border-gray-200 p-4 mb-6 shadow-sm">
              <div className="flex flex-wrap gap-2">
                {/* Single Export Button */}
                <button
                  onClick={() => setShowExportSelector(!showExportSelector)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm text-gray-700 bg-white hover:bg-gray-50 border border-gray-300 rounded-md"
                >
                  <i className="fas fa-download text-xs"></i>
                  <span className="sm:hidden">Export</span>
                  <span className="hidden sm:inline">Export Meeting Minutes</span>
                  {selectedMeetingsForExport.size > 0 && (
                    <span className="px-1.5 py-0.5 rounded-full text-xs bg-white/20 text-white">
                      {selectedMeetingsForExport.size}
                    </span>
                  )}
                </button>
                
                {/* AI Titles Button - Only show when titles need updating */}
                {aiTitleStatus.needsAI && (
                  <button
                    onClick={generateAITitles}
                    disabled={isGeneratingTitles}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm text-gray-700 bg-white hover:bg-gray-50 border border-gray-300 rounded-md"
                    title={`Generate titles for ${aiTitleStatus.emptyTitles} items`}
                  >
                    <i className={`fas fa-robot text-green-600 text-xs ${isGeneratingTitles ? 'animate-spin' : ''}`}></i>
                    <span>AI Titles ({aiTitleStatus.emptyTitles})</span>
                  </button>
                )}

                <div className="flex-1"></div>

                {/* Single Toggle Button */}
                <button
                  onClick={() => {
                    const allExpanded = expandedMeetings.size === groupedByMeetingAndCategory.length;
                    if (allExpanded) {
                      collapseAllMeetings();
                    } else {
                      expandAllMeetings();
                    }
                  }}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-sm text-gray-700 bg-white hover:bg-gray-50 border border-gray-300 rounded-md"
                >
                  <i className={`fas fa-chevron-${expandedMeetings.size === groupedByMeetingAndCategory.length ? 'up' : 'down'} text-xs`}></i>
                  <span>{expandedMeetings.size === groupedByMeetingAndCategory.length ? 'Collapse All' : 'Expand All'}</span>
                </button>
              </div>

              {/* Export Selector */}
              {showExportSelector && (
                <div className="mt-3 pt-3 border-t border-gray-200">
                  <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className="text-sm text-gray-600">Select meetings and format:</span>
                      <span className="text-xs text-gray-500 bg-gray-100 px-2 py-1 rounded">v1.3.0</span>
                    </div>
                    <div className="text-xs text-gray-500 w-full sm:w-auto">
                      HTML: Browser view • CSV: Data analysis • MD: Typora PDF with navigation
                    </div>
                    {selectedMeetingsForExport.size > 0 && (
                      <div className="flex gap-2">
                        <button
                          onClick={exportToHTML}
                          className="export-html-btn inline-flex items-center gap-1.5 bg-blue-600 hover:bg-blue-700 text-white border border-blue-600 px-3 py-1.5 rounded text-xs"
                        >
                          <i className="fas fa-file-code text-xs"></i>
                          <span>HTML ({selectedMeetingsForExport.size})</span>
                        </button>
                        <button
                          onClick={exportToCSV}
                          className="export-csv-btn inline-flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white border border-green-600 px-3 py-1.5 rounded text-xs"
                        >
                          <i className="fas fa-file-csv text-xs"></i>
                          <span>CSV ({selectedMeetingsForExport.size})</span>
                        </button>

                      </div>
                    )}
                  </div>
                  

                  <div className="max-h-32 overflow-y-auto space-y-1">
                    {meetingDates.map(date => (
                      <label
                        key={date}
                        className="flex items-center gap-2 p-2 hover:bg-gray-50 rounded cursor-pointer"
                      >
                        <input
                          type="checkbox"
                          checked={selectedMeetingsForExport.has(date)}
                          onChange={() => toggleMeetingForExport(date)}
                          disabled={!selectedMeetingsForExport.has(date) && selectedMeetingsForExport.size >= 3}
                          className="h-4 w-4 text-green-600 border-gray-300 rounded"
                        />
                        <div className="flex-1">
                          <span className="text-sm text-gray-700">
                            {formatDate(date)}
                          </span>
                          <span className="ml-2 text-xs text-gray-500">
                            ({dataToUse.filter(item => isSameDay(item.meetingDate, date)).length})
                          </span>
                        </div>
                      </label>
                    ))}
                  </div>
                  {selectedMeetingsForExport.size >= 3 && (
                    <div className="mt-2 text-xs text-amber-600">
                      <i className="fas fa-info-circle mr-1"></i>
                      Max 3 meetings
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>





          <div className={`transition-all duration-300 ${groupedByMeetingAndCategory.length === 0 ? 'opacity-100' : 'opacity-100'}`}>
            {groupedByMeetingAndCategory.length === 0 ? (
              <div className="bg-white rounded-xl shadow-sm p-6 sm:p-8 text-center border border-gray-200 animate-in fade-in duration-200">
                <FileText className="mx-auto h-10 w-10 sm:h-12 sm:w-12 text-gray-400 mb-4" />
                <h3 className="text-base sm:text-lg font-medium text-gray-900 mb-2">
                  {searchQuery ? 'No Search Results' : 'No Records Found'}
                </h3>
                <p className="text-sm sm:text-base text-gray-600 max-w-md mx-auto">
                  {searchQuery 
                    ? `No items found matching "${searchQuery}". Try a different search term or clear the search to see all records.`
                    : isConfigured 
                      ? "No closed items found in your SharePoint lists matching the current filters."
                      : "No meeting records match your current filters."
                  }
                </p>
                {searchQuery && (
                  <button
                    onClick={() => setSearchQuery('')}
                    className="mt-4 inline-flex items-center px-4 py-2 border border-transparent text-sm font-medium rounded-md text-white bg-blue-600 hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    Clear Search
                  </button>
                )}
              </div>
            ) : (
              <div className="space-y-6">
                {groupedByMeetingAndCategory.map(({ meetingDate, categorized, totalItems }) => (
              <Card key={meetingDate} className="overflow-hidden">
                {/* Print-only meeting header */}
                <div className="print-header" style={{display: 'none'}}>
                  <div className="print-title">CRANFIELD GLASS CHRISTCHURCH</div>
                  <div className="print-subtitle">Health & Safety Committee Meeting Minutes</div>
                  <div className="print-date">{formatDate(meetingDate)}</div>
                  
                  <div className="print-meeting-details">
                    <div>
                      <h4><strong>Meeting Attendees:</strong></h4>
                      {Array.from(new Set(Object.values(categorized).flat().map((item: MeetingItem) => item.submittedBy))).map((person, index) => (
                        <div key={index}>{person} - Present</div>
                      ))}
                      <div>_________________________ - _______</div>
                      <div>_________________________ - _______</div>
                    </div>
                    <div>
                      <h4><strong>Meeting Details:</strong></h4>
                      <div><strong>Date:</strong> {formatDate(meetingDate)}</div>
                      <div><strong>Time:</strong> ________________</div>
                      <div><strong>Location:</strong> ________________</div>
                      <div><strong>Chair:</strong> ________________</div>
                      <div><strong>Secretary:</strong> ________________</div>
                    </div>
                  </div>
                </div>

                <CardHeader 
                  className={`${getMeetingHeaderStyle(meetingDate, getMeetingMovableStatus(categorized).allItemsClosed && getMeetingStatus(meetingDate).isPast)} p-4 sm:p-6 relative`}
                >
                  <CardTitle className="flex items-center gap-3">
                    <Calendar className="h-5 w-5 flex-shrink-0" />
                    <div 
                      className="min-w-0 flex-1 cursor-pointer"
                      onClick={() => toggleMeetingExpansion(meetingDate)}
                    >
                      <div className="text-base sm:text-lg font-bold truncate">
                        {getMeetingTitle(meetingDate)}: {formatDate(meetingDate)}
                      </div>
                      <div className="text-white/80 text-sm font-normal">{getStatusSummary(categorized)}</div>
                    </div>
                    <div className="flex items-center gap-2 flex-shrink-0">
                      {/* Reschedule pencil - only on upcoming (green) meetings */}
                      {getMeetingStatus(meetingDate).isUpcoming && (
                        <button
                          onClick={(e) => {
                            e.stopPropagation();
                            handleRescheduleMeeting(meetingDate);
                          }}
                          className="p-1 rounded hover:bg-white/20 transition-colors text-white/70 hover:text-white"
                          title="Change meeting date"
                        >
                          <i className="fas fa-pencil-alt text-xs"></i>
                        </button>
                      )}
                      {/* Move Agenda Items Button - Only show on past (purple) meetings with movable items */}
                      {(() => {
                        const movableStatus = getMeetingMovableStatus(categorized);
                        if (!movableStatus.hasMovableItems || !getMeetingStatus(meetingDate).isPast) return null;
                        return (
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              const allMeetingItems = Object.values(categorized).flat();
                              const actualSharePointDate = allMeetingItems.length > 0 ? allMeetingItems[0].meetingDate : meetingDate;
                              handleMergeToNextMeeting(actualSharePointDate, categorized);
                            }}
                            className="px-3 py-1 rounded-md transition-colors duration-200 border flex items-center gap-2 text-xs font-medium bg-white/20 hover:bg-white/30 text-white border-white/30"
                            title={`Move ${movableStatus.movableCount} Submitted item${movableStatus.movableCount !== 1 ? 's' : ''} to the next scheduled meeting`}
                          >
                            <i className="fas fa-arrow-right text-xs"></i>
                            <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-white/30">
                              {movableStatus.movableCount}
                            </span>
                          </button>
                        );
                      })()}
                      <div 
                        className="text-white/60 text-xs hidden sm:inline cursor-pointer"
                        onClick={() => toggleMeetingExpansion(meetingDate)}
                      >
                        {expandedMeetings.has(meetingDate) ? 'Click to collapse' : 'Click to expand'}
                      </div>
                      <div 
                        className={`transform transition-transform duration-200 cursor-pointer ${expandedMeetings.has(meetingDate) ? 'rotate-180' : ''}`}
                        onClick={() => toggleMeetingExpansion(meetingDate)}
                      >
                        <i className="fas fa-chevron-down text-white/80"></i>
                      </div>
                    </div>
                  </CardTitle>
                </CardHeader>
                
                {/* Meeting Content - Collapsible */}
                {expandedMeetings.has(meetingDate) && (
                  <CardContent className="p-0">
                    {/* Attendance Management Section - Always under Meeting Header */}
                    <div className="bg-white border-b border-gray-100 p-3 sm:p-6">
                      <div className="flex flex-col sm:flex-row gap-3 sm:gap-4">
                        <Button
                          variant="outline"
                          onClick={() => {
                            if (showAttendanceSelector === meetingDate) {
                              setShowAttendanceSelector('');
                            } else {
                              initializeAttendance(meetingDate);
                              setShowAttendanceSelector(meetingDate);
                            }
                          }}
                          className="w-full sm:w-auto flex items-center justify-center gap-2 text-gray-700 bg-white hover:bg-gray-50 hover:text-blue-700 hover:border-blue-300"
                        >
                          <Users className="h-4 w-4" />
                          <span>{showAttendanceSelector === meetingDate ? 'Hide' : 'Manage'} Attendance</span>
                          {showAttendanceSelector === meetingDate ? <ChevronDown className="h-4 w-4" /> : <ChevronRight className="h-4 w-4" />}
                        </Button>

                        {(() => {
                          const sigs = meetingSignatures[meetingDate] ?? {};
                          const sigCount = Object.keys(sigs).length;
                          const isLocked = !!lockedAttendance[normaliseLockDate(meetingDate)];
                          return (
                            <Button
                              variant="outline"
                              onClick={() => !isLocked && setShowSignatureCarousel(meetingDate)}
                              disabled={isLocked}
                              title={isLocked ? 'Attendance is locked — signatures cannot be modified' : undefined}
                              className={`w-full sm:w-auto flex items-center justify-center gap-2 border-indigo-200 ${isLocked ? 'opacity-50 cursor-not-allowed text-indigo-400 bg-white' : 'text-indigo-700 bg-white hover:bg-indigo-50 hover:border-indigo-300'}`}
                            >
                              <PenLine className="h-4 w-4" />
                              <span>Collect Signatures</span>
                              {sigCount > 0 && (
                                <span className="ml-1 bg-indigo-100 text-indigo-700 text-xs font-semibold px-2 py-0.5 rounded-full">{sigCount}</span>
                              )}
                            </Button>
                          );
                        })()}
                      </div>
                    </div>

                    {/* Attendance Selector - Within collapsed content */}
                    {showAttendanceSelector === meetingDate && (() => {
                      const isLocked = lockedAttendance[normaliseLockDate(meetingDate)];
                      
                      return (
                        <div className="bg-gradient-to-br from-white to-blue-50 border-b border-blue-200 p-4 sm:p-6">
                          <div className="flex items-center gap-3 mb-4">
                            <div className="w-8 h-8 bg-gradient-to-r from-blue-500 to-blue-600 rounded-lg flex items-center justify-center shadow-md">
                              <Users className="h-4 w-4 text-white" />
                            </div>
                            <h3 className="text-base sm:text-lg font-bold text-gray-900">Meeting Attendance</h3>
                            <div className="flex items-center gap-2 ml-auto">
                              <Button
                                variant="outline"
                                size="sm"
                                onClick={() => toggleAttendanceLock(meetingDate)}
                                className={`flex items-center gap-2 ${
                                  isLocked 
                                    ? 'text-red-700 hover:text-red-800 border-red-300 hover:border-red-400 bg-red-50 hover:bg-red-100'
                                    : 'text-green-700 hover:text-green-800 border-green-300 hover:border-green-400 bg-green-50 hover:bg-green-100'
                                }`}
                              >
                                {isLocked ? (
                                  <>
                                    <Lock className="h-4 w-4" />
                                    <span>Unlock</span>
                                  </>
                                ) : (
                                  <>
                                    <Unlock className="h-4 w-4" />
                                    <span>Lock</span>
                                  </>
                                )}
                              </Button>
                              {isLocked && (
                                <div className="flex items-center gap-2">
                                  <div className="w-4 h-4 bg-red-400 rounded-full flex items-center justify-center">
                                    <Lock className="h-2 w-2 text-white" />
                                  </div>
                                  <span className="text-sm text-red-600 font-medium">Locked</span>
                                </div>
                              )}
                            </div>
                          </div>
                          
                          {/* Compact single section with all attendees */}
                          <div className="bg-white rounded-xl border border-blue-200 p-4 shadow-sm">
                            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2">
                              {[...meetingAttendees.management, ...meetingAttendees.glaziers].map((attendee) => {
                                const sig = (meetingSignatures[meetingDate] ?? {})[attendee.name];
                                const sigBadge = sig
                                  ? sig.status === 'signed' ? { label: '✍️', cls: 'bg-green-100 text-green-700' }
                                  : sig.status === 'remote' ? { label: '🖥️', cls: 'bg-purple-100 text-purple-700' }
                                  : { label: '—', cls: 'bg-gray-100 text-gray-400' }
                                  : null;
                                return (
                                  <label
                                    key={attendee.name}
                                    className={`flex items-center gap-2 p-2 rounded-lg transition-colors text-sm ${
                                      isLocked
                                        ? 'cursor-not-allowed opacity-60'
                                        : 'cursor-pointer hover:bg-blue-50'
                                    }`}
                                  >
                                    <input
                                      type="checkbox"
                                      checked={isAttending(meetingDate, attendee.name)}
                                      onChange={() => !isLocked && toggleAttendee(meetingDate, attendee.name)}
                                      disabled={isLocked}
                                      className={`h-4 w-4 text-blue-600 border-2 border-gray-300 rounded transition-all ${
                                        isLocked ? 'cursor-not-allowed' : 'focus:ring-2 focus:ring-blue-500'
                                      }`}
                                    />
                                    <div className="flex-1 min-w-0">
                                      <div className="font-medium text-gray-900 truncate">{attendee.name}</div>
                                      <div className="text-xs text-blue-600 truncate">{attendee.role}</div>
                                    </div>
                                    {sigBadge && (
                                      <div className="flex flex-col items-end gap-1 flex-shrink-0">
                                        <span className={`text-xs px-1.5 py-0.5 rounded-full font-medium ${sigBadge.cls}`}>
                                          {sigBadge.label}
                                        </span>
                                        {sig?.status === 'signed' && sig.signatureData && (
                                          <img
                                            src={sig.signatureData}
                                            alt={`${attendee.name} signature`}
                                            className="h-6 w-16 object-contain border border-gray-200 rounded bg-white"
                                          />
                                        )}
                                      </div>
                                    )}
                                  </label>
                                );
                              })}
                            </div>
                          </div>
                        </div>
                      );
                    })()}
                
                  <div className="space-y-3 sm:space-y-4 p-3 sm:p-6">
                    {/* Pending/optimistic items skeleton - appear immediately after saving */}
                    {pendingItems
                      .filter(p => p.meetingDate === meetingDate)
                      .map(p => (
                        <div key={p.id} className="rounded-xl border border-blue-200 bg-blue-50 p-4 shadow-sm">
                          <div className="flex items-center gap-2 mb-3">
                            <Loader2 className="h-4 w-4 text-blue-500 animate-spin flex-shrink-0" />
                            <span className="text-sm font-semibold text-blue-700">Saving to SharePoint…</span>
                            <span className="text-xs bg-blue-200 text-blue-800 px-2 py-0.5 rounded-full ml-auto">{p.type}</span>
                          </div>
                          <div className="space-y-2 animate-pulse">
                            <div className="h-3 bg-blue-200 rounded-full w-2/3"></div>
                            <div className="h-3 bg-blue-200 rounded-full w-1/2"></div>
                          </div>
                          <p className="text-xs text-blue-500 mt-3 line-clamp-1 italic">
                            "{p.description.length > 70 ? p.description.substring(0, 70) + '…' : p.description}"
                          </p>
                          <p className="text-xs text-blue-400 mt-1">Submitted by {p.submittedBy}</p>
                        </div>
                      ))
                    }

                    {Object.entries(categorized).map(([categoryType, categoryItems]) => {
                      if (categoryItems.length === 0) return null;
                      
                      const categoryKey = `${meetingDate}-${categoryType}`;
                      const isExpanded = expandedCategories.has(categoryKey);
                      const IconComponent = getCategoryIcon(categoryType);
                      
                      return (
                        <div key={categoryType} className="w-full">
                          <Collapsible open={isExpanded} onOpenChange={(open) => {
                            if (open) {
                              toggleCategoryExpansion(categoryKey, meetingDate);
                            } else {
                              collapseCategory(categoryKey);
                            }
                          }}>
                            <CollapsibleTrigger asChild>
                              <Button
                                variant="ghost"
                                className={`w-full justify-between p-3 sm:p-4 h-auto ${getCategoryColor(categoryType)} border rounded-lg hover:bg-opacity-80 mb-2 touch-manipulation min-h-[48px] sm:min-h-[44px]`}
                              >
                                <div className="flex items-center gap-2 sm:gap-3 min-w-0 flex-1">
                                  <IconComponent className="h-4 w-4 sm:h-5 sm:w-5 flex-shrink-0" />
                                  <div className="text-left min-w-0 flex-1">
                                    <div className="font-semibold text-sm sm:text-base leading-tight line-clamp-1">{categoryType}</div>
                                    <div className="text-xs opacity-75 mt-0.5">{categoryItems.length} item{categoryItems.length !== 1 ? 's' : ''}</div>
                                  </div>
                                </div>
                                {isExpanded ? <ChevronDown className="h-4 w-4 flex-shrink-0" /> : <ChevronRight className="h-4 w-4 flex-shrink-0" />}
                              </Button>
                            </CollapsibleTrigger>
                            <CollapsibleContent className="w-full">
                              <div className="space-y-3 pl-1 sm:pl-2">
                                {categoryItems.map((item: MeetingItem) => {
                                  // Determine if card should be flipped
                                  // All items default to showing details (front face)
                                  // User can manually flip to see actions by clicking the flip button
                                  // No auto-flip - everything is manual
                                  const isFlipped = flippedCards.has(item.id);
                                  
                                  return (
                                    <div key={item.id} className="flip-card-container" style={{ perspective: '1000px' }}>
                                      <div 
                                        className={`flip-card-inner relative transition-transform duration-500`}
                                        style={{ 
                                          transformStyle: 'preserve-3d',
                                          transform: isFlipped ? 'rotateY(180deg)' : 'rotateY(0deg)'
                                        }}
                                      >
                                        {/* FRONT FACE - Details */}
                                        <div 
                                          className="flip-card-front bg-white rounded-lg border border-gray-200 shadow-sm hover:shadow-md transition-shadow"
                                          style={{ 
                                            backfaceVisibility: 'hidden',
                                            WebkitBackfaceVisibility: 'hidden'
                                          }}
                                        >
                                          {/* Item Header */}
                                          <div className="p-3 sm:p-4">
                                            <div className="flex flex-col gap-3">
                                              {/* Title and Badges Row */}
                                              <div className="flex items-start justify-between gap-2">
                                                <h4 className="font-semibold text-gray-900 text-sm sm:text-base leading-tight flex-1 min-w-0 pr-2">
                                                  {item.title && item.title.trim() !== '' ? (
                                                    <span className="line-clamp-2">{highlightSearchTerm(item.title, searchQuery)}</span>
                                                  ) : (
                                                    <span className="text-gray-400 italic">No title</span>
                                                  )}
                                                </h4>
                                                <div className="flex flex-wrap gap-1.5 flex-shrink-0">
                                                  {item.ideaType && (
                                                    <Badge variant="secondary" className="text-xs bg-blue-100 text-blue-800 border-blue-200 whitespace-nowrap">
                                                      {item.ideaType}
                                                    </Badge>
                                                  )}
                                                  <Badge variant="outline" className={`text-xs whitespace-nowrap ${getStatusBadgeColor(item.status, item.type)}`}>
                                                    {item.status}
                                                  </Badge>
                                                </div>
                                              </div>
                                              
                                              {/* Status Update Row */}
                                              <div className="flex items-center gap-2">
                                                <label className="text-xs text-gray-600 font-medium">Status:</label>
                                                <select
                                                  value={item.status}
                                                  onChange={(e) => updateItemStatus(item, e.target.value)}
                                                  disabled={isUpdatingStatus === item.id}
                                                  className="text-xs border border-gray-300 rounded-md px-2 py-1.5 bg-white hover:bg-gray-50 transition-colors min-w-0 flex-1 sm:flex-none sm:min-w-[120px]"
                                                  title="Update status"
                                                  data-testid={`select-status-${item.id}`}
                                                >
                                                  {statusOptions
                                                    .filter(option => {
                                                      if (option === 'Closed') {
                                                        return item.type === 'Business Ideas' || item.type === 'Safety Ideas' || item.type === 'Near Miss';
                                                      }
                                                      return true;
                                                    })
                                                    .map(option => (
                                                      <option key={option} value={option}>{option}</option>
                                                    ))}
                                                </select>
                                                {/* Flip to Actions button - only show for Actioned status OR Closed with action data */}
                                                {(() => {
                                                  const hasActionData = !!(item.actionPriority || item.actionStatus || item.actionAssignedTo || item.actionStartDate || item.actionDueDate || item.actionNotes);
                                                  const canFlip = item.status === 'Actioned' || (item.status === 'Closed' && hasActionData);
                                                  
                                                  if (!canFlip) return null;
                                                  
                                                  return (
                                                    <button
                                                      onClick={() => toggleFlipCard(item.id)}
                                                      className={`inline-flex items-center gap-1 px-2 py-1 text-xs rounded-md transition-colors ${
                                                        item.status === 'Closed' 
                                                          ? 'text-green-700 bg-green-50 hover:bg-green-100 border border-green-300' 
                                                          : 'text-orange-600 bg-orange-50 hover:bg-orange-100 border border-orange-200'
                                                      }`}
                                                      title={item.status === 'Closed' ? "View completed action details" : "View action details"}
                                                      data-testid={`button-flip-to-actions-${item.id}`}
                                                    >
                                                      <i className={`fas ${item.status === 'Closed' ? 'fa-check-circle' : 'fa-exchange-alt'} text-xs`}></i>
                                                      <span className="hidden sm:inline">{item.status === 'Closed' ? 'Actions' : 'Actions'}</span>
                                                      {item.status === 'Closed' && hasActionData && (
                                                        <span className="ml-1 w-2 h-2 bg-green-500 rounded-full animate-pulse"></span>
                                                      )}
                                                    </button>
                                                  );
                                                })()}
                                              </div>
                                            </div>
                                          </div>
                                          
                                          {/* Details Content */}
                                          <div className="p-3 sm:p-4 pt-0">
                                            {/* Near Miss Special Layout */}
                                            {item.type === 'Near Miss' ? (
                                              <div className="space-y-3 mb-3">
                                                <div className="bg-orange-50 border border-orange-200 rounded-lg p-3">
                                                  <div className="flex items-center gap-2 mb-2">
                                                    <AlertTriangle className="h-4 w-4 text-orange-600" />
                                                    <div className="font-semibold text-orange-900 text-sm">What Happened</div>
                                                  </div>
                                                  <div className="text-sm text-orange-800 leading-relaxed">{highlightSearchTerm(item.description, searchQuery)}</div>
                                                </div>
                                                {item.secondaryDescription && item.secondaryDescription.trim() && (
                                                  <div className="bg-red-50 border border-red-200 rounded-lg p-3">
                                                    <div className="flex items-center gap-2 mb-2">
                                                      <AlertTriangle className="h-4 w-4 text-red-600" />
                                                      <div className="font-semibold text-red-900 text-sm">How It Happened</div>
                                                    </div>
                                                    <div className="text-sm text-red-800 leading-relaxed">{highlightSearchTerm(item.secondaryDescription, searchQuery)}</div>
                                                  </div>
                                                )}
                                              </div>
                                            ) : (
                                              <div className="bg-gray-50 border border-gray-200 rounded-lg p-3 mb-3">
                                                <p className="text-sm text-gray-700 leading-relaxed">{highlightSearchTerm(item.description, searchQuery)}</p>
                                              </div>
                                            )}
                                            
                                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-sm text-gray-600 mb-3">
                                              <div className="bg-white border border-gray-100 rounded-lg p-3 space-y-2">
                                                <div className="flex flex-col sm:flex-row sm:items-center gap-1">
                                                  <span className="font-medium text-gray-700 text-xs uppercase tracking-wide">Submitted by:</span> 
                                                  <span className="line-clamp-1">{highlightSearchTerm(item.submittedBy, searchQuery)}</span>
                                                </div>
                                                <div className="flex flex-col sm:flex-row sm:items-center gap-1">
                                                  <span className="font-medium text-gray-700 text-xs uppercase tracking-wide">Date:</span> 
                                                  <span>{formatDate(item.submittedDate)}</span>
                                                </div>
                                              </div>
                                              <div className="bg-white border border-gray-100 rounded-lg p-3">
                                                <div className="flex flex-col sm:flex-row sm:items-center gap-1">
                                                  <span className="font-medium text-gray-700 text-xs uppercase tracking-wide">Status:</span> 
                                                  <span>{highlightSearchTerm(item.status, searchQuery)}</span>
                                                </div>
                                              </div>
                                            </div>
                                            
                                            {/* Meeting Notes */}
                                            <div className="bg-gray-50 border border-gray-200 rounded-lg p-3 mt-3">
                                              <div className="flex items-center justify-between mb-2">
                                                <div className="flex items-center gap-2">
                                                  <FileText className="h-4 w-4 text-gray-600" />
                                                  <div className="font-semibold text-gray-900 text-sm">Meeting Notes</div>
                                                </div>
                                                <button
                                                  onClick={() => openNotesModal(item)}
                                                  className="inline-flex items-center gap-1 px-3 py-1 text-xs text-gray-600 bg-white hover:bg-gray-50 border border-gray-300 rounded-md transition-colors shadow-sm"
                                                  title={item.meetingNotes?.trim() ? "Edit meeting notes" : "Add meeting notes"}
                                                  data-testid={`button-edit-notes-${item.id}`}
                                                >
                                                  <i className={`fas ${item.meetingNotes?.trim() ? 'fa-edit' : 'fa-plus'} text-xs`}></i>
                                                  <span>{item.meetingNotes?.trim() ? 'Edit Notes' : 'Add Notes'}</span>
                                                </button>
                                              </div>
                                              {item.meetingNotes && item.meetingNotes.trim() ? (
                                                <div className={`text-sm text-gray-700 leading-relaxed rounded border p-2 transition-all duration-500 ${
                                                  recentlySavedNotes.has(item.id) 
                                                    ? 'bg-green-50 border-green-300 shadow-md' 
                                                    : 'bg-white border-gray-200'
                                                }`}>
                                                  {recentlySavedNotes.has(item.id) && (
                                                    <div className="flex items-center text-green-700 mb-1 text-xs">
                                                      <CheckCircle className="h-3 w-3 mr-1" />
                                                      Notes saved successfully
                                                    </div>
                                                  )}
                                                  {highlightSearchTerm(item.meetingNotes, searchQuery)}
                                                </div>
                                              ) : (
                                                <div className="text-sm text-gray-500 italic">
                                                  No meeting notes added yet. Click "Add Notes" to add discussion points.
                                                </div>
                                              )}
                                            </div>
                                          </div>
                                        </div>
                                        
                                        {/* BACK FACE - Actions (Yellow/Orange Gradient) */}
                                        <div 
                                          className="flip-card-back rounded-lg border-2 border-amber-400 shadow-lg absolute top-0 left-0 w-full h-full"
                                          style={{ 
                                            backfaceVisibility: 'hidden',
                                            WebkitBackfaceVisibility: 'hidden',
                                            transform: 'rotateY(180deg)',
                                            background: 'linear-gradient(135deg, #fef3c7 0%, #fde68a 25%, #fcd34d 50%, #fbbf24 75%, #f59e0b 100%)'
                                          }}
                                        >
                                          {/* Action Card Header */}
                                          <div className="p-3 sm:p-4 border-b border-amber-300">
                                            <div className="flex items-start justify-between gap-2">
                                              <div className="flex-1 min-w-0">
                                                <div className="flex items-center gap-2 mb-1">
                                                  <i className="fas fa-tasks text-amber-700"></i>
                                                  <span className="font-bold text-amber-900 text-sm uppercase tracking-wide">Action Card</span>
                                                </div>
                                                <h4 className="font-semibold text-amber-900 text-sm sm:text-base leading-tight line-clamp-2">
                                                  {item.title && item.title.trim() !== '' ? item.title : 'Untitled Item'}
                                                </h4>
                                              </div>
                                              {/* Flip back button */}
                                              <button
                                                onClick={() => toggleFlipCard(item.id)}
                                                className="inline-flex items-center gap-1 px-2 py-1 text-xs text-amber-800 bg-white/80 hover:bg-white border border-amber-400 rounded-md transition-colors shadow-sm"
                                                title="View original details"
                                                data-testid={`button-flip-to-details-${item.id}`}
                                              >
                                                <i className="fas fa-exchange-alt text-xs"></i>
                                                <span className="hidden sm:inline">Details</span>
                                              </button>
                                            </div>
                                          </div>
                                          
                                          {/* Action Content */}
                                          <div className="p-3 flex flex-col gap-2">
                                            {/* Compact summary row + Edit toggle */}
                                            <div className="flex items-center justify-between gap-2">
                                              <div className="flex flex-wrap gap-1 items-center flex-1 min-w-0">
                                                {item.actionPriority ? (
                                                  <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${item.actionPriority === 'High' ? 'bg-red-200 text-red-800' : item.actionPriority === 'Medium' ? 'bg-amber-200 text-amber-900' : 'bg-green-200 text-green-800'}`}>{item.actionPriority}</span>
                                                ) : null}
                                                {item.actionStatus ? (
                                                  <span className="px-1.5 py-0.5 rounded text-[10px] bg-white/70 text-amber-900 border border-amber-300">{item.actionStatus || 'Not Started'}</span>
                                                ) : null}
                                                {item.actionAssignedTo ? (
                                                  <span className="text-xs text-amber-800 truncate max-w-[100px]">→ {item.actionAssignedTo}</span>
                                                ) : null}
                                                {item.actionDueDate ? (
                                                  <span className="text-xs text-amber-700">Due: {formatDate(item.actionDueDate.split('T')[0])}</span>
                                                ) : null}
                                                {!item.actionPriority && !item.actionStatus && !item.actionAssignedTo && !item.actionDueDate && (
                                                  <span className="text-xs text-amber-600 italic">No action details yet</span>
                                                )}
                                              </div>
                                              <button
                                                onClick={() => setExpandedActionEditors(prev => {
                                                  const next = new Set(prev);
                                                  next.has(item.id) ? next.delete(item.id) : next.add(item.id);
                                                  return next;
                                                })}
                                                className="flex-shrink-0 inline-flex items-center gap-1 px-2 py-1 text-xs text-amber-800 bg-white/80 hover:bg-white border border-amber-300 rounded-md transition-colors"
                                              >
                                                <i className={`fas fa-${expandedActionEditors.has(item.id) ? 'chevron-up' : 'pencil-alt'} text-[10px]`}></i>
                                                <span>{expandedActionEditors.has(item.id) ? 'Collapse' : 'Edit'}</span>
                                              </button>
                                            </div>

                                            {/* Action Notes - visible when collapsed */}
                                            {!expandedActionEditors.has(item.id) && (
                                            <div>
                                              <div className="flex items-center justify-between">
                                                <label className="text-xs font-medium text-amber-800 uppercase tracking-wide flex items-center gap-1">
                                                  Action Notes
                                                  <i className="fas fa-robot text-green-500 text-[10px] opacity-70"></i>
                                                </label>
                                                <div className="flex items-center gap-2">
                                                  {localActionEdits[item.id] && (
                                                    <span className="text-xs text-amber-600 italic flex items-center gap-1">
                                                      <i className="fas fa-pencil-alt text-[10px]"></i>Editing...
                                                    </span>
                                                  )}
                                                  {recentlySavedActions.has(item.id) && !localActionEdits[item.id] && (
                                                    <span className="text-xs text-green-600 flex items-center gap-1">
                                                      <i className="fas fa-check text-[10px]"></i>Saved
                                                    </span>
                                                  )}
                                                </div>
                                              </div>
                                              <textarea
                                                value={getActionFieldValue(item, 'actionNotes')}
                                                onChange={(e) => handleActionTextChange(item, 'actionNotes', e.target.value)}
                                                placeholder="Add action notes..."
                                                rows={2}
                                                className="w-full mt-1 text-xs border border-amber-300 rounded px-2 py-1.5 bg-white/95 hover:bg-white focus:border-amber-500 transition-colors resize-none"
                                                data-testid={`textarea-action-notes-${item.id}`}
                                              />
                                              <div className="flex items-center justify-end mt-1">
                                                <button
                                                  type="button"
                                                  onClick={async () => {
                                                    const currentNotes = getActionFieldValue(item, 'actionNotes');
                                                    if (!currentNotes.trim()) return;
                                                    setIsEnhancingActionNotes(item.id);
                                                    try {
                                                      const response = await fetch('/api/ai-enhance-notes', {
                                                        method: 'POST',
                                                        headers: {
                                                          'Content-Type': 'application/json',
                                                          'Authorization': `Bearer ${await authService.getAccessToken()}`
                                                        },
                                                        body: JSON.stringify({ content: currentNotes, itemType: 'Action Notes' })
                                                      });
                                                      if (response.ok) {
                                                        const data = await response.json();
                                                        setLocalActionEdits(prev => ({ ...prev, [item.id]: { ...prev[item.id], actionNotes: data.enhancedContent } }));
                                                        handleActionTextChange(item, 'actionNotes', data.enhancedContent);
                                                      }
                                                    } catch (error) {
                                                      console.error('AI enhancement failed:', error);
                                                    } finally {
                                                      setIsEnhancingActionNotes('');
                                                    }
                                                  }}
                                                  disabled={!getActionFieldValue(item, 'actionNotes').trim() || isEnhancingActionNotes === item.id}
                                                  className="text-xs bg-green-600 hover:bg-green-700 disabled:bg-gray-300 text-white px-2 py-1 rounded flex items-center gap-1"
                                                >
                                                  {isEnhancingActionNotes === item.id ? (
                                                    <><i className="fas fa-spinner fa-spin text-[10px]"></i>Enhancing...</>
                                                  ) : (
                                                    <><i className="fas fa-wand-magic-sparkles text-[10px]"></i>Finish Text</>
                                                  )}
                                                </button>
                                              </div>
                                            </div>
                                            )}

                                            {/* Expandable fields */}
                                            {expandedActionEditors.has(item.id) && (
                                              <div className="space-y-2 border-t border-amber-300 pt-2">
                                              {/* Row 1: Priority + Due Date */}
                                              <div className="grid grid-cols-2 gap-2">
                                                <div>
                                                  <label className="text-xs font-medium text-amber-800 uppercase tracking-wide">Priority</label>
                                                  <select
                                                    value={item.actionPriority || ''}
                                                    onChange={(e) => {
                                                      const priority = e.target.value;
                                                      let dueDate = '';
                                                      if (priority) {
                                                        const today = new Date();
                                                        const daysToAdd = priority === 'High' ? 7 : priority === 'Medium' ? 14 : 30;
                                                        today.setDate(today.getDate() + daysToAdd);
                                                        dueDate = today.toISOString().split('T')[0];
                                                      }
                                                      updateActionFields(item, { actionPriority: priority, actionDueDate: dueDate });
                                                    }}
                                                    className="w-full mt-1 text-xs border border-amber-300 rounded px-2 py-1.5 bg-white/95 hover:bg-white focus:border-amber-500 transition-colors"
                                                    data-testid={`select-action-priority-${item.id}`}
                                                  >
                                                    <option value="">Select...</option>
                                                    <option value="High">High (7 days)</option>
                                                    <option value="Medium">Medium (14 days)</option>
                                                    <option value="Low">Low (30 days)</option>
                                                  </select>
                                                </div>
                                                <div>
                                                  <label className="text-xs font-medium text-amber-800 uppercase tracking-wide">Due Date</label>
                                                  <input
                                                    type="date"
                                                    value={item.actionDueDate ? item.actionDueDate.split('T')[0] : ''}
                                                    onChange={(e) => updateActionFields(item, { actionDueDate: e.target.value })}
                                                    className="w-full mt-1 text-xs border border-amber-300 rounded px-2 py-1.5 bg-white/95 hover:bg-white focus:border-amber-500 transition-colors"
                                                    data-testid={`input-action-due-date-${item.id}`}
                                                  />
                                                </div>
                                              </div>
                                              
                                              {/* Row 2: Status + Assigned To */}
                                              <div className="grid grid-cols-2 gap-2">
                                                <div>
                                                  <label className="text-xs font-medium text-amber-800 uppercase tracking-wide">Status</label>
                                                  <select
                                                    value={item.actionStatus || ''}
                                                    onChange={(e) => updateActionFields(item, { actionStatus: e.target.value })}
                                                    className="w-full mt-1 text-xs border border-amber-300 rounded px-2 py-1.5 bg-white/95 hover:bg-white focus:border-amber-500 transition-colors"
                                                    data-testid={`select-action-status-${item.id}`}
                                                  >
                                                    <option value="">Not Started</option>
                                                    <option value="In Progress">In Progress</option>
                                                    <option value="On Hold">On Hold</option>
                                                    <option value="Completed">Completed</option>
                                                    <option value="Ready to Close">Ready to Close</option>
                                                  </select>
                                                </div>
                                                <div>
                                                  <label className="text-xs font-medium text-amber-800 uppercase tracking-wide">Assigned To</label>
                                                  <input
                                                    type="text"
                                                    value={getActionFieldValue(item, 'actionAssignedTo')}
                                                    onChange={(e) => handleActionTextChange(item, 'actionAssignedTo', e.target.value)}
                                                    placeholder="Enter name..."
                                                    className="w-full mt-1 text-xs border border-amber-300 rounded px-2 py-1.5 bg-white/95 hover:bg-white focus:border-amber-500 transition-colors"
                                                    data-testid={`input-action-assigned-${item.id}`}
                                                  />
                                                </div>
                                              </div>
                                              
                                            </div>
                                            )}

                                            {/* Meeting Discussion - visible when collapsed */}
                                            {!expandedActionEditors.has(item.id) && item.meetingNotes && item.meetingNotes.trim() && (
                                              <div className="bg-white/60 rounded-lg border border-amber-200 p-2">
                                                <div className="flex items-start gap-2">
                                                  <i className="fas fa-comment-dots text-amber-600 text-xs mt-0.5"></i>
                                                  <div className="flex-1 min-w-0">
                                                    <span className="text-xs font-medium text-amber-700">Meeting Discussion: </span>
                                                    <span className="text-xs text-amber-800">{item.meetingNotes}</span>
                                                  </div>
                                                </div>
                                              </div>
                                            )}
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            </CollapsibleContent>
                          </Collapsible>
                        </div>
                      );
                    })}
                  </div>

                  {/* Completed Actions - Group Review container */}
                  {(() => {
                    const allMeetingItems = Object.values(categorized).flat() as MeetingItem[];
                    const readyToCloseItems = allMeetingItems.filter(item => item.actionStatus === 'Ready to Close');
                    if (readyToCloseItems.length === 0) return null;
                    const containerKey = `${meetingDate}-ready-to-close`;
                    const isExpanded = expandedCategories.has(containerKey);
                    return (
                      <div className="mt-2">
                        <Collapsible open={isExpanded} onOpenChange={(open) => {
                          if (open) {
                            setExpandedCategories(prev => new Set([...prev, containerKey]));
                          } else {
                            setExpandedCategories(prev => { const n = new Set(prev); n.delete(containerKey); return n; });
                          }
                        }}>
                          <CollapsibleTrigger asChild>
                            <Button
                              variant="ghost"
                              className="w-full justify-between p-3 sm:p-4 h-auto bg-emerald-50 text-emerald-900 border border-emerald-300 rounded-lg hover:bg-emerald-100 mb-2 touch-manipulation min-h-[48px] sm:min-h-[44px]"
                            >
                              <div className="flex items-center gap-2 sm:gap-3 min-w-0 flex-1">
                                <CheckCircle className="h-4 w-4 sm:h-5 sm:w-5 flex-shrink-0 text-emerald-600" />
                                <div className="text-left min-w-0 flex-1">
                                  <div className="font-semibold text-sm sm:text-base leading-tight">Completed Actions – Group Review</div>
                                  <div className="text-xs opacity-75 mt-0.5">{readyToCloseItems.length} item{readyToCloseItems.length !== 1 ? 's' : ''} ready for sign-off</div>
                                </div>
                              </div>
                              {isExpanded ? <ChevronDown className="h-4 w-4 flex-shrink-0" /> : <ChevronRight className="h-4 w-4 flex-shrink-0" />}
                            </Button>
                          </CollapsibleTrigger>
                          <CollapsibleContent>
                            <div className="space-y-3 pl-1 sm:pl-2">
                              {readyToCloseItems.map((item: MeetingItem) => (
                                <div key={`rtc-${item.id}`} className="bg-white rounded-lg border border-emerald-200 shadow-sm p-3 sm:p-4">
                                  {/* Header row */}
                                  <div className="flex items-start justify-between gap-2 mb-3">
                                    <h4 className="font-semibold text-gray-900 text-sm sm:text-base leading-tight flex-1 min-w-0 pr-2 line-clamp-2">
                                      {item.title?.trim() || <span className="text-gray-400 italic">No title</span>}
                                    </h4>
                                    <div className="flex flex-wrap gap-1.5 flex-shrink-0">
                                      <Badge variant="outline" className={`text-xs whitespace-nowrap ${
                                        item.type === 'Safety Ideas' ? 'bg-red-50 text-red-700 border-red-200' :
                                        item.type === 'Near Miss' ? 'bg-orange-50 text-orange-700 border-orange-200' :
                                        'bg-blue-50 text-blue-700 border-blue-200'
                                      }`}>
                                        {item.type}
                                      </Badge>
                                      <Badge className="text-xs whitespace-nowrap bg-emerald-100 text-emerald-800 border-emerald-300">
                                        Ready to Close
                                      </Badge>
                                    </div>
                                  </div>
                                  {/* Detail rows */}
                                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-xs text-gray-600 mb-3">
                                    {item.actionAssignedTo && (
                                      <div className="flex items-center gap-1.5">
                                        <i className="fas fa-user text-emerald-600 w-3"></i>
                                        <span><span className="font-medium">Actioned by:</span> {item.actionAssignedTo}</span>
                                      </div>
                                    )}
                                    {item.actionDueDate && (
                                      <div className="flex items-center gap-1.5">
                                        <i className="fas fa-calendar-check text-emerald-600 w-3"></i>
                                        <span><span className="font-medium">Due:</span> {formatDate(item.actionDueDate.split('T')[0])}</span>
                                      </div>
                                    )}
                                    {item.submittedBy && (
                                      <div className="flex items-center gap-1.5">
                                        <i className="fas fa-user-edit text-gray-400 w-3"></i>
                                        <span><span className="font-medium">Submitted by:</span> {item.submittedBy}</span>
                                      </div>
                                    )}
                                    {item.status && (
                                      <div className="flex items-center gap-1.5">
                                        <i className="fas fa-tag text-gray-400 w-3"></i>
                                        <span><span className="font-medium">Item status:</span> {item.status}</span>
                                      </div>
                                    )}
                                  </div>
                                  {/* Action notes */}
                                  {item.actionNotes && (
                                    <div className="bg-emerald-50 border border-emerald-100 rounded p-2 mb-2">
                                      <div className="flex items-start gap-2">
                                        <i className="fas fa-clipboard-check text-emerald-600 text-xs mt-0.5"></i>
                                        <div>
                                          <span className="text-xs font-medium text-emerald-800">What was done: </span>
                                          <span className="text-xs text-emerald-900">{item.actionNotes}</span>
                                        </div>
                                      </div>
                                    </div>
                                  )}
                                  {/* Meeting notes */}
                                  {item.meetingNotes && (
                                    <div className="bg-gray-50 border border-gray-100 rounded p-2">
                                      <div className="flex items-start gap-2">
                                        <i className="fas fa-comment-dots text-gray-400 text-xs mt-0.5"></i>
                                        <div>
                                          <span className="text-xs font-medium text-gray-600">Discussion notes: </span>
                                          <span className="text-xs text-gray-700">{item.meetingNotes}</span>
                                        </div>
                                      </div>
                                    </div>
                                  )}
                                  {/* Call to action */}
                                  <div className="mt-3 pt-2 border-t border-emerald-100 flex items-center gap-2 text-xs text-emerald-700">
                                    <i className="fas fa-users text-emerald-500"></i>
                                    <span>Group to discuss, confirm outcome and close off this item.</span>
                                  </div>
                                </div>
                              ))}
                            </div>
                          </CollapsibleContent>
                        </Collapsible>
                      </div>
                    );
                  })()}

                    {/* Print-only signature section */}
                    <div className="print-signatures" style={{display: 'none'}}>
                  <h3><strong>Meeting Approval & Signatures</strong></h3>
                  <div className="print-signature-grid">
                    <div>
                      <div className="print-signature-label">Committee Chair:</div>
                      <div className="print-signature-line"></div>
                      <div style={{fontSize: '8px', marginBottom: '15px'}}>Signature & Date</div>
                      
                      <div className="print-signature-label">Health & Safety Manager:</div>
                      <div className="print-signature-line"></div>
                      <div style={{fontSize: '8px'}}>Signature & Date</div>
                    </div>
                    <div>
                      <div className="print-signature-label">Secretary:</div>
                      <div className="print-signature-line"></div>
                      <div style={{fontSize: '8px', marginBottom: '15px'}}>Signature & Date</div>
                      
                      <div className="print-signature-label">General Manager:</div>
                      <div className="print-signature-line"></div>
                      <div style={{fontSize: '8px'}}>Signature & Date</div>
                    </div>
                  </div>
                </div>
                  </CardContent>
                )}
              </Card>
            ))}
              </div>
            )}
          </div>
        </div>

        {/* Print-only footer */}
        <div className="print-footer" style={{display: 'none'}}>
          Generated: {new Date().toLocaleDateString('en-GB')} | Cranfield Glass Christchurch HSEQ & Compliance System | Document ID: CG-HS-{new Date().toISOString().split('T')[0].replace(/-/g, '')}
        </div>

        {/* Compliance Footer */}
        <div className="mt-6 sm:mt-8 bg-gradient-to-r from-gray-800 to-gray-900 text-white rounded-xl shadow-lg p-4 sm:p-6">
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between space-y-3 sm:space-y-0">
            <div className="text-center sm:text-left">
              <h3 className="text-base sm:text-lg font-bold">Cranfield Glass Christchurch</h3>
              <p className="text-xs sm:text-sm text-gray-300 mt-1">Health & Safety Compliance Records</p>
            </div>
            <div className="flex flex-col sm:flex-row sm:items-center space-y-1 sm:space-y-0 sm:space-x-6 text-center sm:text-right">
              <div className="flex items-center justify-center sm:justify-end space-x-2">
                <i className="fas fa-calendar-alt text-gray-400 text-xs"></i>
                <p className="text-xs sm:text-sm text-gray-300">Generated: {new Date().toLocaleDateString('en-GB')}</p>
              </div>
              <div className="flex items-center justify-center sm:justify-end space-x-2">
                <i className="fas fa-file-alt text-gray-400 text-xs"></i>
                <p className="text-xs sm:text-sm text-gray-300">
                  <span className="font-medium">{filteredItems.length}</span> {filteredItems.length === 1 ? 'Record' : 'Records'}
                </p>
              </div>
            </div>
          </div>
        </div>
        
        {/* AI Generation Modal */}
        <Dialog open={isGeneratingTitles} onOpenChange={() => {}}>
          <DialogContent className="sm:max-w-md">
            <DialogHeader>
              <DialogTitle className="flex items-center gap-2">
                <Bot className="h-5 w-5 text-blue-600" />
                Generating AI Titles
              </DialogTitle>
            </DialogHeader>
            <div className="space-y-4">
              <div className="flex items-center gap-3">
                <Loader2 className="h-4 w-4 animate-spin text-blue-600" />
                <span className="text-sm text-gray-600">
                  {aiGenerationProgress.currentItem || 'Processing items with OpenAI GPT-4o...'}
                </span>
              </div>
              
              {aiGenerationProgress.total > 0 && (
                <div className="space-y-2">
                  <div className="flex justify-between text-sm">
                    <span>Progress</span>
                    <span>{aiGenerationProgress.current} of {aiGenerationProgress.total}</span>
                  </div>
                  <Progress 
                    value={(aiGenerationProgress.current / aiGenerationProgress.total) * 100} 
                    className="h-2"
                  />
                </div>
              )}
              
              <div className="text-xs text-gray-500">
                This process uses OpenAI GPT-4o to generate intelligent titles from content. 
                Please wait while we enhance your meeting records.
              </div>
            </div>
          </DialogContent>
        </Dialog>

        {/* Reschedule Meeting Modal */}
        <Dialog open={showRescheduleMeetingModal} onOpenChange={setShowRescheduleMeetingModal}>
          <DialogContent className="sm:max-w-md">
            <DialogHeader>
              <DialogTitle className="flex items-center gap-2">
                <Calendar className="h-5 w-5 text-blue-600" />
                Reschedule Meeting
              </DialogTitle>
            </DialogHeader>
            <div className="space-y-4">
              <div className="text-sm text-gray-600">
                Change the date for <span className="font-medium">{rescheduleMeetingSourceDate ? formatDate(rescheduleMeetingSourceDate) : ''}</span>.
                All items on this meeting date will be updated in SharePoint.
              </div>

              <div className="space-y-2">
                <label className="text-sm font-medium text-gray-700 flex items-center gap-2">
                  <Calendar className="h-4 w-4 text-blue-600" />
                  New Meeting Date
                </label>
                <input
                  type="date"
                  value={rescheduleTargetDate}
                  onChange={(e) => setRescheduleTargetDate(e.target.value)}
                  className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                />
                {rescheduleTargetDate && (
                  <div className="text-xs text-gray-500">
                    New date: <strong>{new Date(rescheduleTargetDate).toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</strong>
                  </div>
                )}
              </div>

              <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
                <div className="text-xs text-amber-800">
                  <strong><i className="fas fa-exclamation-triangle mr-1"></i>This will update all items on this meeting date in SharePoint, including closed items.</strong>
                </div>
              </div>

              <div className="flex gap-3 pt-2">
                <Button
                  variant="outline"
                  onClick={() => setShowRescheduleMeetingModal(false)}
                  className="flex-1"
                >
                  Cancel
                </Button>
                <Button
                  onClick={executeRescheduleMeeting}
                  disabled={!rescheduleTargetDate || rescheduleTargetDate === new Date(rescheduleMeetingSourceDate || '').toISOString().split('T')[0]}
                  className="flex-1 bg-blue-600 hover:bg-blue-700 text-white disabled:bg-gray-300"
                >
                  <Calendar className="h-4 w-4 mr-2" />
                  Reschedule Meeting
                </Button>
              </div>
            </div>
          </DialogContent>
        </Dialog>

        {/* Rescheduling Progress Overlay */}
        <Dialog open={isRescheduling} onOpenChange={() => {}}>
          <DialogContent className="sm:max-w-sm">
            <DialogHeader>
              <DialogTitle className="flex items-center gap-2">
                <Loader2 className="h-5 w-5 animate-spin text-blue-600" />
                Rescheduling Meeting...
              </DialogTitle>
            </DialogHeader>
            <div className="text-sm text-gray-600">
              Updating all items in SharePoint. Please wait.
            </div>
          </DialogContent>
        </Dialog>

        {/* Move Agenda Items Confirmation Modal */}
        <Dialog open={showMergeConfirmation} onOpenChange={setShowMergeConfirmation}>
          <DialogContent className="sm:max-w-md">
            <DialogHeader>
              <DialogTitle className="flex items-center gap-2">
                <Calendar className="h-5 w-5 text-blue-600" />
                Move to Next Meeting
              </DialogTitle>
            </DialogHeader>
            <div className="space-y-4">
              {mergeConfirmationData && (
                <>
                  <div className="text-sm text-gray-600">
                    Move <span className="font-semibold text-blue-600">{mergeConfirmationData.movableItemCount}</span> <span className="font-medium">Submitted</span> item{mergeConfirmationData.movableItemCount !== 1 ? 's' : ''} from{' '}
                    <span className="font-medium">{formatDate(mergeConfirmationData.currentMeetingDate)}</span> to the next scheduled meeting.
                    {mergeConfirmationData.movableItemCount < mergeConfirmationData.itemCount && (
                      <span className="text-gray-500 italic"> ({mergeConfirmationData.itemCount - mergeConfirmationData.movableItemCount} item{mergeConfirmationData.itemCount - mergeConfirmationData.movableItemCount !== 1 ? 's' : ''} with other statuses will remain)</span>
                    )}
                  </div>

                  <div className="bg-green-50 border border-green-200 rounded-lg p-3 flex items-center gap-3">
                    <div className="w-8 h-8 bg-green-600 rounded-full flex items-center justify-center flex-shrink-0">
                      <Calendar className="h-4 w-4 text-white" />
                    </div>
                    <div>
                      <div className="text-xs text-green-700 font-medium uppercase tracking-wide">Next Meeting</div>
                      <div className="text-sm font-semibold text-green-900">{mergeConfirmationData.formattedNextDate}</div>
                    </div>
                  </div>
                  
                  <div className="bg-blue-50 border border-blue-200 rounded-lg p-3">
                    <div className="text-xs text-blue-800">
                      <i className="fas fa-info-circle mr-1"></i>
                      Only <strong>Submitted</strong> items will be moved. Items with any other status (In Discussion, Actioned, Closed) stay on the original date.
                    </div>
                  </div>
                  
                  <div className="flex gap-3 pt-2">
                    <Button 
                      variant="outline" 
                      onClick={() => setShowMergeConfirmation(false)}
                      className="flex-1"
                    >
                      Cancel
                    </Button>
                    <Button 
                      onClick={executeMerge}
                      className="flex-1 bg-blue-600 hover:bg-blue-700 text-white"
                    >
                      <ArrowRight className="h-4 w-4 mr-2" />
                      Move {mergeConfirmationData.movableItemCount} Item{mergeConfirmationData.movableItemCount !== 1 ? 's' : ''}
                    </Button>
                  </div>
                </>
              )}
            </div>
          </DialogContent>
        </Dialog>

        {/* Moving Agenda Items Progress Modal */}
        <Dialog open={isMergingMeeting} onOpenChange={() => {}}>
          <DialogContent className="sm:max-w-md">
            <DialogHeader>
              <DialogTitle className="flex items-center gap-2">
                <Calendar className="h-5 w-5 text-blue-600" />
                Moving Agenda Items
              </DialogTitle>
            </DialogHeader>
            <div className="space-y-4">
              <div className="flex items-center gap-3">
                <Loader2 className="h-4 w-4 animate-spin text-blue-600" />
                <span className="text-sm text-gray-600">
                  {mergeProgress.currentItem || 'Updating meeting dates in SharePoint...'}
                </span>
              </div>
              
              {mergeProgress.total > 0 && (
                <div className="space-y-2">
                  <div className="flex justify-between text-sm">
                    <span>Progress</span>
                    <span>{mergeProgress.current} of {mergeProgress.total}</span>
                  </div>
                  <Progress 
                    value={(mergeProgress.current / mergeProgress.total) * 100} 
                    className="h-2"
                  />
                </div>
              )}
              
              <div className="text-xs text-gray-500">
                Moving agenda items to the selected meeting date. 
                This will update meeting dates in SharePoint for Business Ideas, Safety Ideas, and Near Miss items.
              </div>
            </div>
          </DialogContent>
        </Dialog>

        {/* Mobile & Print Styles */}
        <style>{`
          /* Flip Card Styles */
          .flip-card-container {
            perspective: 1000px;
          }
          
          .flip-card-inner {
            position: relative;
            width: 100%;
            transition: transform 0.6s cubic-bezier(0.4, 0, 0.2, 1);
            transform-style: preserve-3d;
          }
          
          .flip-card-front,
          .flip-card-back {
            -webkit-backface-visibility: hidden;
            backface-visibility: hidden;
          }
          
          .flip-card-back {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            transform: rotateY(180deg);
          }
          
          /* Ensure proper height calculation for flip cards */
          .flip-card-container {
            min-height: auto;
          }
          
          .flip-card-inner.flipped {
            transform: rotateY(180deg);
          }
          
          /* Touch-friendly mobile enhancements */
          .touch-manipulation {
            touch-action: manipulation;
            -webkit-tap-highlight-color: rgba(0, 0, 0, 0.1);
          }
          
          /* Custom scrollbar for mobile */
          ::-webkit-scrollbar {
            width: 4px;
          }
          
          ::-webkit-scrollbar-track {
            background: #f1f1f1;
            border-radius: 2px;
          }
          
          ::-webkit-scrollbar-thumb {
            background: #c1c1c1;
            border-radius: 2px;
          }
          
          ::-webkit-scrollbar-thumb:hover {
            background: #a8a8a8;
          }
          
          /* Professional PDF Print Styles */
          @media print {
            @page {
              margin: 15mm;
              size: A4;
            }
            
            * {
              -webkit-print-color-adjust: exact !important;
              color-adjust: exact !important;
            }
            
            body { 
              background: white !important; 
              font-family: 'Arial', sans-serif !important;
              font-size: 11px !important;
              line-height: 1.3 !important;
              color: #000 !important;
            }
            
            /* Hide UI elements for print */
            .no-print,
            button,
            .bg-gray-50:first-child,
            nav,
            header {
              display: none !important; 
            }
            
            /* Show print-only elements */
            .print-header,
            .print-signatures,
            .print-footer {
              display: block !important;
            }
            
            /* Meeting header styling */
            .print-header {
              text-align: center !important;
              border-bottom: 2px solid #000 !important;
              padding-bottom: 10px !important;
              margin-bottom: 20px !important;
            }
            
            .print-title {
              font-size: 18px !important;
              font-weight: bold !important;
              margin-bottom: 5px !important;
            }
            
            .print-subtitle {
              font-size: 14px !important;
              font-weight: bold !important;
              margin-bottom: 5px !important;
            }
            
            .print-date {
              font-size: 12px !important;
              margin-bottom: 15px !important;
            }
            
            /* Meeting details in columns */
            .print-meeting-details {
              display: grid !important;
              grid-template-columns: 1fr 1fr !important;
              gap: 20px !important;
              margin-bottom: 20px !important;
              padding: 10px !important;
              border: 1px solid #ccc !important;
            }
            
            /* Category sections */
            .print-category {
              page-break-inside: avoid !important;
              margin-bottom: 20px !important;
            }
            
            .print-category-header {
              background: #f0f0f0 !important;
              padding: 8px !important;
              border: 1px solid #000 !important;
              font-weight: bold !important;
              margin-bottom: 10px !important;
            }
            
            /* Items in table format */
            .print-items-table {
              width: 100% !important;
              border-collapse: collapse !important;
              margin-bottom: 15px !important;
            }
            
            .print-items-table th {
              background: #e0e0e0 !important;
              border: 1px solid #000 !important;
              padding: 5px !important;
              font-size: 10px !important;
              font-weight: bold !important;
              text-align: left !important;
            }
            
            .print-items-table td {
              border: 1px solid #666 !important;
              padding: 4px !important;
              font-size: 9px !important;
              vertical-align: top !important;
            }
            
            /* Signature section */
            .print-signatures {
              page-break-inside: avoid !important;
              margin-top: 30px !important;
              border-top: 1px solid #000 !important;
              padding-top: 15px !important;
            }
            
            .print-signature-grid {
              display: grid !important;
              grid-template-columns: 1fr 1fr !important;
              gap: 30px !important;
            }
            
            .print-signature-line {
              border-bottom: 1px solid #000 !important;
              height: 25px !important;
              margin-bottom: 5px !important;
            }
            
            .print-signature-label {
              font-size: 10px !important;
              font-weight: bold !important;
            }
            
            /* Document footer */
            .print-footer {
              position: fixed !important;
              bottom: 10mm !important;
              left: 0 !important;
              right: 0 !important;
              text-align: center !important;
              font-size: 8px !important;
              border-top: 1px solid #ccc !important;
              padding-top: 5px !important;
            }
            
            /* Transform existing content for print */
            .bg-gradient-to-r {
              background: #f0f0f0 !important;
              border: 1px solid #000 !important;
              color: #000 !important;
            }
            
            .bg-yellow-50,
            .bg-green-50,
            .bg-red-50,
            .bg-blue-50,
            .bg-orange-50 {
              background: #f8f8f8 !important;
              border: 1px solid #ccc !important;
            }
            
            .text-white {
              color: #000 !important;
            }
            
            .shadow-lg {
              box-shadow: none !important;
            }
            
            .rounded-xl,
            .rounded-lg {
              border-radius: 0 !important;
            }
            
            /* Hide collapsible controls */
            [data-state] {
              display: block !important;
            }
            
            /* Ensure all content is visible */
            .space-y-4 > * {
              display: block !important;
            }
          }
        `}</style>

        {/* Error Modal */}
        <Dialog open={showErrorModal} onOpenChange={setShowErrorModal}>
          <DialogContent className="mx-4 max-w-sm rounded-2xl border-0 shadow-2xl" aria-describedby="error-description">
            <DialogHeader>
              <div className="flex flex-col items-center text-center gap-3 pt-2">
                <div className="h-12 w-12 rounded-full bg-red-100 flex items-center justify-center">
                  <CalendarX className="h-6 w-6 text-red-600" />
                </div>
                <DialogTitle className="text-lg font-semibold text-gray-900">
                  {modalTitle}
                </DialogTitle>
                <div id="error-description" className="sr-only">Error notification dialog</div>
              </div>
            </DialogHeader>
            <div className="space-y-4 pb-2">
              <p className="text-sm text-gray-600 text-center">{modalMessage}</p>
              <Button 
                onClick={() => setShowErrorModal(false)}
                className="w-full bg-red-600 hover:bg-red-700 text-white rounded-xl h-11"
              >
                OK
              </Button>
            </div>
          </DialogContent>
        </Dialog>

        {/* Success Modal */}
        <Dialog open={showSuccessModal} onOpenChange={setShowSuccessModal}>
          <DialogContent className="mx-4 max-w-sm rounded-2xl border-0 shadow-2xl" aria-describedby="success-description">
            <DialogHeader>
              <div className="flex flex-col items-center text-center gap-3 pt-2">
                <div className="h-12 w-12 rounded-full bg-green-100 flex items-center justify-center">
                  <CheckCircle2 className="h-6 w-6 text-green-600" />
                </div>
                <DialogTitle className="text-lg font-semibold text-gray-900">
                  {modalTitle}
                </DialogTitle>
                <div id="success-description" className="sr-only">Success notification dialog</div>
              </div>
            </DialogHeader>
            <div className="space-y-4 pb-2">
              <p className="text-sm text-gray-600 text-center">{modalMessage}</p>
              <Button 
                onClick={() => setShowSuccessModal(false)}
                className="w-full bg-green-600 hover:bg-green-700 text-white rounded-xl h-11"
              >
                Done
              </Button>
            </div>
          </DialogContent>
        </Dialog>

        {/* Shareable URL Success Modal */}
        <Dialog open={showShareableModal} onOpenChange={setShowShareableModal}>
          <DialogContent className="sm:max-w-lg" aria-describedby="shareable-description">
            <DialogHeader>
              <DialogTitle className="flex items-center gap-2 text-blue-600">
                <i className="fas fa-share-alt"></i>
                HTML Export Complete
              </DialogTitle>
              <div id="shareable-description" className="sr-only">Shareable URL success dialog</div>
            </DialogHeader>
            <div className="space-y-4">
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
                <div className="flex items-center gap-2 mb-2">
                  <i className="fas fa-download text-blue-600"></i>
                  <span className="font-medium text-blue-900">File Downloaded</span>
                </div>
                <p className="text-sm text-blue-800">{shareableFilename}</p>
              </div>
              
              <div className="bg-green-50 border border-green-200 rounded-lg p-4">
                <div className="flex items-center gap-2 mb-2">
                  <i className="fas fa-link text-green-600"></i>
                  <span className="font-medium text-green-900">Shareable URL Created</span>
                </div>
                <p className="text-sm text-green-800 mb-3">
                  Share this link with others to view the meeting minutes online:
                </p>
                <div className="bg-white border border-green-300 rounded-md p-3 flex items-center gap-2">
                  <input
                    type="text"
                    value={shareableUrl}
                    readOnly
                    className="flex-1 text-sm text-gray-700 bg-transparent border-none outline-none"
                  />
                  <button
                    onClick={(e) => {
                      navigator.clipboard.writeText(shareableUrl);
                      // Briefly show copied feedback
                      const btn = e.currentTarget;
                      if (btn) {
                        const originalText = btn.innerHTML;
                        btn.innerHTML = '<i class="fas fa-check text-green-600"></i>';
                        setTimeout(() => btn.innerHTML = originalText, 1000);
                      }
                    }}
                    className="text-sm bg-green-600 hover:bg-green-700 text-white px-3 py-1 rounded flex items-center gap-1 transition-colors"
                    title="Copy to clipboard"
                  >
                    <i className="fas fa-copy"></i>
                    Copy
                  </button>
                </div>
                <p className="text-xs text-green-600 mt-2">
                  <i className="fas fa-info-circle"></i>
                  Link expires after 7 days • Protected from search engines
                </p>
              </div>
              
              <div className="flex justify-center pt-2">
                <Button 
                  onClick={() => setShowShareableModal(false)}
                  className="px-8 bg-blue-600 hover:bg-blue-700 text-white"
                >
                  Close
                </Button>
              </div>
            </div>
          </DialogContent>
        </Dialog>

        {/* Add Idea Modal - Mobile Optimized */}
        <Dialog open={showAddIdeaModal} onOpenChange={setShowAddIdeaModal}>
          <DialogContent className="w-[95vw] max-w-md mx-auto my-4 max-h-[95vh] overflow-y-auto" aria-describedby="add-idea-description">
            <DialogHeader className="space-y-2 pb-3">
              <DialogTitle className="flex items-center gap-2 text-lg font-semibold">
                <i className="fas fa-lightbulb text-blue-600"></i>
                Add New Idea
              </DialogTitle>
              <div id="add-idea-description" className="text-sm text-gray-600">
                Add a new business idea or safety idea to the meeting agenda
              </div>
              <div className="flex items-center gap-2 px-3 py-2 bg-blue-50 border border-blue-200 rounded-lg text-sm font-medium text-blue-800">
                <i className="fas fa-calendar text-blue-600 flex-shrink-0"></i>
                <span className="truncate">Meeting: {newIdeaForm.meetingDate ? formatDisplayDate(newIdeaForm.meetingDate, 'date-only') : 'No date selected'}</span>
              </div>
            </DialogHeader>
            <div className="space-y-4">
              <div>
                <label className="text-sm font-medium text-gray-700 mb-2 block">
                  Idea Type
                </label>
                <select
                  value={newIdeaForm.type}
                  onChange={(e) => handleFormTypeChange(e.target.value as 'Business Ideas' | 'Safety Ideas')}
                  className="w-full border border-gray-300 rounded-lg px-3 py-3 sm:py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 touch-manipulation"
                  style={{ fontSize: '16px' }} // Prevents zoom on iOS
                >
                  <option value="">Select type...</option>
                  <option value="Safety Ideas">Safety Ideas</option>
                  <option value="Business Ideas">Business Ideas</option>
                </select>
              </div>
              
              <div>
                <label className="text-sm font-medium text-gray-700 mb-2 block">
                  Idea Type
                </label>
                <select
                  value={newIdeaForm.ideaType}
                  onChange={(e) => setNewIdeaForm(prev => ({ ...prev, ideaType: e.target.value }))}
                  className="w-full border border-gray-300 rounded-lg px-3 py-3 sm:py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 touch-manipulation"
                  style={{ fontSize: '16px' }} // Prevents zoom on iOS
                >
                  <option value="">Select idea type...</option>
                  {ideaTypeOptions.map(option => (
                    <option key={option} value={option}>{option}</option>
                  ))}
                </select>
              </div>
              
              <div>
                <label className="text-sm font-medium text-gray-700 mb-2 flex items-center gap-2">
                  Description
                  <i className="fas fa-robot text-green-500 text-xs opacity-70"></i>
                </label>
                <InlineTextarea
                  value={newIdeaForm.description}
                  onChange={(value) => setNewIdeaForm(prev => ({ ...prev, description: value }))}
                  placeholder="Describe your idea in detail... Type to see inline completions."
                  rows={4}
                  context={newIdeaForm.type === 'Safety Ideas' ? 'safety' : 'business'}
                  className="w-full text-sm border border-gray-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                />
                <div className="flex items-center justify-between mt-2">
                  <div className="flex items-center gap-2 text-xs text-gray-500">
                    <i className="fas fa-keyboard text-blue-500"></i>
                    <span>Predictive text + </span>
                    <i className="fas fa-magic text-green-500"></i>
                    <span>AI assistance</span>
                  </div>
                  <button
                    type="button"
                    onClick={async () => {
                      if (!newIdeaForm.description.trim()) return;
                      
                      try {
                        // Call AI to enhance the description
                        const response = await fetch('/api/ai-enhance-notes', {
                          method: 'POST',
                          headers: {
                            'Content-Type': 'application/json',
                            'Authorization': `Bearer ${await authService.getAccessToken()}`
                          },
                          body: JSON.stringify({
                            content: newIdeaForm.description,
                            itemType: newIdeaForm.type || 'Idea'
                          })
                        });
                        
                        if (response.ok) {
                          const data = await response.json();
                          setNewIdeaForm(prev => ({ ...prev, description: data.enhancedContent }));
                        }
                      } catch (error) {
                        console.error('AI enhancement failed:', error);
                      }
                    }}
                    disabled={!newIdeaForm.description.trim()}
                    className="text-xs bg-green-600 hover:bg-green-700 disabled:bg-gray-300 text-white px-2 py-1 rounded flex items-center gap-1"
                  >
                    <i className="fas fa-wand-magic-sparkles"></i>
                    Enhance
                  </button>
                </div>
              </div>
              
              <div>
                <label className="text-sm font-medium text-gray-700 mb-2 block">
                  Submitted By
                </label>
                <select
                  value={newIdeaForm.submittedById}
                  onChange={(e) => {
                    const val = e.target.value;
                    const spUser = sharepointUsers.find(user => user.id.toString() === val);
                    const name = spUser?.title || (val.startsWith('static_') ? val.slice(7) : '');
                    setNewIdeaForm(prev => ({ 
                      ...prev, 
                      submittedById: val,
                      submittedBy: name
                    }));
                  }}
                  className="w-full border border-gray-300 rounded-lg px-3 py-3 sm:py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 touch-manipulation"
                  style={{ fontSize: '16px' }}
                >
                  <option value="">Select person...</option>
                  {(() => {
                    const spNames = new Set(sharepointUsers.map(u => u.title?.toLowerCase().trim()));
                    const staticFallbacks = [
                      ...meetingAttendees.management,
                      ...meetingAttendees.glaziers
                    ].filter(s => !spNames.has(s.name.toLowerCase().trim()));
                    const allSpUsers = sharepointUsers.filter(u => u.title?.trim());
                    const combined = [
                      ...allSpUsers.map(u => ({ value: u.id.toString(), label: u.title })),
                      ...staticFallbacks.map(s => ({ value: `static_${s.name}`, label: s.name }))
                    ].sort((a, b) => a.label.localeCompare(b.label));
                    return combined.map(opt => (
                      <option key={opt.value} value={opt.value}>{opt.label}</option>
                    ));
                  })()}
                </select>
              </div>
              
              <div className="flex flex-col sm:flex-row gap-3 pt-4">
                <Button 
                  onClick={() => {
                    setShowAddIdeaModal(false);
                    setNewIdeaForm({
                      type: 'Business Ideas',
                      description: '',
                      submittedBy: '',
                      submittedById: '',
                      assignedTo: '',
                      ideaType: '',
                      meetingDate: ''
                    });
                    setIdeaTypeOptions([]);
                  }}
                  variant="outline"
                  className="flex-1 min-h-[44px] touch-manipulation"
                  disabled={isCreatingItem}
                >
                  Cancel
                </Button>
                <Button 
                  onClick={() => {
                    if (newIdeaForm.type && newIdeaForm.description && newIdeaForm.submittedBy) {
                      createSharePointItem(newIdeaForm);
                      setShowAddIdeaModal(false);
                      setNewIdeaForm({
                        type: 'Business Ideas',
                        description: '',
                        submittedBy: '',
                        submittedById: '',
                        assignedTo: '',
                        ideaType: '',
                        meetingDate: ''
                      });
                      setIdeaTypeOptions([]);
                    } else {
                      showError('Validation Error', 'Please fill in all required fields');
                    }
                  }}
                  className="flex-1 bg-blue-600 hover:bg-blue-700 text-white min-h-[44px] touch-manipulation"
                  disabled={isCreatingItem || !newIdeaForm.type || !newIdeaForm.description || !newIdeaForm.submittedBy}
                >
                  {isCreatingItem ? (
                    <>
                      <i className="fas fa-spinner fa-spin mr-2"></i>
                      Creating...
                    </>
                  ) : (
                    <>
                      <i className="fas fa-plus mr-2"></i>
                      Create Idea
                    </>
                  )}
                </Button>
              </div>
            </div>
          </DialogContent>
        </Dialog>

        {/* Meeting Notes Modal - Mobile Optimized */}
        <Dialog open={showNotesModal} onOpenChange={(open) => { setShowNotesModal(open); if (!open) setShowSaveDropdown(false); }}>
          <DialogContent className="w-[95vw] max-w-lg mx-auto my-4 max-h-[95vh] overflow-y-auto" aria-describedby="notes-description">
            <DialogHeader className="pb-3">
              <DialogTitle className="flex items-center gap-2 text-lg">
                <i className="fas fa-edit text-blue-600"></i>
                Edit Meeting Notes
              </DialogTitle>
              <div id="notes-description" className="sr-only">Edit meeting notes for agenda item</div>
            </DialogHeader>
            <div className="space-y-4">
              <div>
                <label className="text-sm font-medium text-gray-700 mb-2 flex items-center gap-2">
                  Meeting Notes for: {editingNotesItem?.title}
                  <i className="fas fa-robot text-green-500 text-xs opacity-70"></i>
                </label>
                <InlineTextarea
                  value={editingMeetingNotes}
                  onChange={(value) => setEditingMeetingNotes(value)}
                  placeholder="Add meeting discussion notes... Type to see inline completions."
                  rows={6}
                  context="meeting_notes"
                  className="w-full border border-gray-300 rounded-lg px-3 py-3 sm:py-2 focus:ring-2 focus:ring-blue-500 focus:border-blue-500 touch-manipulation text-base" // text-base prevents zoom on iOS
                />
                <div className="flex items-center justify-between mt-2">
                  <div className="flex items-center gap-2 text-xs text-gray-500">
                    <i className="fas fa-keyboard text-blue-500"></i>
                    <span>Predictive text + </span>
                    <i className="fas fa-magic text-green-500"></i>
                    <span>AI assistance</span>
                  </div>
                  <button
                    type="button"
                    onClick={async () => {
                      if (!editingMeetingNotes.trim() || isEnhancingNotes) return;
                      
                      setIsEnhancingNotes(true);
                      try {
                        // Call AI to enhance the notes
                        const response = await fetch('/api/ai-enhance-notes', {
                          method: 'POST',
                          headers: {
                            'Content-Type': 'application/json',
                            'Authorization': `Bearer ${await authService.getAccessToken()}`
                          },
                          body: JSON.stringify({
                            content: editingMeetingNotes,
                            itemType: editingNotesItem?.type || 'Meeting Notes'
                          })
                        });
                        
                        if (response.ok) {
                          const data = await response.json();
                          setEditingMeetingNotes(data.enhancedContent);
                        } else {
                          console.error('AI enhancement failed:', response.status);
                        }
                      } catch (error) {
                        console.error('AI enhancement failed:', error);
                      } finally {
                        setIsEnhancingNotes(false);
                      }
                    }}
                    disabled={!editingMeetingNotes.trim() || isEnhancingNotes}
                    className="text-xs bg-green-600 hover:bg-green-700 disabled:bg-gray-300 text-white px-2 py-1 rounded flex items-center gap-1"
                  >
                    {isEnhancingNotes ? (
                      <>
                        <i className="fas fa-spinner fa-spin"></i>
                        Enhancing...
                      </>
                    ) : (
                      <>
                        <i className="fas fa-wand-magic-sparkles"></i>
                        Enhance
                      </>
                    )}
                  </button>
                </div>
              </div>
              
              <div className="flex flex-col sm:flex-row gap-3 pt-4">
                <Button 
                  onClick={() => {
                    setShowNotesModal(false);
                    setEditingNotesItem(null);
                    setEditingMeetingNotes('');
                  }}
                  variant="outline"
                  className="flex-1 min-h-[44px] touch-manipulation"
                >
                  Cancel
                </Button>
                <div className="flex-1 flex relative">
                  {/* Main save button */}
                  <Button
                    onClick={() => {
                      if (editingNotesItem) {
                        updateMeetingNotes(editingNotesItem, editingMeetingNotes);
                        setShowNotesModal(false);
                        setEditingNotesItem(null);
                        setEditingMeetingNotes('');
                        setShowSaveDropdown(false);
                      }
                    }}
                    disabled={!editingMeetingNotes.trim()}
                    className="flex-1 bg-blue-600 hover:bg-blue-700 text-white min-h-[44px] touch-manipulation rounded-r-none border-r border-blue-500 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <i className="fas fa-save mr-2"></i>
                    Save Notes
                  </Button>
                  {/* Dropdown trigger */}
                  <Button
                    onClick={() => setShowSaveDropdown(prev => !prev)}
                    disabled={!editingMeetingNotes.trim()}
                    className="bg-blue-600 hover:bg-blue-700 text-white min-h-[44px] touch-manipulation rounded-l-none px-3 border-l border-blue-500 disabled:opacity-50 disabled:cursor-not-allowed"
                    aria-label="Save and update status"
                  >
                    <i className="fas fa-chevron-down text-xs"></i>
                  </Button>
                  {/* Dropdown menu */}
                  {showSaveDropdown && (
                    <div className="absolute bottom-full right-0 mb-1 w-56 bg-white border border-gray-200 rounded-lg shadow-lg z-50 overflow-hidden">
                      <div className="px-3 py-2 text-xs font-semibold text-gray-500 uppercase tracking-wide border-b border-gray-100 bg-gray-50">
                        Save &amp; Change Status
                      </div>
                      {['In Discussion', 'Actioned', 'Closed'].map((status) => {
                        const icons: Record<string, string> = {
                          'In Discussion': 'fa-comments',
                          'Actioned': 'fa-check-circle',
                          'Closed': 'fa-lock',
                        };
                        const colors: Record<string, string> = {
                          'In Discussion': 'text-blue-600',
                          'Actioned': 'text-green-600',
                          'Closed': 'text-gray-500',
                        };
                        return (
                          <button
                            key={status}
                            onClick={() => {
                              if (editingNotesItem) {
                                updateMeetingNotes(editingNotesItem, editingMeetingNotes);
                                updateItemStatus(editingNotesItem, status);
                                setShowNotesModal(false);
                                setEditingNotesItem(null);
                                setEditingMeetingNotes('');
                                setShowSaveDropdown(false);
                              }
                            }}
                            className="w-full flex items-center gap-3 px-4 py-3 text-sm text-gray-700 hover:bg-blue-50 hover:text-blue-700 transition-colors text-left"
                          >
                            <i className={`fas ${icons[status]} ${colors[status]} w-4 text-center`}></i>
                            <span>Save &amp; Set <strong>{status}</strong></span>
                          </button>
                        );
                      })}
                    </div>
                  )}
                </div>
              </div>
            </div>
          </DialogContent>
        </Dialog>

        {/* Signature Carousel Modal */}
        {showSignatureCarousel && (() => {
          const presentAttendees = [
            ...meetingAttendees.management,
            ...meetingAttendees.glaziers
          ].filter(a => isAttending(showSignatureCarousel, a.name));
          const existingSigs = meetingSignatures[showSignatureCarousel] ?? {};
          const displayDate = (() => { try { return new Date(showSignatureCarousel).toLocaleDateString('en-NZ', { day: '2-digit', month: 'short', year: 'numeric' }); } catch { return showSignatureCarousel; } })();
          return (
            <SignatureCarousel
              meetingDate={displayDate}
              attendees={presentAttendees}
              existingSignatures={existingSigs}
              onSaveSignature={async (attendeeName, status, signatureData, signedAt) => {
                await saveSignatureMutation.mutateAsync({
                  meetingDate: showSignatureCarousel,
                  attendeeName,
                  status,
                  signatureData,
                  signedAt
                });
              }}
              onComplete={() => setShowSignatureCarousel('')}
              onClose={() => setShowSignatureCarousel('')}
            />
          );
        })()}

        {/* Floating "Saving..." banner - shows immediately when submitting to SharePoint */}
        {isCreatingItem && (
          <div className="fixed top-4 left-1/2 -translate-x-1/2 z-50 flex items-center gap-2 bg-blue-700 text-white text-sm font-medium px-5 py-3 rounded-full shadow-2xl animate-in fade-in slide-in-from-top-2 duration-300">
            <Loader2 className="h-4 w-4 animate-spin" />
            Saving to SharePoint…
          </div>
        )}

        {/* Floating Add Button with Scroll Detection */}
        {showFloatingAdd && floatingMeetingDate && !showSignatureCarousel && (
          <div className="fixed bottom-6 right-6 z-50">
            <button
              onClick={() => {
                setNewIdeaForm(prev => ({ ...prev, meetingDate: floatingMeetingDate }));
                setShowAddIdeaModal(true);
                fetchSharePointFormData();
              }}
              className="group bg-gradient-to-r from-blue-600 to-blue-700 hover:from-blue-700 hover:to-blue-800 text-white p-4 rounded-full shadow-lg hover:shadow-2xl transition-all duration-300 transform hover:scale-110 active:scale-95"
              title={`Add item to ${formatDisplayDate(floatingMeetingDate, 'date-only')} meeting`}
            >
              <Plus className="h-6 w-6 transition-transform group-hover:rotate-90" />
            </button>
            {/* Enhanced tooltip */}
            <div className="absolute bottom-full right-0 mb-3 px-4 py-2 bg-gray-900 text-white text-sm font-medium rounded-lg shadow-lg opacity-0 group-hover:opacity-100 transition-opacity duration-200 whitespace-nowrap pointer-events-none">
              Add to {formatDisplayDate(floatingMeetingDate, 'date-only')}
              <div className="absolute top-full right-4 w-0 h-0 border-l-4 border-r-4 border-t-4 border-transparent border-t-gray-900"></div>
            </div>
          </div>
        )}

        </div>
      </div>
    </div>
  );
}