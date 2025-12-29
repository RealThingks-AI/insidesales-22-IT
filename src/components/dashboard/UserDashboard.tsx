import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { useNavigate } from "react-router-dom";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { useAuth } from "@/hooks/useAuth";
import { 
  Users, FileText, Briefcase, TrendingUp, Clock, Plus, Settings2, Calendar, Activity, Bell, AlertCircle, Info, 
  Mail, Building2, ListTodo, CalendarClock, ClipboardList, Check, X
} from "lucide-react";
import { useState, useRef, useEffect, useCallback } from "react";
import { WidgetKey, WidgetLayoutConfig, DEFAULT_WIDGETS } from "./DashboardCustomizeModal";
import { ResizableDashboard } from "./ResizableDashboard";
import { toast } from "sonner";
import { format, isBefore, addDays, startOfWeek, endOfWeek, isToday } from "date-fns";
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from "@/components/ui/tooltip";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { ScrollArea } from "@/components/ui/scroll-area";
import { TaskModal } from "@/components/tasks/TaskModal";
import { MeetingModal } from "@/components/MeetingModal";
import { LeadModal } from "@/components/LeadModal";
import { ContactModal } from "@/components/ContactModal";
import { AccountModal } from "@/components/AccountModal";
import { useTasks } from "@/hooks/useTasks";
import { Task } from "@/types/task";
import { EmptyState } from "@/components/shared/EmptyState";

const GRID_COLS = 12;

// Utility: Compact layouts to remove all gaps (both vertical and horizontal)
const compactLayoutsUtil = (layouts: WidgetLayoutConfig, visibleKeys: WidgetKey[]): WidgetLayoutConfig => {
  const items = visibleKeys
    .filter(key => layouts[key])
    .map(key => ({ key, ...layouts[key] }))
    .sort((a, b) => a.y - b.y || a.x - b.x);
  
  const compacted: WidgetLayoutConfig = {};
  const grid: boolean[][] = [];
  
  const canPlace = (x: number, y: number, w: number, h: number): boolean => {
    if (x < 0 || x + w > GRID_COLS) return false;
    for (let dy = 0; dy < h; dy++) {
      for (let dx = 0; dx < w; dx++) {
        if (grid[y + dy]?.[x + dx]) return false;
      }
    }
    return true;
  };
  
  const occupy = (x: number, y: number, w: number, h: number) => {
    for (let dy = 0; dy < h; dy++) {
      if (!grid[y + dy]) grid[y + dy] = new Array(GRID_COLS).fill(false);
      for (let dx = 0; dx < w; dx++) {
        grid[y + dy][x + dx] = true;
      }
    }
  };
  
  items.forEach(item => {
    let placed = false;
    for (let y = 0; y < 100 && !placed; y++) {
      for (let x = 0; x <= GRID_COLS - item.w && !placed; x++) {
        if (canPlace(x, y, item.w, item.h)) {
          occupy(x, y, item.w, item.h);
          compacted[item.key] = { x, y, w: item.w, h: item.h };
          placed = true;
        }
      }
    }
    if (!placed) {
      const fallbackY = Object.keys(compacted).length * 2;
      occupy(0, fallbackY, item.w, item.h);
      compacted[item.key] = { x: 0, y: fallbackY, w: item.w, h: item.h };
    }
  });
  
  return compacted;
};

const UserDashboard = () => {
  const { user } = useAuth();
  const navigate = useNavigate();
  const queryClient = useQueryClient();
  const [isResizeMode, setIsResizeMode] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);
  const [containerWidth, setContainerWidth] = useState(1200);
  
  const [pendingWidgetChanges, setPendingWidgetChanges] = useState<Set<WidgetKey>>(new Set());
  const [originalState, setOriginalState] = useState<{
    visible: WidgetKey[];
    order: WidgetKey[];
    layouts: WidgetLayoutConfig;
  } | null>(null);
  
  // Modal states
  const [selectedTask, setSelectedTask] = useState<Task | null>(null);
  const [taskModalOpen, setTaskModalOpen] = useState(false);
  const [selectedMeeting, setSelectedMeeting] = useState<any>(null);
  const [meetingModalOpen, setMeetingModalOpen] = useState(false);
  const [leadModalOpen, setLeadModalOpen] = useState(false);
  const [contactModalOpen, setContactModalOpen] = useState(false);
  const [accountModalOpen, setAccountModalOpen] = useState(false);
  const [createMeetingModalOpen, setCreateMeetingModalOpen] = useState(false);
  
  const { createTask, updateTask } = useTasks();

  useEffect(() => {
    const updateWidth = () => {
      if (containerRef.current) {
        setContainerWidth(containerRef.current.offsetWidth - 48);
      }
    };
    updateWidth();
    window.addEventListener('resize', updateWidth);
    return () => window.removeEventListener('resize', updateWidth);
  }, []);
  
  const { data: userName } = useQuery({
    queryKey: ['user-profile-name', user?.id],
    queryFn: async () => {
      if (!user?.id) return null;
      const { data, error } = await supabase
        .from('profiles')
        .select('full_name')
        .eq('id', user.id)
        .maybeSingle();
      if (error) throw error;
      const name = data?.full_name;
      if (!name || name.includes('@')) {
        return user.email?.split('@')[0] || null;
      }
      return name;
    },
    enabled: !!user?.id,
  });

  const { data: dashboardPrefs } = useQuery({
    queryKey: ['dashboard-prefs', user?.id],
    queryFn: async () => {
      if (!user?.id) return null;
      const { data, error } = await supabase
        .from('dashboard_preferences')
        .select('visible_widgets, card_order, layout_view')
        .eq('user_id', user.id)
        .maybeSingle();
      if (error) throw error;
      return data;
    },
    enabled: !!user?.id,
  });

  const defaultWidgetKeys = DEFAULT_WIDGETS.map((w) => w.key);
  const defaultVisibleWidgets = defaultWidgetKeys.filter(
    (k) => DEFAULT_WIDGETS.find((w) => w.key === k)?.visible
  );

  const [visibleWidgets, setVisibleWidgets] = useState<WidgetKey[]>(defaultVisibleWidgets);
  const [widgetOrder, setWidgetOrder] = useState<WidgetKey[]>(defaultWidgetKeys);

  const parseWidgetLayouts = (): WidgetLayoutConfig => {
    if (!dashboardPrefs?.layout_view) return {};
    if (typeof dashboardPrefs.layout_view === "object") {
      return dashboardPrefs.layout_view as WidgetLayoutConfig;
    }
    if (typeof dashboardPrefs.layout_view === "string") {
      try {
        const parsed = JSON.parse(dashboardPrefs.layout_view);
        if (typeof parsed === "object" && parsed !== null) {
          return parsed as WidgetLayoutConfig;
        }
      } catch {
        // Legacy string value
      }
    }
    return {};
  };

  const [widgetLayouts, setWidgetLayouts] = useState<WidgetLayoutConfig>(parseWidgetLayouts());

  useEffect(() => {
    setIsResizeMode(false);
    if (!user?.id) return;

    const sanitizeKeys = (keys: WidgetKey[]) => {
      const allowed = new Set(defaultWidgetKeys);
      const uniq: WidgetKey[] = [];
      const seen = new Set<string>();
      keys.forEach((k) => {
        if (!allowed.has(k)) return;
        if (seen.has(k)) return;
        seen.add(k);
        uniq.push(k);
      });
      return uniq;
    };

    const nextVisibleRaw: WidgetKey[] = dashboardPrefs?.visible_widgets
      ? (dashboardPrefs.visible_widgets as WidgetKey[])
      : defaultVisibleWidgets;

    const nextOrderRaw: WidgetKey[] = dashboardPrefs?.card_order
      ? (dashboardPrefs.card_order as WidgetKey[])
      : defaultWidgetKeys;

    const nextVisible = sanitizeKeys(nextVisibleRaw);
    const nextOrderBase = sanitizeKeys(nextOrderRaw);
    const missingVisible = nextVisible.filter((k) => !nextOrderBase.includes(k));
    const nextOrder = [...nextOrderBase, ...missingVisible];

    const loadedLayouts = parseWidgetLayouts();
    const compactedLayouts = compactLayoutsUtil(loadedLayouts, nextVisible);

    setVisibleWidgets(nextVisible);
    setWidgetOrder(nextOrder);
    setWidgetLayouts(compactedLayouts);
  }, [user?.id, dashboardPrefs?.visible_widgets, dashboardPrefs?.card_order, dashboardPrefs?.layout_view]);

  const savePreferencesMutation = useMutation({
    mutationFn: async ({ widgets, order, layouts }: { widgets: WidgetKey[], order: WidgetKey[], layouts: WidgetLayoutConfig }) => {
      if (!user?.id) throw new Error("User not authenticated");
      const { data, error } = await supabase
        .from('dashboard_preferences')
        .upsert({
          user_id: user.id,
          visible_widgets: widgets,
          card_order: order,
          layout_view: JSON.stringify(layouts),
          updated_at: new Date().toISOString(),
        }, { onConflict: 'user_id' })
        .select();
      if (error) throw error;
      return data;
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['dashboard-prefs', user?.id] });
      toast.success("Dashboard layout saved");
    },
    onError: () => {
      toast.error("Failed to save layout");
    },
  });

  const handleLayoutChange = useCallback((newLayouts: WidgetLayoutConfig) => {
    const compacted = compactLayoutsUtil(newLayouts, visibleWidgets);
    setWidgetLayouts(compacted);
  }, [visibleWidgets]);

  const handleWidgetRemove = useCallback((key: WidgetKey) => {
    const isCurrentlyVisible = visibleWidgets.includes(key);
    setPendingWidgetChanges((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      toast(isCurrentlyVisible 
        ? (next.has(key) ? "Marked for removal" : "Removal undone")
        : (next.has(key) ? "Marked to add" : "Add undone"));
      return next;
    });
  }, [visibleWidgets]);

  const togglePendingWidget = useCallback((key: WidgetKey) => {
    setPendingWidgetChanges(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  }, []);

  const willWidgetBeVisible = useCallback((key: WidgetKey) => {
    const isCurrentlyVisible = visibleWidgets.includes(key);
    const isPending = pendingWidgetChanges.has(key);
    return isPending ? !isCurrentlyVisible : isCurrentlyVisible;
  }, [visibleWidgets, pendingWidgetChanges]);

  const findNextGridPosition = useCallback((existingLayouts: WidgetLayoutConfig, widgetWidth: number, widgetHeight: number) => {
    const COLS = 12;
    const grid: boolean[][] = [];
    Object.values(existingLayouts).forEach(layout => {
      if (!layout) return;
      for (let row = layout.y; row < layout.y + layout.h; row++) {
        if (!grid[row]) grid[row] = new Array(COLS).fill(false);
        for (let col = layout.x; col < Math.min(layout.x + layout.w, COLS); col++) {
          grid[row][col] = true;
        }
      }
    });
    for (let y = 0; y < 100; y++) {
      if (!grid[y]) grid[y] = new Array(COLS).fill(false);
      for (let x = 0; x <= COLS - widgetWidth; x++) {
        let fits = true;
        for (let dy = 0; dy < widgetHeight && fits; dy++) {
          if (!grid[y + dy]) grid[y + dy] = new Array(COLS).fill(false);
          for (let dx = 0; dx < widgetWidth && fits; dx++) {
            if (grid[y + dy][x + dx]) fits = false;
          }
        }
        if (fits) return { x, y };
      }
    }
    return { x: 0, y: Object.keys(existingLayouts).length * 2 };
  }, []);

  const handleSaveLayout = () => {
    let finalVisible = [...visibleWidgets];
    let finalOrder = [...widgetOrder];
    let finalLayouts = { ...widgetLayouts };
    
    pendingWidgetChanges.forEach(key => {
      const isCurrentlyVisible = visibleWidgets.includes(key);
      if (isCurrentlyVisible) {
        finalVisible = finalVisible.filter(w => w !== key);
        finalOrder = finalOrder.filter(w => w !== key);
        delete finalLayouts[key];
      } else {
        finalVisible.push(key);
        if (!finalOrder.includes(key)) finalOrder.push(key);
        const position = findNextGridPosition(finalLayouts, 3, 2);
        finalLayouts[key] = { x: position.x, y: position.y, w: 3, h: 2 };
      }
    });
    
    const compactedLayouts = compactLayoutsUtil(finalLayouts, finalVisible);
    setVisibleWidgets(finalVisible);
    setWidgetOrder(finalOrder);
    setWidgetLayouts(compactedLayouts);
    savePreferencesMutation.mutate({ widgets: finalVisible, order: finalOrder, layouts: compactedLayouts });
    setPendingWidgetChanges(new Set());
    setOriginalState(null);
    setIsResizeMode(false);
  };

  const handleEnterCustomizeMode = useCallback(() => {
    setOriginalState({ visible: [...visibleWidgets], order: [...widgetOrder], layouts: { ...widgetLayouts } });
    setIsResizeMode(true);
  }, [visibleWidgets, widgetOrder, widgetLayouts]);

  const handleCancelCustomize = useCallback(() => {
    if (originalState) {
      setVisibleWidgets(originalState.visible);
      setWidgetOrder(originalState.order);
      setWidgetLayouts(originalState.layouts);
    }
    setPendingWidgetChanges(new Set());
    setOriginalState(null);
    setIsResizeMode(false);
    toast.info("Changes discarded");
  }, [originalState]);

  useEffect(() => {
    if (!isResizeMode) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        handleCancelCustomize();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isResizeMode, handleCancelCustomize]);

  // ================== DATA QUERIES ==================

  // Leads data - enhanced
  const { data: leadsData, isLoading: leadsLoading } = useQuery({
    queryKey: ['user-leads-enhanced', user?.id],
    queryFn: async () => {
      const { data, error } = await supabase.from('leads').select('id, lead_status, lead_name, created_time').eq('created_by', user?.id);
      if (error) throw error;
      const leads = data || [];
      const recentLead = leads.sort((a, b) => new Date(b.created_time || 0).getTime() - new Date(a.created_time || 0).getTime())[0];
      return {
        total: leads.length,
        new: leads.filter(l => l.lead_status === 'New').length,
        contacted: leads.filter(l => l.lead_status === 'Contacted').length,
        qualified: leads.filter(l => l.lead_status === 'Qualified').length,
        converted: leads.filter(l => l.lead_status === 'Converted').length,
        recentLead: recentLead?.lead_name || null
      };
    },
    enabled: !!user?.id
  });

  // Contacts data - enhanced
  const { data: contactsData, isLoading: contactsLoading } = useQuery({
    queryKey: ['user-contacts-enhanced', user?.id],
    queryFn: async () => {
      const { data, error } = await supabase.from('contacts').select('id, contact_name, email, phone_no, segment, created_time').eq('created_by', user?.id);
      if (error) throw error;
      const contacts = data || [];
      const withEmail = contacts.filter(c => c.email).length;
      const withPhone = contacts.filter(c => c.phone_no).length;
      const prospects = contacts.filter(c => c.segment === 'prospect').length;
      const customers = contacts.filter(c => c.segment === 'customer').length;
      const recentContact = contacts.sort((a, b) => new Date(b.created_time || 0).getTime() - new Date(a.created_time || 0).getTime())[0];
      return { total: contacts.length, withEmail, withPhone, prospects, customers, recentContact: recentContact?.contact_name || null };
    },
    enabled: !!user?.id
  });

  // Deals data - enhanced
  const { data: dealsData, isLoading: dealsLoading } = useQuery({
    queryKey: ['user-deals-enhanced', user?.id],
    queryFn: async () => {
      const { data, error } = await supabase.from('deals').select('id, stage, total_contract_value, deal_name, created_by, lead_owner, expected_closing_date');
      if (error) throw error;
      const userDeals = (data || []).filter(d => d.created_by === user?.id || d.lead_owner === user?.id);
      const activeDeals = userDeals.filter(d => !['Won', 'Lost', 'Dropped'].includes(d.stage));
      const wonDeals = userDeals.filter(d => d.stage === 'Won');
      const totalPipeline = activeDeals.reduce((sum, d) => sum + (d.total_contract_value || 0), 0);
      const wonValue = wonDeals.reduce((sum, d) => sum + (d.total_contract_value || 0), 0);
      
      // Deals closing this month
      const now = new Date();
      const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
      const closingThisMonth = activeDeals.filter(d => {
        if (!d.expected_closing_date) return false;
        const closeDate = new Date(d.expected_closing_date);
        return closeDate <= monthEnd && closeDate >= now;
      });
      
      return {
        total: userDeals.length,
        active: activeDeals.length,
        won: wonDeals.length,
        lost: userDeals.filter(d => d.stage === 'Lost').length,
        totalPipeline,
        wonValue,
        closingThisMonth: closingThisMonth.length,
        closingValue: closingThisMonth.reduce((sum, d) => sum + (d.total_contract_value || 0), 0),
        byStage: {
          lead: userDeals.filter(d => d.stage === 'Lead').length,
          qualified: userDeals.filter(d => d.stage === 'Qualified').length,
          discussions: userDeals.filter(d => d.stage === 'Discussions').length,
        }
      };
    },
    enabled: !!user?.id
  });

  // Accounts data - enhanced
  const { data: accountsData, isLoading: accountsLoading } = useQuery({
    queryKey: ['user-accounts-enhanced', user?.id],
    queryFn: async () => {
      const now = new Date();
      const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
      const { data, error } = await supabase.from('accounts').select('id, company_name, segment, status, created_at, total_revenue').eq('created_by', user?.id);
      if (error) throw error;
      const accounts = data || [];
      const newThisMonth = accounts.filter(a => new Date(a.created_at || 0) >= monthStart).length;
      const bySegment = {
        prospect: accounts.filter(a => a.segment === 'prospect').length,
        customer: accounts.filter(a => a.segment === 'customer').length,
        partner: accounts.filter(a => a.segment === 'partner').length,
      };
      const totalRevenue = accounts.reduce((sum, a) => sum + (a.total_revenue || 0), 0);
      return { total: accounts.length, newThisMonth, bySegment, totalRevenue };
    },
    enabled: !!user?.id
  });

  // Action items - enhanced
  const { data: actionItemsData, isLoading: actionItemsLoading } = useQuery({
    queryKey: ['user-action-items-enhanced', user?.id],
    queryFn: async () => {
      const { data: dealItems } = await supabase.from('deal_action_items').select('id, status, due_date, next_action').eq('assigned_to', user?.id).eq('status', 'Open').order('due_date', { ascending: true }).limit(5);
      const { data: leadItems } = await supabase.from('lead_action_items').select('id, status, due_date, next_action').eq('assigned_to', user?.id).eq('status', 'Open').order('due_date', { ascending: true }).limit(5);
      const allItems = [...(dealItems || []), ...(leadItems || [])];
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const overdue = allItems.filter(item => item.due_date && new Date(item.due_date) < today).length;
      const dueToday = allItems.filter(item => item.due_date && isToday(new Date(item.due_date))).length;
      const topItems = allItems.slice(0, 3);
      return { total: allItems.length, overdue, dueToday, topItems };
    },
    enabled: !!user?.id
  });

  // Upcoming meetings - enhanced
  const { data: upcomingMeetings } = useQuery({
    queryKey: ['user-upcoming-meetings-enhanced', user?.id],
    queryFn: async () => {
      const now = new Date().toISOString();
      const weekFromNow = addDays(new Date(), 7).toISOString();
      const { data, error } = await supabase
        .from('meetings')
        .select('id, subject, start_time, end_time, status, attendees')
        .eq('created_by', user?.id)
        .gte('start_time', now)
        .lte('start_time', weekFromNow)
        .order('start_time', { ascending: true })
        .limit(5);
      if (error) throw error;
      return (data || []).map(m => ({
        ...m,
        isToday: isToday(new Date(m.start_time)),
        attendeeCount: Array.isArray(m.attendees) ? m.attendees.length : 0
      }));
    },
    enabled: !!user?.id
  });

  // Today's meetings for agenda
  const { data: todaysMeetings } = useQuery({
    queryKey: ['user-todays-meetings', user?.id],
    queryFn: async () => {
      const todayStart = new Date();
      todayStart.setHours(0, 0, 0, 0);
      const todayEnd = new Date();
      todayEnd.setHours(23, 59, 59, 999);
      const { data, error } = await supabase
        .from('meetings')
        .select('id, subject, start_time, end_time, status')
        .eq('created_by', user?.id)
        .gte('start_time', todayStart.toISOString())
        .lte('start_time', todayEnd.toISOString())
        .order('start_time', { ascending: true });
      if (error) throw error;
      return data || [];
    },
    enabled: !!user?.id
  });

  // Today's tasks for agenda
  const { data: todaysTasks } = useQuery({
    queryKey: ['user-todays-tasks', user?.id],
    queryFn: async () => {
      const today = format(new Date(), 'yyyy-MM-dd');
      const { data, error } = await supabase
        .from('tasks')
        .select('id, title, due_date, priority, status')
        .or(`assigned_to.eq.${user?.id},created_by.eq.${user?.id}`)
        .in('status', ['open', 'in_progress'])
        .eq('due_date', today)
        .order('priority', { ascending: true });
      if (error) throw error;
      return data || [];
    },
    enabled: !!user?.id
  });

  // Overdue tasks for agenda
  const { data: overdueTasks } = useQuery({
    queryKey: ['user-overdue-tasks', user?.id],
    queryFn: async () => {
      const today = format(new Date(), 'yyyy-MM-dd');
      const { data, error } = await supabase
        .from('tasks')
        .select('id, title, due_date, priority, status')
        .or(`assigned_to.eq.${user?.id},created_by.eq.${user?.id}`)
        .in('status', ['open', 'in_progress'])
        .lt('due_date', today)
        .order('due_date', { ascending: true })
        .limit(5);
      if (error) throw error;
      return data || [];
    },
    enabled: !!user?.id
  });

  // Task reminders
  const { data: taskReminders } = useQuery({
    queryKey: ['user-task-reminders-enhanced', user?.id],
    queryFn: async () => {
      const weekFromNow = format(addDays(new Date(), 7), 'yyyy-MM-dd');
      const today = format(new Date(), 'yyyy-MM-dd');
      const { data, error } = await supabase
        .from('tasks')
        .select('id, title, due_date, priority, status')
        .or(`assigned_to.eq.${user?.id},created_by.eq.${user?.id}`)
        .in('status', ['open', 'in_progress'])
        .lte('due_date', weekFromNow)
        .order('due_date', { ascending: true })
        .limit(10);
      if (error) throw error;
      const tasks = data || [];
      const overdue = tasks.filter(t => t.due_date && t.due_date < today).length;
      const dueToday = tasks.filter(t => t.due_date === today).length;
      const highPriority = tasks.filter(t => t.priority === 'high').length;
      return { tasks: tasks.slice(0, 5), overdue, dueToday, highPriority, total: tasks.length };
    },
    enabled: !!user?.id
  });

  // Email stats - enhanced
  const { data: emailStats } = useQuery({
    queryKey: ['user-email-stats-enhanced', user?.id],
    queryFn: async () => {
      const { data, error } = await supabase
        .from('email_history')
        .select('id, status, open_count, click_count, subject, sent_at')
        .eq('sent_by', user?.id)
        .order('sent_at', { ascending: false });
      if (error) throw error;
      const emails = data || [];
      const sent = emails.length;
      const opened = emails.filter(e => (e.open_count || 0) > 0).length;
      const clicked = emails.filter(e => (e.click_count || 0) > 0).length;
      const openRate = sent > 0 ? Math.round((opened / sent) * 100) : 0;
      const clickRate = sent > 0 ? Math.round((clicked / sent) * 100) : 0;
      const recentEmail = emails[0];
      return { sent, opened, clicked, openRate, clickRate, recentSubject: recentEmail?.subject || null };
    },
    enabled: !!user?.id
  });

  // Follow-ups due
  const { data: followUpsDue } = useQuery({
    queryKey: ['user-follow-ups-due', user?.id],
    queryFn: async () => {
      const { data, error } = await supabase
        .from('meeting_follow_ups')
        .select('id, title, status, due_date, meeting_id')
        .eq('assigned_to', user?.id)
        .eq('status', 'pending')
        .order('due_date', { ascending: true })
        .limit(5);
      if (error) throw error;
      const followUps = data || [];
      const today = format(new Date(), 'yyyy-MM-dd');
      const overdue = followUps.filter(f => f.due_date && f.due_date < today).length;
      return { followUps, total: followUps.length, overdue };
    },
    enabled: !!user?.id
  });

  // Weekly summary
  const { data: weeklySummary } = useQuery({
    queryKey: ['user-weekly-summary', user?.id],
    queryFn: async () => {
      const weekStart = startOfWeek(new Date(), { weekStartsOn: 1 });
      const weekEnd = endOfWeek(new Date(), { weekStartsOn: 1 });
      const startStr = weekStart.toISOString();
      const endStr = weekEnd.toISOString();
      
      const [leadsRes, contactsRes, dealsRes, meetingsRes, tasksRes] = await Promise.all([
        supabase.from('leads').select('id', { count: 'exact', head: true }).eq('created_by', user?.id).gte('created_time', startStr).lte('created_time', endStr),
        supabase.from('contacts').select('id', { count: 'exact', head: true }).eq('created_by', user?.id).gte('created_time', startStr).lte('created_time', endStr),
        supabase.from('deals').select('id', { count: 'exact', head: true }).eq('created_by', user?.id).gte('created_at', startStr).lte('created_at', endStr),
        supabase.from('meetings').select('id', { count: 'exact', head: true }).eq('created_by', user?.id).eq('status', 'completed').gte('start_time', startStr).lte('start_time', endStr),
        supabase.from('tasks').select('id', { count: 'exact', head: true }).or(`assigned_to.eq.${user?.id},created_by.eq.${user?.id}`).eq('status', 'completed').gte('completed_at', startStr).lte('completed_at', endStr),
      ]);
      
      return {
        newLeads: leadsRes.count || 0,
        newContacts: contactsRes.count || 0,
        newDeals: dealsRes.count || 0,
        meetingsCompleted: meetingsRes.count || 0,
        tasksCompleted: tasksRes.count || 0,
      };
    },
    enabled: !!user?.id
  });

  // Recent activities
  const { data: userProfiles } = useQuery({
    queryKey: ['all-user-profiles'],
    queryFn: async () => {
      const { data, error } = await supabase.from('profiles').select('id, full_name');
      if (error) throw error;
      return data || [];
    },
    staleTime: 5 * 60 * 1000,
  });

  const getDisplayName = (value: any): string => {
    if (!value || value === 'empty' || value === null) return 'empty';
    if (typeof value !== 'string') return String(value);
    const uuidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (uuidPattern.test(value)) {
      const profile = userProfiles?.find(p => p.id === value);
      return profile?.full_name || 'Unknown User';
    }
    return value;
  };

  const { data: recentActivities } = useQuery({
    queryKey: ['user-recent-activities', user?.id, userProfiles],
    queryFn: async () => {
      const { data, error } = await supabase
        .from('security_audit_log')
        .select('id, action, resource_type, resource_id, created_at, details, user_id')
        .eq('user_id', user?.id)
        .in('action', ['CREATE', 'UPDATE', 'DELETE'])
        .in('resource_type', ['contacts', 'leads', 'deals', 'accounts', 'meetings', 'tasks'])
        .order('created_at', { ascending: false })
        .limit(5);
      if (error) throw error;

      return (data || []).map(log => {
        let detailedSubject = `${log.action} ${log.resource_type}`;
        const details = log.details as any;
        
        if (log.action === 'UPDATE' && details?.field_changes) {
          const changedFields = Object.keys(details.field_changes);
          if (changedFields.length > 0) {
            const fieldSummary = changedFields.slice(0, 2).map(field => {
              const change = details.field_changes[field];
              const oldVal = getDisplayName(change?.old ?? 'empty');
              const newVal = getDisplayName(change?.new ?? 'empty');
              return `${field}: "${oldVal}" → "${newVal}"`;
            }).join(', ');
            detailedSubject = `Updated ${log.resource_type} - ${fieldSummary}${changedFields.length > 2 ? ` (+${changedFields.length - 2} more)` : ''}`;
          }
        } else if (log.action === 'CREATE' && details?.record_data) {
          const recordName = details.record_data.lead_name || details.record_data.contact_name || 
                            details.record_data.deal_name || details.record_data.company_name || 
                            details.record_data.title || details.record_data.subject || '';
          if (recordName) detailedSubject = `Created ${log.resource_type} - "${recordName}"`;
        } else if (log.action === 'DELETE' && details?.deleted_data) {
          const recordName = details.deleted_data.lead_name || details.deleted_data.contact_name || 
                            details.deleted_data.deal_name || details.deleted_data.company_name || 
                            details.deleted_data.title || details.deleted_data.subject || '';
          if (recordName) detailedSubject = `Deleted ${log.resource_type} - "${recordName}"`;
        }
        
        return {
          id: log.id,
          subject: detailedSubject,
          activity_type: log.action,
          activity_date: log.created_at,
          resource_type: log.resource_type,
        };
      });
    },
    enabled: !!user?.id && !!userProfiles
  });

  const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'EUR',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(amount);
  };

  const isLoading = leadsLoading || contactsLoading || dealsLoading || actionItemsLoading || accountsLoading;

  if (isLoading) {
    return (
      <div className="p-6 space-y-6">
        <div className="flex items-center justify-between">
          <div className="h-8 w-64 rounded-md skeleton-shimmer" />
          <div className="h-9 w-24 rounded-md skeleton-shimmer" />
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
          {[...Array(8)].map((_, i) => (
            <div key={i} className="h-32 rounded-lg skeleton-shimmer" style={{ animationDelay: `${i * 0.1}s` }} />
          ))}
        </div>
      </div>
    );
  }

  const renderWidget = (key: WidgetKey) => {
    switch (key) {
      case "leads":
        return (
          <Card className="h-full hover:shadow-lg transition-shadow cursor-pointer animate-fade-in" onClick={() => !isResizeMode && navigate('/leads')}>
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium">My Leads</CardTitle>
              <FileText className="w-4 h-4 text-blue-600" />
            </CardHeader>
            <CardContent className="space-y-2">
              <div className="text-2xl font-bold">{leadsData?.total || 0}</div>
              <div className="flex flex-wrap gap-1 text-xs">
                <span className="px-2 py-0.5 rounded bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-300">{leadsData?.new || 0} New</span>
                <span className="px-2 py-0.5 rounded bg-yellow-100 text-yellow-700 dark:bg-yellow-900/30 dark:text-yellow-300">{leadsData?.contacted || 0} Contacted</span>
                <span className="px-2 py-0.5 rounded bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-300">{leadsData?.qualified || 0} Qualified</span>
              </div>
              {leadsData?.recentLead && (
                <p className="text-xs text-muted-foreground truncate">Latest: {leadsData.recentLead}</p>
              )}
            </CardContent>
          </Card>
        );

      case "contacts":
        return (
          <Card className="h-full hover:shadow-lg transition-shadow cursor-pointer animate-fade-in" onClick={() => !isResizeMode && navigate('/contacts')}>
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium">My Contacts</CardTitle>
              <Users className="w-4 h-4 text-green-600" />
            </CardHeader>
            <CardContent className="space-y-2">
              <div className="text-2xl font-bold">{contactsData?.total || 0}</div>
              <div className="flex flex-wrap gap-1 text-xs">
                <span className="px-2 py-0.5 rounded bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-300">{contactsData?.withEmail || 0} w/ Email</span>
                <span className="px-2 py-0.5 rounded bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-300">{contactsData?.withPhone || 0} w/ Phone</span>
              </div>
              <div className="flex gap-2 text-xs text-muted-foreground">
                <span>{contactsData?.prospects || 0} Prospects</span>
                <span>•</span>
                <span>{contactsData?.customers || 0} Customers</span>
              </div>
            </CardContent>
          </Card>
        );

      case "deals":
        return (
          <Card className="h-full hover:shadow-lg transition-shadow cursor-pointer animate-fade-in" onClick={() => !isResizeMode && navigate('/deals')}>
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium">My Deals</CardTitle>
              <Briefcase className="w-4 h-4 text-purple-600" />
            </CardHeader>
            <CardContent className="space-y-2">
              <div className="flex items-baseline gap-2">
                <span className="text-2xl font-bold">{dealsData?.active || 0}</span>
                <span className="text-sm text-muted-foreground">active</span>
              </div>
              <div className="flex flex-wrap gap-1 text-xs">
                <span className="px-2 py-0.5 rounded bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-300">{dealsData?.won || 0} Won</span>
                <span className="px-2 py-0.5 rounded bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-300">{dealsData?.lost || 0} Lost</span>
              </div>
              <p className="text-xs text-muted-foreground">Pipeline: {formatCurrency(dealsData?.totalPipeline || 0)}</p>
            </CardContent>
          </Card>
        );

      case "accountsSummary":
        return (
          <Card className="h-full hover:shadow-lg transition-shadow cursor-pointer animate-fade-in" onClick={() => !isResizeMode && navigate('/accounts')}>
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium">Accounts</CardTitle>
              <Building2 className="w-4 h-4 text-indigo-600" />
            </CardHeader>
            <CardContent className="space-y-2">
              <div className="flex items-baseline gap-2">
                <span className="text-2xl font-bold">{accountsData?.total || 0}</span>
                {(accountsData?.newThisMonth || 0) > 0 && (
                  <span className="text-xs text-green-600">+{accountsData?.newThisMonth} this month</span>
                )}
              </div>
              <div className="flex flex-wrap gap-1 text-xs">
                <span className="px-2 py-0.5 rounded bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-300">{accountsData?.bySegment?.prospect || 0} Prospect</span>
                <span className="px-2 py-0.5 rounded bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-300">{accountsData?.bySegment?.customer || 0} Customer</span>
              </div>
            </CardContent>
          </Card>
        );

      case "actionItems":
        return (
          <Card className="h-full animate-fade-in">
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium">Action Items</CardTitle>
              <Clock className="w-4 h-4 text-orange-600" />
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-2xl font-bold">{actionItemsData?.total || 0}</span>
                <div className="flex gap-2 text-xs">
                  {(actionItemsData?.overdue || 0) > 0 && (
                    <span className="px-2 py-1 rounded bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-300 font-medium">
                      {actionItemsData?.overdue} overdue
                    </span>
                  )}
                  {(actionItemsData?.dueToday || 0) > 0 && (
                    <span className="px-2 py-1 rounded bg-orange-100 text-orange-700 dark:bg-orange-900/30 dark:text-orange-300">
                      {actionItemsData?.dueToday} today
                    </span>
                  )}
                </div>
              </div>
              {actionItemsData?.topItems && actionItemsData.topItems.length > 0 && (
                <div className="space-y-1">
                  {actionItemsData.topItems.slice(0, 3).map((item: any) => (
                    <div key={item.id} className="text-xs p-2 rounded bg-muted/50 truncate">
                      {item.next_action}
                    </div>
                  ))}
                </div>
              )}
            </CardContent>
          </Card>
        );

      case "quickActions":
        return (
          <Card className="h-full animate-fade-in">
            <CardHeader className="pb-2">
              <CardTitle className="text-sm font-medium">Quick Actions</CardTitle>
            </CardHeader>
            <CardContent className="grid grid-cols-2 gap-2">
              <Button variant="outline" size="sm" className="justify-start gap-2" onClick={() => !isResizeMode && setLeadModalOpen(true)}>
                <Plus className="w-3 h-3" /> Lead
              </Button>
              <Button variant="outline" size="sm" className="justify-start gap-2" onClick={() => !isResizeMode && setContactModalOpen(true)}>
                <Plus className="w-3 h-3" /> Contact
              </Button>
              <Button variant="outline" size="sm" className="justify-start gap-2" onClick={() => !isResizeMode && setAccountModalOpen(true)}>
                <Plus className="w-3 h-3" /> Account
              </Button>
              <Button variant="outline" size="sm" className="justify-start gap-2" onClick={() => !isResizeMode && setCreateMeetingModalOpen(true)}>
                <Plus className="w-3 h-3" /> Meeting
              </Button>
            </CardContent>
          </Card>
        );

      case "myPipeline":
        return (
          <Card className="h-full hover:shadow-lg transition-shadow cursor-pointer animate-fade-in" onClick={() => !isResizeMode && navigate('/deals')}>
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium">My Pipeline</CardTitle>
              <TrendingUp className="w-4 h-4 text-emerald-600" />
            </CardHeader>
            <CardContent className="space-y-2">
              <div className="text-2xl font-bold">{formatCurrency(dealsData?.totalPipeline || 0)}</div>
              <div className="flex items-center justify-between text-xs">
                <span className="text-muted-foreground">{dealsData?.active || 0} active deals</span>
                {(dealsData?.closingThisMonth || 0) > 0 && (
                  <span className="text-green-600 font-medium">{dealsData?.closingThisMonth} closing soon</span>
                )}
              </div>
              <div className="flex items-center gap-2 text-xs">
                <span className="text-green-600 font-medium">Won: {formatCurrency(dealsData?.wonValue || 0)}</span>
              </div>
            </CardContent>
          </Card>
        );

      case "todaysAgenda":
        const totalAgendaItems = (todaysMeetings?.length || 0) + (todaysTasks?.length || 0) + (overdueTasks?.length || 0);
        return (
          <Card className="h-full animate-fade-in">
            <CardHeader className="flex flex-row items-center justify-between">
              <CardTitle className="flex items-center gap-2">
                <CalendarClock className="w-5 h-5 text-primary" />
                Today's Agenda
              </CardTitle>
              <span className="text-xs text-muted-foreground">{format(new Date(), 'EEE, MMM d')}</span>
            </CardHeader>
            <CardContent>
              {totalAgendaItems > 0 ? (
                <div className="space-y-3 max-h-[200px] overflow-y-auto">
                  {(overdueTasks?.length || 0) > 0 && (
                    <div>
                      <p className="text-xs font-medium text-red-600 mb-1">⚠️ Overdue ({overdueTasks?.length})</p>
                      {overdueTasks?.slice(0, 2).map((task: any) => (
                        <div key={task.id} className="text-xs p-2 rounded bg-red-50 dark:bg-red-900/20 text-red-700 dark:text-red-300 mb-1 truncate">
                          {task.title}
                        </div>
                      ))}
                    </div>
                  )}
                  {(todaysMeetings?.length || 0) > 0 && (
                    <div>
                      <p className="text-xs font-medium text-muted-foreground mb-1">Meetings ({todaysMeetings?.length})</p>
                      {todaysMeetings?.slice(0, 2).map((meeting: any) => (
                        <div key={meeting.id} className="text-xs p-2 rounded bg-blue-50 dark:bg-blue-900/20 mb-1 flex items-center gap-2">
                          <Calendar className="w-3 h-3 text-blue-600" />
                          <span className="truncate">{meeting.subject}</span>
                          <span className="text-muted-foreground ml-auto">{format(new Date(meeting.start_time), 'HH:mm')}</span>
                        </div>
                      ))}
                    </div>
                  )}
                  {(todaysTasks?.length || 0) > 0 && (
                    <div>
                      <p className="text-xs font-medium text-muted-foreground mb-1">Tasks Due ({todaysTasks?.length})</p>
                      {todaysTasks?.slice(0, 2).map((task: any) => (
                        <div key={task.id} className="text-xs p-2 rounded bg-orange-50 dark:bg-orange-900/20 mb-1 truncate">
                          {task.title}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              ) : (
                <EmptyState
                  title="Clear day ahead"
                  description="No meetings or tasks scheduled for today"
                  illustration="calendar"
                  variant="compact"
                />
              )}
            </CardContent>
          </Card>
        );

      case "upcomingMeetings":
        return (
          <Card className="h-full animate-fade-in">
            <CardHeader className="flex flex-row items-center justify-between">
              <CardTitle className="flex items-center gap-2">
                <Calendar className="w-5 h-5 text-primary" />
                Upcoming Meetings
              </CardTitle>
              <Button variant="ghost" size="sm" onClick={() => !isResizeMode && navigate('/meetings')}>View All</Button>
            </CardHeader>
            <CardContent>
              {upcomingMeetings && upcomingMeetings.length > 0 ? (
                <div className="space-y-2 max-h-[200px] overflow-y-auto">
                  {upcomingMeetings.slice(0, 4).map((meeting: any) => (
                    <div 
                      key={meeting.id} 
                      className={`p-2 rounded-lg cursor-pointer transition-colors ${meeting.isToday ? 'bg-primary/10 border border-primary/20' : 'bg-muted/50 hover:bg-muted'}`}
                      onClick={() => { if (!isResizeMode) { setSelectedMeeting(meeting); setMeetingModalOpen(true); }}}
                    >
                      <div className="flex items-center justify-between">
                        <p className="text-sm font-medium truncate flex-1">{meeting.subject}</p>
                        {meeting.isToday && <span className="text-xs px-1.5 py-0.5 rounded bg-primary text-primary-foreground ml-2">Today</span>}
                      </div>
                      <div className="flex items-center gap-2 text-xs text-muted-foreground mt-1">
                        <span>{format(new Date(meeting.start_time), 'EEE, MMM d • HH:mm')}</span>
                        {meeting.attendeeCount > 0 && <span>• {meeting.attendeeCount} attendee{meeting.attendeeCount > 1 ? 's' : ''}</span>}
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <EmptyState
                  title="No upcoming meetings"
                  description="Schedule a meeting to get started"
                  illustration="calendar"
                  actionLabel="Schedule Meeting"
                  onAction={() => !isResizeMode && setCreateMeetingModalOpen(true)}
                  variant="compact"
                />
              )}
            </CardContent>
          </Card>
        );

      case "taskReminders":
        return (
          <Card className="h-full animate-fade-in">
            <CardHeader className="flex flex-row items-center justify-between">
              <CardTitle className="flex items-center gap-2">
                <Bell className="w-5 h-5 text-primary" />
                Task Reminders
              </CardTitle>
              <div className="flex gap-1 text-xs">
                {(taskReminders?.overdue || 0) > 0 && (
                  <span className="px-2 py-0.5 rounded bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-300">{taskReminders?.overdue} overdue</span>
                )}
                {(taskReminders?.highPriority || 0) > 0 && (
                  <span className="px-2 py-0.5 rounded bg-orange-100 text-orange-700 dark:bg-orange-900/30 dark:text-orange-300">{taskReminders?.highPriority} high</span>
                )}
              </div>
            </CardHeader>
            <CardContent>
              {taskReminders?.tasks && taskReminders.tasks.length > 0 ? (
                <div className="space-y-2 max-h-[200px] overflow-y-auto">
                  {taskReminders.tasks.map((task: any) => {
                    const isOverdue = task.due_date && isBefore(new Date(task.due_date), new Date());
                    const isDueToday = task.due_date && isToday(new Date(task.due_date));
                    return (
                      <div 
                        key={task.id} 
                        className={`p-2 rounded-lg cursor-pointer transition-colors ${
                          isOverdue ? 'bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800' :
                          isDueToday ? 'bg-orange-50 dark:bg-orange-900/20 border border-orange-200 dark:border-orange-800' : 'bg-muted/50 hover:bg-muted'
                        }`}
                        onClick={() => { if (!isResizeMode) { setSelectedTask(task as Task); setTaskModalOpen(true); }}}
                      >
                        <div className="flex items-center justify-between">
                          <p className="text-sm font-medium truncate flex-1">{task.title}</p>
                          <span className={`text-xs px-1.5 py-0.5 rounded font-medium ${
                            task.priority === 'high' ? 'bg-red-500 text-white' :
                            task.priority === 'medium' ? 'bg-amber-500 text-white' : 'bg-slate-500 text-white'
                          }`}>{task.priority}</span>
                        </div>
                        <p className={`text-xs mt-1 ${isOverdue ? 'text-red-600' : isDueToday ? 'text-orange-600' : 'text-muted-foreground'}`}>
                          {isOverdue ? 'OVERDUE - ' : isDueToday ? 'Due Today - ' : ''}
                          {task.due_date ? format(new Date(task.due_date), 'MMM d') : 'No date'}
                        </p>
                      </div>
                    );
                  })}
                </div>
              ) : (
                <EmptyState
                  title="No pending tasks"
                  description="You're all caught up!"
                  illustration="tasks"
                  actionLabel="Create Task"
                  onAction={() => { if (!isResizeMode) { setSelectedTask(null); setTaskModalOpen(true); }}}
                  variant="compact"
                />
              )}
            </CardContent>
          </Card>
        );

      case "recentActivities":
        return (
          <Card className="h-full animate-fade-in">
            <CardHeader className="flex flex-row items-center justify-between">
              <CardTitle className="flex items-center gap-2">
                <Activity className="w-5 h-5 text-primary" />
                Recent Activities
              </CardTitle>
              <Button variant="ghost" size="sm" onClick={() => !isResizeMode && navigate('/notifications')}>View All</Button>
            </CardHeader>
            <CardContent>
              {recentActivities && recentActivities.length > 0 ? (
                <div className="space-y-2 max-h-[200px] overflow-y-auto">
                  {recentActivities.slice(0, 5).map((activity) => (
                    <div key={activity.id} className="flex items-start gap-2 p-2 rounded-lg bg-muted/50">
                      <div className="w-6 h-6 rounded-full bg-primary/10 flex items-center justify-center flex-shrink-0 mt-0.5">
                        <Activity className="w-3 h-3 text-primary" />
                      </div>
                      <div className="min-w-0 flex-1">
                        <p className="text-xs font-medium line-clamp-2">{activity.subject}</p>
                        <p className="text-xs text-muted-foreground">{format(new Date(activity.activity_date), 'MMM d, HH:mm')}</p>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <EmptyState
                  title="No recent activities"
                  description="Activities will appear as you work"
                  illustration="activities"
                  variant="compact"
                />
              )}
            </CardContent>
          </Card>
        );

      case "leadStatus":
        return (
          <Card className="h-full animate-fade-in">
            <CardHeader>
              <CardTitle className="flex items-center gap-2">
                Lead Status Overview
                <TooltipProvider>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Info className="w-4 h-4 text-muted-foreground cursor-help" />
                    </TooltipTrigger>
                    <TooltipContent><p>Your leads by status</p></TooltipContent>
                  </Tooltip>
                </TooltipProvider>
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="grid grid-cols-2 gap-3">
                <div className="text-center p-3 bg-blue-50 dark:bg-blue-950/20 rounded-lg">
                  <p className="text-xl font-bold text-blue-600">{leadsData?.new || 0}</p>
                  <p className="text-xs text-muted-foreground">New</p>
                </div>
                <div className="text-center p-3 bg-yellow-50 dark:bg-yellow-950/20 rounded-lg">
                  <p className="text-xl font-bold text-yellow-600">{leadsData?.contacted || 0}</p>
                  <p className="text-xs text-muted-foreground">Contacted</p>
                </div>
                <div className="text-center p-3 bg-green-50 dark:bg-green-950/20 rounded-lg">
                  <p className="text-xl font-bold text-green-600">{leadsData?.qualified || 0}</p>
                  <p className="text-xs text-muted-foreground">Qualified</p>
                </div>
                <div className="text-center p-3 bg-purple-50 dark:bg-purple-950/20 rounded-lg">
                  <p className="text-xl font-bold text-purple-600">{leadsData?.converted || 0}</p>
                  <p className="text-xs text-muted-foreground">Converted</p>
                </div>
              </div>
            </CardContent>
          </Card>
        );

      case "emailStats":
        return (
          <Card className="h-full animate-fade-in">
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium">Email Statistics</CardTitle>
              <Mail className="w-4 h-4 text-blue-600" />
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="grid grid-cols-3 gap-2 text-center">
                <div>
                  <p className="text-xl font-bold">{emailStats?.sent || 0}</p>
                  <p className="text-xs text-muted-foreground">Sent</p>
                </div>
                <div>
                  <p className="text-xl font-bold text-green-600">{emailStats?.opened || 0}</p>
                  <p className="text-xs text-muted-foreground">Opened</p>
                </div>
                <div>
                  <p className="text-xl font-bold text-blue-600">{emailStats?.clicked || 0}</p>
                  <p className="text-xs text-muted-foreground">Clicked</p>
                </div>
              </div>
              <div className="flex justify-between text-xs text-muted-foreground border-t pt-2">
                <span>Open Rate: <span className="font-medium text-foreground">{emailStats?.openRate || 0}%</span></span>
                <span>Click Rate: <span className="font-medium text-foreground">{emailStats?.clickRate || 0}%</span></span>
              </div>
            </CardContent>
          </Card>
        );

      case "weeklySummary":
        return (
          <Card className="h-full animate-fade-in">
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium">This Week</CardTitle>
              <ListTodo className="w-4 h-4 text-teal-600" />
            </CardHeader>
            <CardContent>
              <div className="grid grid-cols-5 gap-2 text-center">
                <div className="p-2 rounded bg-blue-50 dark:bg-blue-950/20">
                  <p className="text-lg font-bold text-blue-600">{weeklySummary?.newLeads || 0}</p>
                  <p className="text-xs text-muted-foreground">Leads</p>
                </div>
                <div className="p-2 rounded bg-green-50 dark:bg-green-950/20">
                  <p className="text-lg font-bold text-green-600">{weeklySummary?.newContacts || 0}</p>
                  <p className="text-xs text-muted-foreground">Contacts</p>
                </div>
                <div className="p-2 rounded bg-purple-50 dark:bg-purple-950/20">
                  <p className="text-lg font-bold text-purple-600">{weeklySummary?.newDeals || 0}</p>
                  <p className="text-xs text-muted-foreground">Deals</p>
                </div>
                <div className="p-2 rounded bg-indigo-50 dark:bg-indigo-950/20">
                  <p className="text-lg font-bold text-indigo-600">{weeklySummary?.meetingsCompleted || 0}</p>
                  <p className="text-xs text-muted-foreground">Meetings</p>
                </div>
                <div className="p-2 rounded bg-emerald-50 dark:bg-emerald-950/20">
                  <p className="text-lg font-bold text-emerald-600">{weeklySummary?.tasksCompleted || 0}</p>
                  <p className="text-xs text-muted-foreground">Tasks Done</p>
                </div>
              </div>
            </CardContent>
          </Card>
        );

      case "followUpsDue":
        return (
          <Card className="h-full animate-fade-in">
            <CardHeader className="flex flex-row items-center justify-between pb-2">
              <CardTitle className="text-sm font-medium">Follow-Ups Due</CardTitle>
              <div className="flex items-center gap-2">
                {(followUpsDue?.overdue || 0) > 0 && (
                  <span className="text-xs px-2 py-0.5 rounded bg-red-100 text-red-700 dark:bg-red-900/30 dark:text-red-300">
                    {followUpsDue?.overdue} overdue
                  </span>
                )}
                <ClipboardList className="w-4 h-4 text-amber-600" />
              </div>
            </CardHeader>
            <CardContent>
              {followUpsDue?.followUps && followUpsDue.followUps.length > 0 ? (
                <div className="space-y-2 max-h-[150px] overflow-y-auto">
                  {followUpsDue.followUps.map((followUp: any) => {
                    const isOverdue = followUp.due_date && isBefore(new Date(followUp.due_date), new Date());
                    return (
                      <div key={followUp.id} className={`p-2 rounded text-xs ${isOverdue ? 'bg-red-50 dark:bg-red-900/20' : 'bg-muted/50'}`}>
                        <p className="font-medium truncate">{followUp.title}</p>
                        <p className={`text-muted-foreground ${isOverdue ? 'text-red-600' : ''}`}>
                          Due: {followUp.due_date ? format(new Date(followUp.due_date), 'MMM d') : 'No date'}
                        </p>
                      </div>
                    );
                  })}
                </div>
              ) : (
                <div className="text-center py-4 text-muted-foreground text-sm">
                  No pending follow-ups
                </div>
              )}
            </CardContent>
          </Card>
        );

      default:
        return null;
    }
  };

  return (
    <div className="p-6 space-y-8" ref={containerRef}>
      {/* Welcome Header */}
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div className="min-w-0 flex-1">
          <h1 className="text-xl sm:text-2xl font-bold text-foreground truncate">
            Welcome back{userName ? `, ${userName}` : ''}!
          </h1>
        </div>
        <div className="flex gap-2 flex-shrink-0 items-center">
          {isResizeMode ? (
            <>
              <div className="bg-primary/10 border border-primary/20 rounded-lg px-3 py-1.5 hidden sm:flex items-center">
                <p className="text-xs text-primary font-medium flex items-center gap-1.5">
                  <Settings2 className="w-3.5 h-3.5" />
                  <span className="hidden md:inline">Drag to move, resize edges, or press Escape to cancel</span>
                  <span className="md:hidden">Edit mode</span>
                </p>
              </div>
              <Popover>
                <PopoverTrigger asChild>
                  <Button variant="outline" className="gap-2">
                    <Plus className="w-4 h-4" />
                    Add Widget
                    {pendingWidgetChanges.size > 0 && (
                      <span className="bg-primary text-primary-foreground text-xs px-1.5 py-0.5 rounded-full">
                        {pendingWidgetChanges.size}
                      </span>
                    )}
                  </Button>
                </PopoverTrigger>
                <PopoverContent className="w-64 p-0" align="end">
                  <div className="p-3 border-b">
                    <p className="text-sm font-medium">Toggle Widgets</p>
                    <p className="text-xs text-muted-foreground">Click to add/remove.</p>
                  </div>
                  <ScrollArea className="h-64">
                    <div className="p-2 space-y-1">
                      {DEFAULT_WIDGETS.map(widget => {
                        const willBeVisible = willWidgetBeVisible(widget.key);
                        const isPending = pendingWidgetChanges.has(widget.key);
                        return (
                          <Button
                            key={widget.key}
                            variant="ghost"
                            className={`w-full justify-between gap-2 ${isPending ? 'bg-primary/10' : ''}`}
                            onClick={() => togglePendingWidget(widget.key)}
                          >
                            <span className="flex items-center gap-2">
                              {widget.icon}
                              {widget.label}
                            </span>
                            {willBeVisible && <Check className="w-4 h-4 text-primary" />}
                          </Button>
                        );
                      })}
                    </div>
                  </ScrollArea>
                </PopoverContent>
              </Popover>
              <Button variant="outline" onClick={handleCancelCustomize} className="gap-2">
                <X className="w-4 h-4" /> Cancel
              </Button>
              <Button onClick={handleSaveLayout} className="gap-2" disabled={savePreferencesMutation.isPending}>
                <Check className="w-4 h-4" /> {savePreferencesMutation.isPending ? 'Saving...' : 'Save'}
              </Button>
            </>
          ) : (
            <Button variant="outline" size="sm" onClick={handleEnterCustomizeMode} className="gap-2">
              <Settings2 className="w-4 h-4" /> Customize
            </Button>
          )}
        </div>
      </div>

      {/* Resizable Grid Layout */}
      <ResizableDashboard
        isResizeMode={isResizeMode}
        visibleWidgets={visibleWidgets}
        widgetLayouts={widgetLayouts}
        pendingWidgetChanges={pendingWidgetChanges}
        onLayoutChange={handleLayoutChange}
        onWidgetRemove={handleWidgetRemove}
        renderWidget={renderWidget}
        containerWidth={containerWidth}
      />
      
      {/* Modals */}
      <TaskModal
        open={taskModalOpen}
        onOpenChange={(open) => { setTaskModalOpen(open); if (!open) setSelectedTask(null); }}
        task={selectedTask}
        onSubmit={createTask}
        onUpdate={async (taskId, updates, original) => {
          const result = await updateTask(taskId, updates, original);
          if (result) queryClient.invalidateQueries({ queryKey: ['user-task-reminders-enhanced', user?.id] });
          return result;
        }}
      />
      
      <MeetingModal
        open={meetingModalOpen}
        onOpenChange={(open) => { setMeetingModalOpen(open); if (!open) setSelectedMeeting(null); }}
        meeting={selectedMeeting}
        onSuccess={() => {
          queryClient.invalidateQueries({ queryKey: ['user-upcoming-meetings-enhanced', user?.id] });
          setMeetingModalOpen(false);
          setSelectedMeeting(null);
        }}
      />
      
      <MeetingModal
        open={createMeetingModalOpen}
        onOpenChange={setCreateMeetingModalOpen}
        meeting={null}
        onSuccess={() => {
          queryClient.invalidateQueries({ queryKey: ['user-upcoming-meetings-enhanced', user?.id] });
          setCreateMeetingModalOpen(false);
          toast.success("Meeting scheduled");
        }}
      />
      
      <LeadModal
        open={leadModalOpen}
        onOpenChange={(open) => { setLeadModalOpen(open); if (!open) queryClient.invalidateQueries({ queryKey: ['user-leads-enhanced', user?.id] }); }}
        lead={null}
        onSuccess={() => {
          queryClient.invalidateQueries({ queryKey: ['user-leads-enhanced', user?.id] });
          setLeadModalOpen(false);
          toast.success("Lead created");
        }}
      />
      
      <ContactModal
        open={contactModalOpen}
        onOpenChange={(open) => { setContactModalOpen(open); if (!open) queryClient.invalidateQueries({ queryKey: ['user-contacts-enhanced', user?.id] }); }}
        contact={null}
        onSuccess={() => {
          queryClient.invalidateQueries({ queryKey: ['user-contacts-enhanced', user?.id] });
          setContactModalOpen(false);
          toast.success("Contact created");
        }}
      />
      
      <AccountModal
        open={accountModalOpen}
        onOpenChange={(open) => { setAccountModalOpen(open); if (!open) queryClient.invalidateQueries({ queryKey: ['user-accounts-enhanced', user?.id] }); }}
        account={null}
        onSuccess={() => {
          queryClient.invalidateQueries({ queryKey: ['user-accounts-enhanced', user?.id] });
          setAccountModalOpen(false);
          toast.success("Account created");
        }}
      />
    </div>
  );
};

export default UserDashboard;
