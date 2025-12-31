/**
 * Centralized status color utilities for consistent styling across the app
 * All colors use CSS variables from the design system
 */

// Lead status colors
export const getLeadStatusColor = (status: string): { bg: string; text: string; border?: string } => {
  const statusLower = status?.toLowerCase() || '';
  
  switch (statusLower) {
    case 'new':
      return { bg: 'bg-blue-100 dark:bg-blue-900/30', text: 'text-blue-700 dark:text-blue-300' };
    case 'attempted':
      return { bg: 'bg-amber-100 dark:bg-amber-900/30', text: 'text-amber-700 dark:text-amber-300' };
    case 'follow-up':
      return { bg: 'bg-purple-100 dark:bg-purple-900/30', text: 'text-purple-700 dark:text-purple-300' };
    case 'qualified':
      return { bg: 'bg-emerald-100 dark:bg-emerald-900/30', text: 'text-emerald-700 dark:text-emerald-300' };
    case 'disqualified':
      return { bg: 'bg-slate-100 dark:bg-slate-800/50', text: 'text-slate-600 dark:text-slate-400' };
    default:
      return { bg: 'bg-muted', text: 'text-muted-foreground' };
  }
};

// Deal stage colors - using CSS variables from index.css
export const getDealStageColor = (stage: string): { bg: string; text: string } => {
  const stageLower = stage?.toLowerCase() || '';
  
  switch (stageLower) {
    case 'lead':
      return { bg: 'bg-[hsl(var(--stage-lead))]', text: 'text-[hsl(var(--stage-lead-foreground))]' };
    case 'discussions':
      return { bg: 'bg-[hsl(var(--stage-discussions))]', text: 'text-[hsl(var(--stage-discussions-foreground))]' };
    case 'qualified':
      return { bg: 'bg-[hsl(var(--stage-qualified))]', text: 'text-[hsl(var(--stage-qualified-foreground))]' };
    case 'rfq':
      return { bg: 'bg-[hsl(var(--stage-rfq))]', text: 'text-[hsl(var(--stage-rfq-foreground))]' };
    case 'offered':
      return { bg: 'bg-[hsl(var(--stage-offered))]', text: 'text-[hsl(var(--stage-offered-foreground))]' };
    case 'won':
      return { bg: 'bg-[hsl(var(--stage-won))]', text: 'text-[hsl(var(--stage-won-foreground))]' };
    case 'lost':
      return { bg: 'bg-[hsl(var(--stage-lost))]', text: 'text-[hsl(var(--stage-lost-foreground))]' };
    case 'dropped':
      return { bg: 'bg-[hsl(var(--stage-dropped))]', text: 'text-[hsl(var(--stage-dropped-foreground))]' };
    default:
      return { bg: 'bg-muted', text: 'text-muted-foreground' };
  }
};

// Meeting status colors
export const getMeetingStatusColor = (status: string): { bg: string; text: string } => {
  const statusLower = status?.toLowerCase() || '';
  
  switch (statusLower) {
    case 'scheduled':
      return { bg: 'bg-blue-100 dark:bg-blue-900/30', text: 'text-blue-700 dark:text-blue-300' };
    case 'completed':
      return { bg: 'bg-emerald-100 dark:bg-emerald-900/30', text: 'text-emerald-700 dark:text-emerald-300' };
    case 'cancelled':
      return { bg: 'bg-red-100 dark:bg-red-900/30', text: 'text-red-700 dark:text-red-300' };
    case 'rescheduled':
      return { bg: 'bg-amber-100 dark:bg-amber-900/30', text: 'text-amber-700 dark:text-amber-300' };
    case 'no-show':
      return { bg: 'bg-slate-100 dark:bg-slate-800/50', text: 'text-slate-600 dark:text-slate-400' };
    default:
      return { bg: 'bg-muted', text: 'text-muted-foreground' };
  }
};

// Task status colors
export const getTaskStatusColor = (status: string): { bg: string; text: string } => {
  const statusLower = status?.toLowerCase() || '';
  
  switch (statusLower) {
    case 'open':
      return { bg: 'bg-blue-100 dark:bg-blue-900/30', text: 'text-blue-700 dark:text-blue-300' };
    case 'in_progress':
    case 'in progress':
      return { bg: 'bg-amber-100 dark:bg-amber-900/30', text: 'text-amber-700 dark:text-amber-300' };
    case 'completed':
      return { bg: 'bg-emerald-100 dark:bg-emerald-900/30', text: 'text-emerald-700 dark:text-emerald-300' };
    case 'deferred':
      return { bg: 'bg-slate-100 dark:bg-slate-800/50', text: 'text-slate-600 dark:text-slate-400' };
    default:
      return { bg: 'bg-muted', text: 'text-muted-foreground' };
  }
};

// Task priority colors
export const getTaskPriorityColor = (priority: string): { bg: string; text: string } => {
  const priorityLower = priority?.toLowerCase() || '';
  
  switch (priorityLower) {
    case 'high':
      return { bg: 'bg-red-100 dark:bg-red-900/30', text: 'text-red-700 dark:text-red-300' };
    case 'medium':
      return { bg: 'bg-amber-100 dark:bg-amber-900/30', text: 'text-amber-700 dark:text-amber-300' };
    case 'low':
      return { bg: 'bg-emerald-100 dark:bg-emerald-900/30', text: 'text-emerald-700 dark:text-emerald-300' };
    default:
      return { bg: 'bg-muted', text: 'text-muted-foreground' };
  }
};

// Account status colors
export const getAccountStatusColor = (status: string): { bg: string; text: string } => {
  const statusLower = status?.toLowerCase() || '';
  
  switch (statusLower) {
    case 'new':
      return { bg: 'bg-blue-100 dark:bg-blue-900/30', text: 'text-blue-700 dark:text-blue-300' };
    case 'working':
      return { bg: 'bg-purple-100 dark:bg-purple-900/30', text: 'text-purple-700 dark:text-purple-300' };
    case 'warm':
      return { bg: 'bg-amber-100 dark:bg-amber-900/30', text: 'text-amber-700 dark:text-amber-300' };
    case 'hot':
      return { bg: 'bg-orange-100 dark:bg-orange-900/30', text: 'text-orange-700 dark:text-orange-300' };
    case 'nurture':
      return { bg: 'bg-cyan-100 dark:bg-cyan-900/30', text: 'text-cyan-700 dark:text-cyan-300' };
    case 'closed-won':
      return { bg: 'bg-emerald-100 dark:bg-emerald-900/30', text: 'text-emerald-700 dark:text-emerald-300' };
    case 'closed-lost':
      return { bg: 'bg-slate-100 dark:bg-slate-800/50', text: 'text-slate-600 dark:text-slate-400' };
    default:
      return { bg: 'bg-muted', text: 'text-muted-foreground' };
  }
};

// Generic status badge class generator
export const getStatusBadgeClasses = (colors: { bg: string; text: string; border?: string }): string => {
  return `${colors.bg} ${colors.text} ${colors.border || ''} px-2 py-0.5 rounded text-xs font-medium`.trim();
};
