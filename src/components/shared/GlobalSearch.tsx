import { useState, useRef, useEffect } from "react";
import { Search, Users, FileText, Briefcase, Calendar, ListTodo, Settings, Building2, X } from "lucide-react";
import { Input } from "@/components/ui/input";
import { useNavigate } from "react-router-dom";
import { supabase } from "@/integrations/supabase/client";
import { useQuery } from "@tanstack/react-query";
import { ScrollArea } from "@/components/ui/scroll-area";
import { cn } from "@/lib/utils";

interface SearchResult {
  id: string;
  type: 'lead' | 'contact' | 'deal' | 'meeting' | 'task' | 'account' | 'setting';
  title: string;
  subtitle?: string;
  route: string;
}

const settingsPages = [
  { id: 'profile', title: 'Profile Settings', subtitle: 'Manage your profile', route: '/settings?tab=profile' },
  { id: 'display', title: 'Display Settings', subtitle: 'Theme and appearance', route: '/settings?tab=display' },
  { id: 'notifications', title: 'Notification Settings', subtitle: 'Email and push notifications', route: '/settings?tab=notifications' },
  { id: 'security', title: 'Security Settings', subtitle: 'Password and authentication', route: '/settings?tab=security' },
  { id: 'integrations', title: 'Integration Settings', subtitle: 'Connect external services', route: '/settings?tab=integrations' },
  { id: 'pipeline', title: 'Pipeline Settings', subtitle: 'Manage deal stages', route: '/settings?tab=pipeline' },
  { id: 'email-templates', title: 'Email Templates', subtitle: 'Manage email templates', route: '/settings?tab=email-templates' },
];

export const GlobalSearch = () => {
  const [query, setQuery] = useState("");
  const [isOpen, setIsOpen] = useState(false);
  const [selectedIndex, setSelectedIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const navigate = useNavigate();

  // Fetch all searchable data
  const { data: leads = [] } = useQuery({
    queryKey: ['global-search-leads'],
    queryFn: async () => {
      const { data } = await supabase.from('leads').select('id, lead_name, company_name, email').limit(100);
      return data || [];
    },
  });

  const { data: contacts = [] } = useQuery({
    queryKey: ['global-search-contacts'],
    queryFn: async () => {
      const { data } = await supabase.from('contacts').select('id, contact_name, company_name, email').limit(100);
      return data || [];
    },
  });

  const { data: deals = [] } = useQuery({
    queryKey: ['global-search-deals'],
    queryFn: async () => {
      const { data } = await supabase.from('deals').select('id, deal_name, customer_name, stage').limit(100);
      return data || [];
    },
  });

  const { data: meetings = [] } = useQuery({
    queryKey: ['global-search-meetings'],
    queryFn: async () => {
      const { data } = await supabase.from('meetings').select('id, subject, description, start_time').limit(100);
      return data || [];
    },
  });

  const { data: tasks = [] } = useQuery({
    queryKey: ['global-search-tasks'],
    queryFn: async () => {
      const { data } = await supabase.from('tasks').select('id, title, description, status').limit(100);
      return data || [];
    },
  });

  const { data: accounts = [] } = useQuery({
    queryKey: ['global-search-accounts'],
    queryFn: async () => {
      const { data } = await supabase.from('accounts').select('id, company_name, industry, email').limit(100);
      return data || [];
    },
  });

  // Filter results based on query
  const getSearchResults = (): SearchResult[] => {
    if (!query.trim()) return [];
    
    const lowerQuery = query.toLowerCase();
    const results: SearchResult[] = [];

    // Search leads
    leads.forEach(lead => {
      if (
        lead.lead_name?.toLowerCase().includes(lowerQuery) ||
        lead.company_name?.toLowerCase().includes(lowerQuery) ||
        lead.email?.toLowerCase().includes(lowerQuery)
      ) {
        results.push({
          id: lead.id,
          type: 'lead',
          title: lead.lead_name,
          subtitle: lead.company_name || lead.email,
          route: `/leads?highlight=${lead.id}`,
        });
      }
    });

    // Search contacts
    contacts.forEach(contact => {
      if (
        contact.contact_name?.toLowerCase().includes(lowerQuery) ||
        contact.company_name?.toLowerCase().includes(lowerQuery) ||
        contact.email?.toLowerCase().includes(lowerQuery)
      ) {
        results.push({
          id: contact.id,
          type: 'contact',
          title: contact.contact_name,
          subtitle: contact.company_name || contact.email,
          route: `/contacts?highlight=${contact.id}`,
        });
      }
    });

    // Search deals
    deals.forEach(deal => {
      if (
        deal.deal_name?.toLowerCase().includes(lowerQuery) ||
        deal.customer_name?.toLowerCase().includes(lowerQuery)
      ) {
        results.push({
          id: deal.id,
          type: 'deal',
          title: deal.deal_name,
          subtitle: deal.customer_name || deal.stage,
          route: `/deals?highlight=${deal.id}`,
        });
      }
    });

    // Search meetings
    meetings.forEach(meeting => {
      if (
        meeting.subject?.toLowerCase().includes(lowerQuery) ||
        meeting.description?.toLowerCase().includes(lowerQuery)
      ) {
        results.push({
          id: meeting.id,
          type: 'meeting',
          title: meeting.subject,
          subtitle: meeting.start_time ? new Date(meeting.start_time).toLocaleDateString() : undefined,
          route: `/meetings?highlight=${meeting.id}`,
        });
      }
    });

    // Search tasks
    tasks.forEach(task => {
      if (
        task.title?.toLowerCase().includes(lowerQuery) ||
        task.description?.toLowerCase().includes(lowerQuery)
      ) {
        results.push({
          id: task.id,
          type: 'task',
          title: task.title,
          subtitle: task.status,
          route: `/tasks?highlight=${task.id}`,
        });
      }
    });

    // Search accounts
    accounts.forEach(account => {
      if (
        account.company_name?.toLowerCase().includes(lowerQuery) ||
        account.industry?.toLowerCase().includes(lowerQuery) ||
        account.email?.toLowerCase().includes(lowerQuery)
      ) {
        results.push({
          id: account.id,
          type: 'account',
          title: account.company_name,
          subtitle: account.industry || account.email,
          route: `/accounts?highlight=${account.id}`,
        });
      }
    });

    // Search settings
    settingsPages.forEach(setting => {
      if (
        setting.title.toLowerCase().includes(lowerQuery) ||
        setting.subtitle.toLowerCase().includes(lowerQuery)
      ) {
        results.push({
          id: setting.id,
          type: 'setting',
          title: setting.title,
          subtitle: setting.subtitle,
          route: setting.route,
        });
      }
    });

    return results.slice(0, 10); // Limit to 10 results
  };

  const results = getSearchResults();

  const getIcon = (type: SearchResult['type']) => {
    switch (type) {
      case 'lead': return <FileText className="w-4 h-4 text-blue-500" />;
      case 'contact': return <Users className="w-4 h-4 text-green-500" />;
      case 'deal': return <Briefcase className="w-4 h-4 text-purple-500" />;
      case 'meeting': return <Calendar className="w-4 h-4 text-orange-500" />;
      case 'task': return <ListTodo className="w-4 h-4 text-yellow-500" />;
      case 'account': return <Building2 className="w-4 h-4 text-cyan-500" />;
      case 'setting': return <Settings className="w-4 h-4 text-muted-foreground" />;
    }
  };

  const getTypeLabel = (type: SearchResult['type']) => {
    return type.charAt(0).toUpperCase() + type.slice(1);
  };

  const handleSelect = (result: SearchResult) => {
    navigate(result.route);
    setQuery("");
    setIsOpen(false);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (!isOpen || results.length === 0) return;

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        setSelectedIndex(prev => (prev + 1) % results.length);
        break;
      case 'ArrowUp':
        e.preventDefault();
        setSelectedIndex(prev => (prev - 1 + results.length) % results.length);
        break;
      case 'Enter':
        e.preventDefault();
        if (results[selectedIndex]) {
          handleSelect(results[selectedIndex]);
        }
        break;
      case 'Escape':
        setIsOpen(false);
        break;
    }
  };

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (
        dropdownRef.current && 
        !dropdownRef.current.contains(e.target as Node) &&
        inputRef.current &&
        !inputRef.current.contains(e.target as Node)
      ) {
        setIsOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Reset selected index when results change
  useEffect(() => {
    setSelectedIndex(0);
  }, [results.length]);

  return (
    <div className="relative w-full max-w-sm">
      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
        <Input
          ref={inputRef}
          type="text"
          placeholder="Search everything..."
          value={query}
          onChange={(e) => {
            setQuery(e.target.value);
            setIsOpen(true);
          }}
          onFocus={() => query && setIsOpen(true)}
          onKeyDown={handleKeyDown}
          className="pl-9 pr-8 h-9 bg-background border-border"
        />
        {query && (
          <button
            onClick={() => {
              setQuery("");
              setIsOpen(false);
              inputRef.current?.focus();
            }}
            className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
          >
            <X className="w-4 h-4" />
          </button>
        )}
      </div>

      {/* Results Dropdown */}
      {isOpen && query && (
        <div
          ref={dropdownRef}
          className="absolute top-full left-0 right-0 mt-1 bg-popover border border-border rounded-lg shadow-lg z-50 overflow-hidden"
        >
          {results.length > 0 ? (
            <ScrollArea className="max-h-80">
              <div className="p-1">
                {results.map((result, index) => (
                  <button
                    key={`${result.type}-${result.id}`}
                    onClick={() => handleSelect(result)}
                    className={cn(
                      "w-full flex items-center gap-3 px-3 py-2 text-left rounded-md transition-colors",
                      index === selectedIndex 
                        ? "bg-accent text-accent-foreground" 
                        : "hover:bg-muted"
                    )}
                  >
                    {getIcon(result.type)}
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium truncate">{result.title}</p>
                      {result.subtitle && (
                        <p className="text-xs text-muted-foreground truncate">{result.subtitle}</p>
                      )}
                    </div>
                    <span className="text-xs text-muted-foreground bg-muted px-2 py-0.5 rounded">
                      {getTypeLabel(result.type)}
                    </span>
                  </button>
                ))}
              </div>
            </ScrollArea>
          ) : (
            <div className="p-4 text-center text-sm text-muted-foreground">
              No results found for "{query}"
            </div>
          )}
        </div>
      )}
    </div>
  );
};
