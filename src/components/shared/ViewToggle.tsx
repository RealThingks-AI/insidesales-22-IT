import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";
import { LucideIcon } from "lucide-react";

interface ViewOption {
  id: string;
  label: string;
  icon: LucideIcon;
}

interface ViewToggleProps {
  views: ViewOption[];
  activeView: string;
  onViewChange: (view: string) => void;
  className?: string;
}

export const ViewToggle = ({
  views,
  activeView,
  onViewChange,
  className,
}: ViewToggleProps) => {
  return (
    <div className={cn("flex items-center gap-0.5 bg-muted rounded-md p-0.5", className)}>
      {views.map((view) => {
        const Icon = view.icon;
        return (
          <Button
            key={view.id}
            variant={activeView === view.id ? "secondary" : "ghost"}
            size="sm"
            onClick={() => onViewChange(view.id)}
            className="gap-1.5 h-8 px-2.5 text-xs"
            aria-pressed={activeView === view.id}
          >
            <Icon className="h-3.5 w-3.5" />
            {view.label}
          </Button>
        );
      })}
    </div>
  );
};
