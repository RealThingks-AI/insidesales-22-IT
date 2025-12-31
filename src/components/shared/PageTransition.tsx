import { ReactNode } from "react";
import { cn } from "@/lib/utils";

interface PageTransitionProps {
  children: ReactNode;
  className?: string;
}

/**
 * Page transition wrapper that applies a subtle fade-in animation
 * Use this wrapper around page content for consistent transitions
 */
export const PageTransition = ({ children, className }: PageTransitionProps) => {
  return (
    <div className={cn("animate-fade-in", className)}>
      {children}
    </div>
  );
};

export default PageTransition;
