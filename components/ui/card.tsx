import * as React from "react";
import { cn } from "@/lib/utils";

const Card = React.forwardRef<HTMLDivElement, React.HTMLAttributes<HTMLDivElement>>(
  ({ className, ...props }, ref) => (
    <div
      ref={ref}
      className={cn(
        "rounded-2xl border border-slate-800 bg-slate-900/80 text-slate-100 shadow-xl backdrop-blur-md transition-all duration-200 hover:border-slate-700/80",
        className
      )}
      {...props}
    />
  )
);
Card.displayName = "Card";

export { Card };
