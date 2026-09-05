import * as React from "react";
import { cva, type VariantProps } from "class-variance-authority";
import { cn } from "@/lib/utils";

const buttonVariants = cva(
  "inline-flex items-center justify-center whitespace-nowrap rounded-lg text-sm font-medium transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500 disabled:pointer-events-none disabled:opacity-50 active:scale-[0.98]",
  {
    variants: {
      variant: {
        default:
          "bg-gradient-to-r from-emerald-600 to-teal-600 text-white shadow-lg shadow-emerald-500/20 hover:from-emerald-500 hover:to-teal-500",
        destructive:
          "bg-rose-600 text-white hover:bg-rose-500 shadow-md shadow-rose-500/20",
        outline:
          "border border-slate-700 bg-slate-800/50 text-slate-200 hover:bg-slate-700/60 hover:text-white hover:border-slate-600",
        secondary:
          "bg-slate-800 text-slate-100 hover:bg-slate-700",
        ghost:
          "text-slate-300 hover:bg-slate-800/80 hover:text-white",
        gradientSuccess:
          "bg-gradient-to-r from-emerald-600 to-teal-500 text-white shadow-lg shadow-emerald-500/20 hover:from-emerald-500 hover:to-teal-400",
      },
      size: {
        default: "h-10 px-4 py-2",
        sm: "h-8 rounded-md px-3 text-xs",
        lg: "h-12 rounded-xl px-8 text-base",
        icon: "h-9 w-9 p-0",
      },
    },
    defaultVariants: {
      variant: "default",
      size: "default",
    },
  }
);

export interface ButtonProps
  extends React.ButtonHTMLAttributes<HTMLButtonElement>,
    VariantProps<typeof buttonVariants> {}

const Button = React.forwardRef<HTMLButtonElement, ButtonProps>(
  ({ className, variant, size, ...props }, ref) => {
    return (
      <button
        className={cn(buttonVariants({ variant, size, className }))}
        ref={ref}
        {...props}
      />
    );
  }
);
Button.displayName = "Button";

export { Button, buttonVariants };
