import React from "react";
import { cn } from "@/lib/utils";

export const BentoCard = ({
  children,
  className,
  title,
  subtitle,
  badge,
  icon,
}: {
  children: React.ReactNode;
  className?: string;
  title?: string;
  subtitle?: string;
  badge?: React.ReactNode;
  icon?: React.ReactNode;
}) => {
  return (
    <div
      className={cn(
        "group relative rounded-2xl border border-slate-800/80 bg-slate-900/60 p-6 backdrop-blur-xl shadow-xl transition-all duration-300 hover:border-sky-500/40 hover:shadow-sky-500/5",
        className
      )}
    >
      <div className="absolute inset-0 -z-10 rounded-2xl bg-gradient-to-b from-white/[0.04] to-transparent pointer-events-none" />
      
      {(title || badge || icon) && (
        <div className="flex items-center justify-between gap-3 mb-4 pb-3 border-b border-slate-800/80">
          <div className="flex items-center gap-2.5">
            {icon && <div className="p-2 rounded-xl bg-slate-800/80 text-sky-400 border border-slate-700/60">{icon}</div>}
            <div>
              {title && <h4 className="text-xs font-black uppercase tracking-widest text-slate-200">{title}</h4>}
              {subtitle && <p className="text-[11px] text-slate-400 mt-0.5">{subtitle}</p>}
            </div>
          </div>
          {badge}
        </div>
      )}
      {children}
    </div>
  );
};