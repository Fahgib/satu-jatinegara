"use client";

import React, { useEffect, useRef } from "react";
import gsap from "gsap";
import { Users, Receipt } from "lucide-react";

interface Props {
  title: string;
  totalPlgn: number;
  amrPlgn: number;
  nonAmrPlgn: number;
  totalRp: number;
  amrRp: number;
  nonAmrRp: number;
  theme: "blue" | "orange";
}

export const MetricCard: React.FC<Props> = ({
  title,
  totalPlgn,
  amrPlgn,
  nonAmrPlgn,
  totalRp,
  amrRp,
  nonAmrRp,
  theme,
}) => {
  const plgnRef = useRef<HTMLSpanElement>(null);
  const rpRef = useRef<HTMLSpanElement>(null);

  useEffect(() => {
    const countObj = { val: 0 };
    gsap.to(countObj, {
      val: totalPlgn,
      duration: 1.5,
      ease: "power2.out",
      onUpdate: () => {
        if (plgnRef.current) plgnRef.current.innerText = Math.floor(countObj.val).toLocaleString("id-ID");
      },
    });

    const rpObj = { val: 0 };
    gsap.to(rpObj, {
      val: totalRp,
      duration: 1.8,
      ease: "power3.out",
      onUpdate: () => {
        if (rpRef.current) rpRef.current.innerText = `Rp ${Math.floor(rpObj.val).toLocaleString("id-ID")}`;
      },
    });
  }, [totalPlgn, totalRp]);

  const isBlue = theme === "blue";

  return (
    <div className={`bg-white rounded-xl shadow-xs border border-slate-200 overflow-hidden border-t-4 ${isBlue ? "border-t-[#005C8A]" : "border-t-[#EA580C]"}`}>
      <div className={`px-5 py-2.5 flex justify-between items-center ${isBlue ? "bg-sky-50 text-[#005C8A]" : "bg-orange-50 text-[#EA580C]"}`}>
        <span className="font-extrabold text-xs tracking-wider uppercase">{title}</span>
        <span className="text-[10px] bg-white px-2 py-0.5 rounded shadow-xs font-semibold">Real-time</span>
      </div>

      <div className="p-5 grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="border-r border-slate-100 pr-4">
          <div className="flex items-center gap-1.5 text-slate-500 text-xs font-semibold uppercase tracking-wider mb-1">
            <Users className="w-4 h-4" />
            <span>Total Pelanggan</span>
          </div>
          <div className="text-3xl font-black text-slate-800">
            <span ref={plgnRef}>0</span> <span className="text-xs text-slate-400 font-medium">Plgn</span>
          </div>
          <div className="mt-2 text-xs text-slate-500">
            AMR: <strong className="text-slate-700">{amrPlgn}</strong> | Non-AMR: <strong className="text-slate-700">{nonAmrPlgn}</strong>
          </div>
        </div>

        <div className="pl-2">
          <div className="flex items-center gap-1.5 text-slate-500 text-xs font-semibold uppercase tracking-wider mb-1">
            <Receipt className="w-4 h-4" />
            <span>Rupiah Tagihan</span>
          </div>
          <div className={`text-2xl font-black ${isBlue ? "text-[#005C8A]" : "text-[#EA580C]"}`}>
            <span ref={rpRef}>Rp 0</span>
          </div>
          <div className="mt-2 text-xs text-slate-500 flex flex-col gap-0.5">
            <span>AMR: <strong>Rp {amrRp.toLocaleString("id-ID")}</strong></span>
            <span>Non-AMR: <strong>Rp {nonAmrRp.toLocaleString("id-ID")}</strong></span>
          </div>
        </div>
      </div>
    </div>
  );
};