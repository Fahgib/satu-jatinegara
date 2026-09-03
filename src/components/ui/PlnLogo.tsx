import React from "react";

export const PlnLogo: React.FC<{ className?: string }> = ({ className = "w-9 h-11" }) => {
  return (
    <svg
      viewBox="0 0 100 125"
      className={className}
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      {/* Bingkai Luar Kuning Emas */}
      <path
        d="M10 8 H90 L80 115 L20 115 Z"
        fill="#FFDD00"
        stroke="#F59E0B"
        strokeWidth="2"
      />
      {/* Tiga Gelombang Air Biru */}
      <path
        d="M24 78 C32 74, 40 82, 48 78 C56 74, 64 82, 74 78"
        stroke="#005C8A"
        strokeWidth="4.5"
        strokeLinecap="round"
        fill="none"
      />
      <path
        d="M23 88 C31 84, 39 92, 47 88 C55 84, 63 92, 75 88"
        stroke="#005C8A"
        strokeWidth="4.5"
        strokeLinecap="round"
        fill="none"
      />
      <path
        d="M22 98 C30 94, 38 102, 46 98 C54 94, 62 102, 76 98"
        stroke="#005C8A"
        strokeWidth="4.5"
        strokeLinecap="round"
        fill="none"
      />
      {/* Kilat Petir Merah */}
      <path
        d="M58 14 L32 54 L52 54 L42 82 L72 42 L52 42 Z"
        fill="#E11D48"
        stroke="#BE123C"
        strokeWidth="1.5"
        strokeLinejoin="round"
      />
    </svg>
  );
};