import pptxgen from "pptxgenjs";
import { HasilKogol } from "./parserKogol";

export const downloadPresentationPptx = (
  data: HasilKogol,
  mode: "BULANAN" | "TAHUNAN",
  allHistory?: Record<string, HasilKogol>
) => {
  const pres = new pptxgen();

  // Dimensi paten Full HD Widescreen 16:9
  pres.defineLayout({ name: "FULL_HD", width: 13.33, height: 7.5 });
  pres.layout = "FULL_HD";

  // Corporate Color Palette (PLN & Modern Executive)
  const C_NAVY_DARK = "0B192C";
  const C_PLN_BLUE = "005C8A";
  const C_CYAN = "00A3E0";
  const C_ORANGE = "EA580C";
  const C_YELLOW = "FFD100";
  const C_WHITE = "FFFFFF";
  const C_CARD_BG = "FFFFFF";
  const C_CARD_BORDER = "CBD5E1";
  const C_TEXT_MAIN = "0F172A";
  const C_TEXT_MUTED = "64748B";
  const C_TEXT_LIGHT = "E2E8F0";

  const targetSaldo = 982000000;
  const persenAmr = ((data.denganKcic.amrRp / data.denganKcic.totalRp) * 100).toFixed(2);
  const capaianPersen = (((targetSaldo - data.denganKcic.totalRp) / targetSaldo) * 100).toFixed(1);

  // ==============================================================
  // SLIDE 1: COVER EKSEKUTIF BERKELAS
  // ==============================================================
  const slide1 = pres.addSlide();
  slide1.background = { color: C_NAVY_DARK };

  // Aksen garis kuning PLN & biru di kiri
  slide1.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 0.35, h: 7.5,
    fill: { color: C_YELLOW }, line: { color: C_YELLOW },
  });
  slide1.addShape(pres.ShapeType.rect, {
    x: 0.35, y: 0, w: 0.35, h: 7.5,
    fill: { color: C_CYAN }, line: { color: C_CYAN },
  });

  // Badge Kategori Slide
  slide1.addShape(pres.ShapeType.roundRect, {
    x: 1.5, y: 1.5, w: 3.4, h: 0.45,
    fill: { color: "1E293B" }, line: { color: C_CYAN, width: 1 },
  });
  slide1.addText(`LAPORAN KINERJA ${mode} RESMI`, {
    x: 1.5, y: 1.55, w: 3.4, h: 0.35,
    fontSize: 10, bold: true, color: C_CYAN, align: "center", fontFace: "Arial",
  });

  slide1.addText("⚡ SATU JATINEGARA", {
    x: 1.5, y: 2.1, w: 10.5, h: 1.1,
    fontSize: 44, bold: true, color: C_WHITE, fontFace: "Arial",
  });

  slide1.addText("Sistem Monitoring & Otomasi Analisis Saldo Akhir Tunggakan KOGOL", {
    x: 1.5, y: 3.2, w: 10.5, h: 0.5,
    fontSize: 19, color: C_YELLOW, fontFace: "Arial", bold: true,
  });

  slide1.addShape(pres.ShapeType.line, {
    x: 1.5, y: 3.9, w: 10.0, h: 0,
    line: { color: "334155", width: 1.5 },
  });

  slide1.addText(
    [
      { text: "Periode Cut-Off Data : ", options: { bold: true, color: C_TEXT_LIGHT } },
      { text: `${data.periode.toUpperCase()}\n`, options: { bold: true, color: C_CYAN } },
      { text: "Unit Kerja              : ", options: { bold: true, color: C_TEXT_LIGHT } },
      { text: "PT PLN (Persero) UP3 Jatinegara — UID Jakarta Raya\n", options: { color: C_TEXT_LIGHT } },
      { text: "Fokus Analisis          : ", options: { bold: true, color: C_TEXT_LIGHT } },
      { text: "Dekomposisi Saldo AMR vs Non-AMR & Evaluasi Piutang KCIC", options: { color: C_TEXT_LIGHT } },
    ],
    { x: 1.5, y: 4.2, w: 10.0, h: 2.3, fontSize: 13, fontFace: "Arial", lineSpacingMultiple: 1.3 }
  );

  // ==============================================================
  // SLIDE 2: DASHBOARD METRIK DUAL-PERSPECTIVE
  // ==============================================================
  const slide2 = pres.addSlide();
  slide2.background = { color: "F4F7F9" };

  // Header Nav Bar
  slide2.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 0.4, w: 11.73, h: 0.85,
    fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 },
  });
  slide2.addText("⚡ SATU JATINEGARA — KINERJA SALDO TUNGGAKAN", {
    x: 1.1, y: 0.48, w: 6.5, h: 0.35,
    fontSize: 15, bold: true, color: C_PLN_BLUE, fontFace: "Arial",
  });
  slide2.addText(`CUT-OFF: ${data.periode.toUpperCase()}  |  PT PLN (PERSERO) UP3 JATINEGARA`, {
    x: 1.1, y: 0.83, w: 6.5, h: 0.25,
    fontSize: 9.5, bold: true, color: C_TEXT_MUTED, fontFace: "Arial",
  });

  // KPI Target Card Kanan
  slide2.addShape(pres.ShapeType.roundRect, {
    x: 8.8, y: 0.48, w: 3.5, h: 0.7,
    fill: { color: C_NAVY_DARK }, line: { color: C_NAVY_DARK },
  });
  slide2.addText(`CAPAIAN TARGET: ${capaianPersen}%`, {
    x: 8.8, y: 0.52, w: 3.5, h: 0.3,
    fontSize: 12, bold: true, color: "F87171", align: "center", fontFace: "Arial",
  });
  slide2.addText(`Target: Rp 982 Jt | Realisasi: Rp ${(data.denganKcic.totalRp / 1e9).toFixed(2)} M`, {
    x: 8.8, y: 0.82, w: 3.5, h: 0.25,
    fontSize: 9, color: "BAE6FD", align: "center", fontFace: "Arial",
  });

  // KARTU 1: DENGAN KCIC
  slide2.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 1.45, w: 5.7, h: 2.15,
    fill: { color: C_CARD_BG }, line: { color: C_PLN_BLUE, width: 2 },
  });
  slide2.addShape(pres.ShapeType.rect, {
    x: 0.8, y: 1.45, w: 5.7, h: 0.32,
    fill: { color: "E0F2FE" }, line: { color: "E0F2FE" },
  });
  slide2.addText("PERSPEKTIF 1: TOTAL REALISASI (DENGAN KCIC)", {
    x: 1.0, y: 1.5, w: 5.3, h: 0.25,
    fontSize: 10, bold: true, color: C_PLN_BLUE, fontFace: "Arial",
  });

  slide2.addText(
    [
      { text: "TOTAL PELANGGAN\n", options: { fontSize: 8.5, color: C_TEXT_MUTED, bold: true } },
      { text: `${data.denganKcic.totalPlgn.toLocaleString("id-ID")}`, options: { fontSize: 22, bold: true, color: C_TEXT_MAIN } },
      { text: " Plgn\n", options: { fontSize: 10, color: C_TEXT_MUTED } },
      { text: `AMR: ${data.denganKcic.amrPlgn}  |  Non-AMR: ${data.denganKcic.nonAmrPlgn}`, options: { fontSize: 9.5, color: C_TEXT_MUTED } },
    ],
    { x: 1.0, y: 1.85, w: 2.6, h: 1.6, fontFace: "Arial" }
  );

  slide2.addText(
    [
      { text: "RUPIAH SALDO AKHIR\n", options: { fontSize: 8.5, color: C_TEXT_MUTED, bold: true } },
      { text: `Rp ${data.denganKcic.totalRp.toLocaleString("id-ID")}\n`, options: { fontSize: 15, bold: true, color: C_PLN_BLUE } },
      { text: `AMR : Rp ${data.denganKcic.amrRp.toLocaleString("id-ID")}\nNon : Rp ${data.denganKcic.nonAmrRp.toLocaleString("id-ID")}`, options: { fontSize: 8.5, color: C_TEXT_MUTED } },
    ],
    { x: 3.6, y: 1.85, w: 2.8, h: 1.6, fontFace: "Arial" }
  );

  // KARTU 2: TANPA KCIC
  slide2.addShape(pres.ShapeType.roundRect, {
    x: 6.83, y: 1.45, w: 5.7, h: 2.15,
    fill: { color: C_CARD_BG }, line: { color: C_ORANGE, width: 2 },
  });
  slide2.addShape(pres.ShapeType.rect, {
    x: 6.83, y: 1.45, w: 5.7, h: 0.32,
    fill: { color: "FFEDD5" }, line: { color: "FFEDD5" },
  });
  slide2.addText("PERSPEKTIF 2: BEBAN REGULER (TANPA KCIC)", {
    x: 7.03, y: 1.5, w: 5.3, h: 0.25,
    fontSize: 10, bold: true, color: C_ORANGE, fontFace: "Arial",
  });

  slide2.addText(
    [
      { text: "TOTAL PELANGGAN\n", options: { fontSize: 8.5, color: C_TEXT_MUTED, bold: true } },
      { text: `${data.tanpaKcic.totalPlgn.toLocaleString("id-ID")}`, options: { fontSize: 22, bold: true, color: C_TEXT_MAIN } },
      { text: " Plgn\n", options: { fontSize: 10, color: C_TEXT_MUTED } },
      { text: `AMR: ${data.tanpaKcic.amrPlgn}  |  Non-AMR: ${data.tanpaKcic.nonAmrPlgn}`, options: { fontSize: 9.5, color: C_TEXT_MUTED } },
    ],
    { x: 7.03, y: 1.85, w: 2.6, h: 1.6, fontFace: "Arial" }
  );

  slide2.addText(
    [
      { text: "RUPIAH SALDO AKHIR\n", options: { fontSize: 8.5, color: C_TEXT_MUTED, bold: true } },
      { text: `Rp ${data.tanpaKcic.totalRp.toLocaleString("id-ID")}\n`, options: { fontSize: 15, bold: true, color: C_ORANGE } },
      { text: `AMR : Rp ${data.tanpaKcic.amrRp.toLocaleString("id-ID")}\nNon : Rp ${data.tanpaKcic.nonAmrRp.toLocaleString("id-ID")}`, options: { fontSize: 8.5, color: C_TEXT_MUTED } },
    ],
    { x: 9.63, y: 1.85, w: 2.8, h: 1.6, fontFace: "Arial" }
  );

  // Tabel Rincian Data Kiri Bawah
  const tabelData = [
    [
      { text: "Kategori Segmentasi", options: { bold: true, fill: "F1F5F9" } },
      { text: "Tipe", options: { bold: true, fill: "F1F5F9" } },
      { text: "Pelanggan", options: { bold: true, fill: "F1F5F9", align: "center" } },
      { text: "Rupiah Tunggakan", options: { bold: true, fill: "F1F5F9", align: "right" } },
    ],
    ["Dengan KCIC", "AMR", `${data.denganKcic.amrPlgn} Plgn`, `Rp ${data.denganKcic.amrRp.toLocaleString("id-ID")}`],
    ["Dengan KCIC", "NON-AMR", `${data.denganKcic.nonAmrPlgn} Plgn`, `Rp ${data.denganKcic.nonAmrRp.toLocaleString("id-ID")}`],
    ["Tanpa KCIC", "AMR", `${data.tanpaKcic.amrPlgn} Plgn`, `Rp ${data.tanpaKcic.amrRp.toLocaleString("id-ID")}`],
    ["Tanpa KCIC", "NON-AMR", `${data.tanpaKcic.nonAmrPlgn} Plgn`, `Rp ${data.tanpaKcic.nonAmrRp.toLocaleString("id-ID")}`],
    ["Khusus Entitas KCIC", "AMR", `${data.kcicOnly.plgn} Plgn`, `Rp ${data.kcicOnly.rp.toLocaleString("id-ID")}`],
  ];

  slide2.addTable(tabelData, {
    x: 0.8, y: 3.8, w: 6.8,
    fontSize: 9,
    border: { pt: 0.5, color: C_CARD_BORDER },
    fill: { color: C_WHITE },
    color: C_TEXT_MAIN,
    autoPage: false,
  });

  // Boks Rekomendasi Eksekutif Kanan Bawah
  slide2.addShape(pres.ShapeType.roundRect, {
    x: 7.8, y: 3.8, w: 4.73, h: 3.2,
    fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 },
  });
  slide2.addText("RINGKASAN & STRATEGI EKSEKUTIF", {
    x: 8.0, y: 3.95, w: 4.3, h: 0.3,
    fontSize: 10, bold: true, color: C_TEXT_MAIN, fontFace: "Arial",
  });

  slide2.addText(
    `📌 Total Realisasi Saldo: Rp ${(data.denganKcic.totalRp / 1e9).toFixed(2)} Miliar terhadap target Rp ${(targetSaldo / 1e6).toFixed(0)} Juta (Deviasi: ${capaianPersen}%).\n\n` +
    `⚠️ Konsentrasi Risiko AMR: Sebesar ${persenAmr}% nominal saldo berpusat pada kategori AMR, memerlukan penagihan intensif dan koordinasi terpusat.\n\n` +
    `🚆 Isolasi Piutang KCIC: Entitas KCIC membukukan saldo Rp ${data.kcicOnly.rp.toLocaleString("id-ID")} (${data.kcicOnly.plgn} IDPEL). Bila piutang KCIC dikeluarkan, saldo tunggakan murni unit berada di posisi aman Rp ${(data.tanpaKcic.totalRp / 1e6).toFixed(1)} Juta.`,
    {
      x: 8.0, y: 4.3, w: 4.3, h: 2.5,
      fontSize: 9.5, color: C_TEXT_MAIN, lineSpacingMultiple: 1.25, fontFace: "Arial",
    }
  );

  // ==============================================================
  // SLIDE 3: ANALISIS STRUKTURAL (DONUT PELANGGAN & NOMINAL)
  // ==============================================================
  const slide3 = pres.addSlide();
  slide3.background = { color: "F4F7F9" };

  slide3.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 0.4, w: 11.73, h: 0.75,
    fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 },
  });
  slide3.addText(`ANALISIS STRUKTURAL: PROPORSI AMR VS NON-AMR (${data.periode.toUpperCase()})`, {
    x: 1.1, y: 0.6, w: 11.0, h: 0.35,
    fontSize: 15, bold: true, color: C_PLN_BLUE, fontFace: "Arial",
  });

  // Donut 1
  slide3.addShape(pres.ShapeType.roundRect, {
    x: 0.8, y: 1.4, w: 5.7, h: 5.4,
    fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 },
  });
  slide3.addChart(
    pres.ChartType.doughnut,
    [
      {
        name: "Pelanggan",
        labels: ["AMR", "NON-AMR"],
        values: [data.denganKcic.amrPlgn, data.denganKcic.nonAmrPlgn],
      },
    ],
    {
      x: 1.0, y: 1.6, w: 5.3, h: 5.0,
      chartColors: [C_PLN_BLUE, C_YELLOW],
      showPercent: true,
      showLegend: true,
      legendPos: "b",
      title: "Proporsi Berdasarkan Kuantitas Pelanggan",
    }
  );

  // Donut 2
  slide3.addShape(pres.ShapeType.roundRect, {
    x: 6.83, y: 1.4, w: 5.7, h: 5.4,
    fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 },
  });
  slide3.addChart(
    pres.ChartType.doughnut,
    [
      {
        name: "Nominal Tagihan",
        labels: ["AMR", "NON-AMR"],
        values: [data.denganKcic.amrRp, data.denganKcic.nonAmrRp],
      },
    ],
    {
      x: 7.03, y: 1.6, w: 5.3, h: 5.0,
      chartColors: [C_PLN_BLUE, C_ORANGE],
      showPercent: true,
      showLegend: true,
      legendPos: "b",
      title: "Proporsi Berdasarkan Nominal Rupiah Tagihan",
    }
  );

  // ==============================================================
  // SLIDE 4: TOP 10 PELANGGAN TUNGGAKAN TERBESAR (PARETO)
  // ==============================================================
  if (data.semuaPelanggan && data.semuaPelanggan.length > 0) {
    const slide4 = pres.addSlide();
    slide4.background = { color: "F4F7F9" };

    slide4.addShape(pres.ShapeType.roundRect, {
      x: 0.8, y: 0.4, w: 11.73, h: 0.75,
      fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 },
    });
    slide4.addText(`PRIORITAS PENAGIHAN: TOP 10 PELANGGAN PARETO — ${data.periode.toUpperCase()}`, {
      x: 1.1, y: 0.6, w: 11.0, h: 0.35,
      fontSize: 15, bold: true, color: C_PLN_BLUE, fontFace: "Arial",
    });

    const top10Rows = data.semuaPelanggan.slice(0, 10);
    const tableTop10Data: any[] = [
      [
        { text: "No", options: { bold: true, fill: "F1F5F9", align: "center" } },
        { text: "ID Pelanggan", options: { bold: true, fill: "F1F5F9" } },
        { text: "Nama Pelanggan", options: { bold: true, fill: "F1F5F9" } },
        { text: "Tarif / Daya", options: { bold: true, fill: "F1F5F9" } },
        { text: "Segmen", options: { bold: true, fill: "F1F5F9", align: "center" } },
        { text: "Nominal Tunggakan (Rp)", options: { bold: true, fill: "F1F5F9", align: "right" } },
        { text: "Kontribusi", options: { bold: true, fill: "F1F5F9", align: "right" } },
      ],
    ];

    top10Rows.forEach((p, idx) => {
      const kontribusi = ((p.rpTunggakan / data.denganKcic.totalRp) * 100).toFixed(2);
      tableTop10Data.push([
        `${idx + 1}`,
        p.idpel,
        p.nama,
        p.tarifDaya,
        p.kategori,
        `Rp ${p.rpTunggakan.toLocaleString("id-ID")}`,
        `${kontribusi}%`,
      ]);
    });

    slide4.addTable(tableTop10Data, {
      x: 0.8, y: 1.45, w: 11.73,
      fontSize: 9.5,
      border: { pt: 0.5, color: C_CARD_BORDER },
      fill: { color: C_WHITE },
      color: C_TEXT_MAIN,
      autoPage: false,
    });
  }

  // ==============================================================
  // SLIDE 5: KOMPARASI MULTI-BULAN (Jika Ada > 1 Bulan)
  // ==============================================================
  if (allHistory && Object.keys(allHistory).length > 1) {
    const slide5 = pres.addSlide();
    slide5.background = { color: "F4F7F9" };

    slide5.addShape(pres.ShapeType.roundRect, {
      x: 0.8, y: 0.4, w: 11.73, h: 0.75,
      fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 },
    });
    slide5.addText("TREN KOMPARASI REALISASI TUNGGAKAN MULTI-PERIODE", {
      x: 1.1, y: 0.6, w: 11.0, h: 0.35,
      fontSize: 15, bold: true, color: C_PLN_BLUE, fontFace: "Arial",
    });

    const sortedKeys = Object.keys(allHistory).sort();
    const chartLabels = sortedKeys.map((k) => allHistory[k].periode);
    const dataDengan = sortedKeys.map((k) => allHistory[k].denganKcic.totalRp);
    const dataTanpa = sortedKeys.map((k) => allHistory[k].tanpaKcic.totalRp);

    slide5.addShape(pres.ShapeType.roundRect, {
      x: 0.8, y: 1.4, w: 11.73, h: 5.4,
      fill: { color: C_CARD_BG }, line: { color: C_CARD_BORDER, width: 1 },
    });

    slide5.addChart(
      pres.ChartType.bar,
      [
        { name: "Dengan KCIC", labels: chartLabels, values: dataDengan },
        { name: "Tanpa KCIC", labels: chartLabels, values: dataTanpa },
      ],
      {
        x: 1.1, y: 1.6, w: 11.13, h: 5.0,
        chartColors: [C_ORANGE, C_PLN_BLUE],
        showLegend: true,
        legendPos: "t",
        title: "Perbandingan Nominal Saldo Tunggakan Antar-Bulan (Rupiah)",
      }
    );
  }

  pres.writeFile({
    fileName: `SATU_Jatinegara_${mode}_${data.periode.replace(/\s+/g, "_")}.pptx`,
  });
};