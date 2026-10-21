const docx = require("docx");
const fs = require("fs");

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, HeadingLevel, BorderStyle, ShadingType,
  TableOfContents, PageBreak, Header, Footer, PageNumber,
} = docx;

const PAGE_WIDTH = 12240;
const MARGIN = 1440;
const CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN;

const COLORS = {
  primary: "1B3A5C",
  secondary: "2E75B6",
  accent: "4472C4",
  lightBg: "E8EFF7",
  darkText: "1F2937",
  grayText: "6B7280",
  white: "FFFFFF",
  border: "B4C7DC",
  headerBg: "1B3A5C",
  rowAlt: "F0F5FA",
  citation: "8B0000",
};

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORS.secondary } },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 28, color: COLORS.primary, font: "Calibri" })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [
      new TextRun({ text: "  ", size: 24 }),
      new TextRun({ text, bold: true, size: 24, color: COLORS.secondary, font: "Calibri" }),
    ],
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, bold: true, size: 22, color: COLORS.accent, font: "Calibri", italics: true })],
  });
}

function richParagraph(runs, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 276 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    indent: opts.indent ? { firstLine: 720 } : undefined,
    children: runs,
  });
}

function bodyText(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120, line: 276 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    indent: opts.indent ? { firstLine: 720 } : undefined,
    children: [new TextRun({ text, size: 22, color: COLORS.darkText, font: "Calibri", italics: opts.italic || false, bold: opts.bold || false })],
  });
}

function cite(tag) {
  return new TextRun({ text: " (" + tag + ")", size: 20, color: COLORS.citation, font: "Calibri", italics: true, bold: true });
}

function t(text, opts = {}) {
  return new TextRun({ text, size: opts.size || 22, color: opts.color || COLORS.darkText, font: "Calibri", italics: opts.italic || false, bold: opts.bold || false });
}

function infoBox(lines) {
  const cell = new TableCell({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: COLORS.lightBg },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
      left: { style: BorderStyle.SINGLE, size: 8, color: COLORS.secondary },
      right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
    },
    margins: { top: 100, bottom: 100, left: 200, right: 200 },
    children: lines.map((l) => new Paragraph({
      spacing: { after: 60 },
      children: [
        new TextRun({ text: l.label + ": ", bold: true, size: 20, color: COLORS.primary, font: "Calibri" }),
        new TextRun({ text: l.value, size: 20, color: COLORS.darkText, font: "Calibri" }),
      ],
    })),
  });
  return new Table({ rows: [new TableRow({ children: [cell] })], width: { size: CONTENT_WIDTH, type: WidthType.DXA } });
}

function lawBox(lines) {
  const cell = new TableCell({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: "FFF8E7" },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: "E0C97F" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "E0C97F" },
      left: { style: BorderStyle.SINGLE, size: 8, color: "C8961E" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "E0C97F" },
    },
    margins: { top: 100, bottom: 100, left: 200, right: 200 },
    children: lines.map((l) => new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: l, size: 20, color: "5C4A00", font: "Calibri", italics: true })],
    })),
  });
  return new Table({ rows: [new TableRow({ children: [cell] })], width: { size: CONTENT_WIDTH, type: WidthType.DXA } });
}

function styledTable(headers, rows) {
  const colCount = headers.length;
  const colWidth = Math.floor(CONTENT_WIDTH / colCount);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h) => new TableCell({
      width: { size: colWidth, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: COLORS.headerBg },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
        left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
        right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
      },
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, size: 20, color: COLORS.white, font: "Calibri" })] })],
    })),
  });
  const dataRows = rows.map((row, idx) => new TableRow({
    children: row.map((cell) => new TableCell({
      width: { size: colWidth, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: idx % 2 === 0 ? COLORS.white : COLORS.rowAlt },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
        left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
        right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
      },
      margins: { top: 40, bottom: 40, left: 100, right: 100 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cell, size: 20, color: COLORS.darkText, font: "Calibri" })] })],
    })),
  }));
  return new Table({ rows: [headerRow, ...dataRows], width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: Array(colCount).fill(colWidth) });
}

function emptyLine() { return new Paragraph({ spacing: { after: 80 } }); }
function separatorLine() {
  return new Paragraph({ spacing: { before: 200, after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: COLORS.border } }, children: [new TextRun({ text: "" })] });
}

// ============= DOCUMENT =============
const doc = new Document({
  creator: "Ekin Harita M\u00fchendislik ve M\u00fc\u015favirlik B\u00fcrosu",
  title: "Marmaraere\u011flisi OSB Parselasyon Plan\u0131 A\u00e7\u0131klama Raporu",
  description: "Tekirda\u011f ili, Marmaraere\u011flisi il\u00e7esi, Marmaraere\u011flisi OSB Parselasyon Plan\u0131 A\u00e7\u0131klama Raporu",
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 22, color: COLORS.darkText }, paragraph: { spacing: { line: 276 } } },
      heading1: { run: { font: "Calibri", size: 28, bold: true, color: COLORS.primary }, paragraph: { spacing: { before: 360, after: 200 } } },
      heading2: { run: { font: "Calibri", size: 24, bold: true, color: COLORS.secondary }, paragraph: { spacing: { before: 280, after: 160 } } },
      heading3: { run: { font: "Calibri", size: 22, bold: true, color: COLORS.accent, italics: true }, paragraph: { spacing: { before: 200, after: 120 } } },
    },
  },
  features: { updateFields: true },
  sections: [
    // ===== KAPAK =====
    {
      properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }, pageNumbers: { start: 1 } } },
      children: [
        emptyLine(), emptyLine(), emptyLine(), emptyLine(),
        new Paragraph({ spacing: { after: 400 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: COLORS.primary } }, children: [new TextRun({ text: "" })] }),
        emptyLine(),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: "T.C.", bold: true, size: 32, color: COLORS.primary, font: "Calibri" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Marmaraere\u011flisi Organize Sanayi B\u00f6lgesi", bold: true, size: 30, color: COLORS.secondary, font: "Calibri" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "Sanayi ve Teknoloji Bakanl\u0131\u011f\u0131 \u2014 Sanayi B\u00f6lgeleri Genel M\u00fcd\u00fcrl\u00fc\u011f\u00fc", size: 22, color: COLORS.grayText, font: "Calibri" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLORS.secondary } }, children: [new TextRun({ text: "" })] }),
        emptyLine(),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "PARSELASYON PLANI", bold: true, size: 48, color: COLORS.primary, font: "Calibri" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "A\u00c7IKLAMA RAPORU", bold: true, size: 44, color: COLORS.secondary, font: "Calibri" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "3194 Say\u0131l\u0131 \u0130mar Kanunu Md. 18 | Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelik", size: 20, color: COLORS.grayText, font: "Calibri", italics: true })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "TKGM 2025/4 Say\u0131l\u0131 Genelge Uyar\u0131nca Haz\u0131rlanm\u0131\u015ft\u0131r", size: 20, color: COLORS.citation, font: "Calibri", bold: true })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: COLORS.accent } }, children: [new TextRun({ text: "" })] }),
        emptyLine(),
        infoBox([
          { label: "\u0130li", value: "Tekirda\u011f" },
          { label: "\u0130l\u00e7esi", value: "Marmaraere\u011flisi" },
          { label: "B\u00f6lge", value: "Marmaraere\u011flisi Organize Sanayi B\u00f6lgesi" },
          { label: "Uygulama Alan\u0131", value: "Yakla\u015f\u0131k 73 Hektar (699.476,07 m\u00b2 kadastral)" },
          { label: "Toplam OSB Alan\u0131", value: "95,75 Hektar" },
          { label: "Kadastro Parsel Say\u0131s\u0131", value: "31 adet" },
        ]),
        emptyLine(), emptyLine(),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Haz\u0131rlayan: Ekin Harita M\u00fchendislik ve M\u00fc\u015favirlik B\u00fcrosu", size: 22, color: COLORS.darkText, font: "Calibri", bold: true })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Alipa\u015fa Mah. \u00c7etin Eme\u00e7 Bulvar\u0131 No:140/3 \u2013 \u00c7ORLU", size: 20, color: COLORS.grayText, font: "Calibri" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Tel: 0532 255 96 47 | ekin.harita@hotmail.com", size: 20, color: COLORS.grayText, font: "Calibri" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "Tarih: 2026", size: 22, color: COLORS.grayText, font: "Calibri" })] }),
        new Paragraph({ spacing: { before: 100 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: COLORS.primary } }, children: [new TextRun({ text: "" })] }),
      ],
    },
    // ===== TOC =====
    {
      properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      children: [
        heading1("\u0130\u00c7\u0130NDEK\u0130LER"),
        new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
        new Paragraph({ spacing: { before: 400 }, children: [new TextRun({ text: "Not: \u0130\u00e7indekiler tablosunu g\u00f6r\u00fcnt\u00fclemek i\u00e7in Word\u2019te sa\u011f t\u0131klay\u0131p \u201cAlan\u0131 G\u00fcncelle\u201d se\u00e7iniz.", size: 18, color: COLORS.grayText, font: "Calibri", italics: true })] }),
      ],
    },
    // ===== KISALTMALAR =====
    {
      properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      children: [
        heading1("KISALTMALAR VE TANIMLAR"),
        emptyLine(),
        styledTable(["K\u0131saltma", "A\u00e7\u0131klama"], [
          ["DOP", "D\u00fczenleme Ortakl\u0131k Pay\u0131 (3194 s.K. md. 18/2)"],
          ["KOP", "Kamu Ortakl\u0131k Pay\u0131 (3194 s.K. md. 18/3)"],
          ["OSB", "Organize Sanayi B\u00f6lgesi (4562 s.K.)"],
          ["TAKS", "Taban Alan\u0131 Kat Say\u0131s\u0131"],
          ["KAKS", "Kat Alan\u0131 Kat Say\u0131s\u0131 (Emsal)"],
          ["TKGM", "Tapu ve Kadastro Genel M\u00fcd\u00fcrl\u00fc\u011f\u00fc"],
          ["TE\u0130A\u015e", "T\u00fcrkiye Elektrik \u0130letim A.\u015e."],
          ["TEDA\u015e", "T\u00fcrkiye Elektrik Da\u011f\u0131t\u0131m A.\u015e."],
          ["ITRF96", "International Terrestrial Reference Frame 1996"],
          ["TUREF", "T\u00fcrkiye Ulusal Referans \u00c7er\u00e7evesi"],
          ["PARS\u0130D", "Parselasyon Plan\u0131 \u0130zleme ve Denetim Sistemi (Y\u00f6n. de\u011f. 11.09.2025)"],
        ]),
      ],
    },
    // ===== ANA İÇERİK =====
    {
      properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: COLORS.border } }, children: [new TextRun({ text: "Marmaraere\u011flisi OSB \u2014 Parselasyon Plan\u0131 A\u00e7\u0131klama Raporu", size: 16, color: COLORS.grayText, font: "Calibri", italics: true })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, border: { top: { style: BorderStyle.SINGLE, size: 2, color: COLORS.border } }, children: [new TextRun({ text: "Sayfa ", size: 16, color: COLORS.grayText, font: "Calibri" }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: COLORS.primary, font: "Calibri", bold: true })] })] }) },
      children: [
        // 1. GİRİŞ
        heading1("1. G\u0130R\u0130\u015e VE AMA\u00c7"),
        heading2("1.1. Raporun Amac\u0131"),
        richParagraph([t("\u0130\u015fbu rapor, 3194 say\u0131l\u0131 \u0130mar Kanunu\u2019nun 18. maddesi"), cite("3194 s.K. md. 18"), t(" ve Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelik h\u00fck\u00fcmleri"), cite("R.G. 22.02.2020/31047; de\u011f. R.G. 11.09.2025/33014"), t(" \u00e7er\u00e7evesinde haz\u0131rlanan, Tekirda\u011f ili Marmaraere\u011flisi il\u00e7esi Marmaraere\u011flisi Organize Sanayi B\u00f6lgesi parselasyon plan\u0131n\u0131n teknik, hukuki ve idari boyutlar\u0131n\u0131 a\u00e7\u0131klamak amac\u0131yla kaleme al\u0131nm\u0131\u015ft\u0131r.")], { indent: true }),
        richParagraph([t("Raporun haz\u0131rlanmas\u0131nda, TKGM 2025/4 say\u0131l\u0131 \u201cTescile Konu Harita ve Planlar\u0131n Yap\u0131m\u0131 ve Kontrol\u00fc Genelgesi\u201d"), cite("TKGM, 12.05.2025, E-32381507-170.03.02-16435018"), t(" esas al\u0131nm\u0131\u015ft\u0131r.")], { indent: true }),
        heading2("1.2. Raporun Kapsam\u0131"),
        richParagraph([t("Rapor, Marmaraere\u011flisi OSB s\u0131n\u0131rlar\u0131 dahilindeki yakla\u015f\u0131k 73 hektarl\u0131k alanda bulunan 31 adet kadastro parselinin imar parsellerine d\u00f6n\u00fc\u015ft\u00fcr\u00fclmesi s\u00fcrecini kapsamaktad\u0131r. Y\u00f6netmeli\u011fin 23. maddesinde belirtilen a\u00e7\u0131klama raporu esaslar\u0131na"), cite("Y\u00f6n. md. 23"), t(" uygun olarak m\u00fclkiyet durumu, imar plan\u0131 kararlar\u0131, DOP hesab\u0131, alan dengelemesi ve da\u011f\u0131t\u0131m-tahsis i\u015flemlerini i\u00e7ermektedir.")], { indent: true }),
        heading2("1.3. OSB T\u00fczel Ki\u015fili\u011fi ve Yer Se\u00e7imi"),
        richParagraph([t("Marmaraere\u011flisi OSB, T.C. Sanayi ve Teknoloji Bakanl\u0131\u011f\u0131\u2019n\u0131n 19.03.2020 tarih ve 37419779-453.01/59 say\u0131l\u0131 yaz\u0131s\u0131 ile \u0131slah \u015fartlar\u0131n\u0131 yerine getirerek OSB t\u00fczel ki\u015fili\u011fi kazanm\u0131\u015ft\u0131r."), cite("4562 s. OSB Kanunu"), t(" 15.09.2022 tarih ve 4024629 say\u0131l\u0131 Bakanl\u0131k karar\u0131 ile yer se\u00e7imi \u00e7al\u0131\u015fmas\u0131 tamamlanarak 95,75 hektarl\u0131k alan\u0131 kapsayan b\u00f6lge toplam 25 parselden olu\u015fmaktad\u0131r.")], { indent: true }),
        heading2("1.4. Uygulaman\u0131n Gerek\u00e7esi"),
        richParagraph([t("\u0130darenin bu alanda 3194 say\u0131l\u0131 \u0130mar Kanunu\u2019nun 18. madde uygulamas\u0131 yapmak istemesinin temel gerek\u00e7esi, onayl\u0131 1/1000 \u00f6l\u00e7ekli Uygulama \u0130mar Plan\u0131nda yap\u0131lan de\u011fi\u015fikli\u011fi uygulamak ve 1/1000 \u00f6l\u00e7ekli Uygulama \u0130mar Plan\u0131na uygun \u015fekilde imar parselleri \u00fcretmektir."), cite("3194 s.K. md. 18/1")], { indent: true }),

        // 2. YASAL DAYANAK
        heading1("2. YASAL DAYANAK VE MEVZUAT \u00c7ER\u00c7EVES\u0130"),
        heading2("2.1. Temel Mevzuat"),
        emptyLine(),
        styledTable(["S\u0131ra", "Mevzuat", "R.G. Tarih/Say\u0131", "Konu"], [
          ["1", "3194 Say\u0131l\u0131 \u0130mar Kanunu (Md. 18)", "09.05.1985 / 18749", "Arazi ve arsa d\u00fczenlemesi"],
          ["2", "Arazi ve Arsa D\u00fcz. Hk. Y\u00f6netmelik", "22.02.2020 / 31047", "Uygulama usul ve esaslar\u0131"],
          ["3", "Y\u00f6netmelik De\u011fi\u015fikli\u011fi", "11.09.2025 / 33014", "PARS\u0130D, imar hakk\u0131 aktar\u0131m\u0131"],
          ["4", "TKGM 2025/4 Say\u0131l\u0131 Genelge", "12.05.2025", "Tescile konu harita/plan kontrol\u00fc"],
          ["5", "4562 Say\u0131l\u0131 OSB Kanunu", "15.04.2000 / 24021", "OSB kurulu\u015f ve i\u015fleyimi"],
          ["6", "3194 s.K. Md. 15-16", "09.05.1985 / 18749", "\u0130fraz, tevhit, yola terk"],
          ["7", "2942 Say\u0131l\u0131 Kamula\u015ft\u0131rma Kanunu", "08.11.1983 / 18215", "Kamula\u015ft\u0131rma i\u015flemleri"],
          ["8", "4721 Say\u0131l\u0131 T\u00fcrk Medeni Kanunu", "08.12.2001 / 24607", "M\u00fclkiyet hakk\u0131"],
          ["9", "3402 Say\u0131l\u0131 Kadastro Kanunu", "09.07.1987 / 19512", "Kadastro tespiti"],
        ]),
        emptyLine(),
        heading2("2.2. TKGM 2025/4 Say\u0131l\u0131 Genelge"),
        richParagraph([t("TKGM\u2019nin 12.05.2025 tarihli ve 2025/4 say\u0131l\u0131 \u201cTescile Konu Harita ve Planlar\u0131n Yap\u0131m\u0131 ve Kontrol\u00fc Genelgesi\u201d"), cite("TKGM 2025/4"), t(", parselasyon planlar\u0131n\u0131n teknik yap\u0131m esaslar\u0131n\u0131, m\u00fclkiyet raporu gerekliliklerini ve tescil \u00f6ncesi kontrol kriterlerini d\u00fczenlemektedir.")], { indent: true }),
        emptyLine(),
        lawBox(["TKGM 2025/4, B\u00f6l\u00fcm 5 \u2014 Parselasyon Haritalar\u0131:", "\u00ab\u0130mar plan\u0131na ba\u011fl\u0131 ay\u0131rma ve/veya birle\u015ftirme, parselasyon, toplu-", "la\u015ft\u0131rma ile kamula\u015ft\u0131rma haritalar\u0131na ili\u015fkin i\u015fleme tabi parseller", "i\u00e7in m\u00fclkiyet raporu d\u00fczenlenir.\u00bb"]),
        emptyLine(),
        heading2("2.3. 11.09.2025 Tarihli Y\u00f6netmelik De\u011fi\u015fikli\u011fi"),
        richParagraph([t("Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelikte 11 Eyl\u00fcl 2025 tarihinde yap\u0131lan de\u011fi\u015fiklik"), cite("R.G. 11.09.2025/33014"), t(" ile getirilen yenilikler:")], { indent: true }),
        bodyText("\u2022  PARS\u0130D sistemi ile elektronik ask\u0131 ve itiraz s\u00fcreci"),
        bodyText("\u2022  Ask\u0131 Da\u011f\u0131t\u0131m Cetveli: ki\u015fisel verilerin korunmas\u0131na uygun yeni belge t\u00fcr\u00fc"),
        bodyText("\u2022  \u0130mar Hakk\u0131 Aktar\u0131m\u0131: verici ve al\u0131c\u0131 parsel mekanizmas\u0131"),
        bodyText("\u2022  DOP + \u00f6nceki kesinti toplam\u0131 %45 \u00fcst s\u0131n\u0131r\u0131n\u0131n peki\u015ftirilmesi"),

        // 3. DÜZENLEME SAHASI
        heading1("3. D\u00dcZENLEME SAHASI B\u0130LG\u0130LER\u0130"),
        heading2("3.1. Co\u011frafi Konum"),
        richParagraph([t("D\u00fczenleme sahas\u0131, Tekirda\u011f ili Marmaraere\u011flisi il\u00e7esi s\u0131n\u0131rlar\u0131 dahilinde, Marmaraere\u011flisi Organize Sanayi B\u00f6lgesi i\u00e7erisinde yer almaktad\u0131r. 15.09.2022 tarih ve 4024629 say\u0131l\u0131 Bakanl\u0131k karar\u0131 ile yer se\u00e7imi tamamlanan b\u00f6lge, toplam 95,75 hektarl\u0131k bir alan\u0131 kapsamakta olup, parselasyon uygulamas\u0131na konu alan yakla\u015f\u0131k 73 hektard\u0131r.")], { indent: true }),
        heading2("3.2. D\u00fczenleme S\u0131n\u0131r\u0131"),
        richParagraph([t("D\u00fczenleme s\u0131n\u0131r\u0131, Y\u00f6netmeli\u011fin 5. maddesi h\u00fck\u00fcmleri do\u011frultusunda"), cite("Y\u00f6n. md. 5"), t(" belirlenmi\u015ftir. Toplam 31 adet kadastro parselinden 30 tanesi d\u00fczenleme alan\u0131na tam olarak girmi\u015f, 1 adet parsel (112/16) i\u00e7in ay\u0131rma \u00e7ap\u0131 d\u00fczenlenmi\u015ftir."), cite("Y\u00f6n. md. 23/1-a")], { indent: true }),
        heading2("3.3. \u0130hdas Durumu"),
        richParagraph([t("Uygulama alan\u0131nda kapanan kadastral yollardan herhangi bir ihdas i\u015flemi yap\u0131lmam\u0131\u015ft\u0131r."), cite("3194 s.K. md. 17")], { indent: true }),

        // 4. MÜLKİYET
        heading1("4. M\u00dcLK\u0130YET DURUMU ANAL\u0130Z\u0130"),
        heading2("4.1. M\u00fclkiyet T\u00fcr\u00fc Da\u011f\u0131l\u0131m\u0131"),
        richParagraph([t("Y\u00f6netmeli\u011fin 23/1-a bendi ve TKGM 2025/4 Genelgesi uyar\u0131nca"), cite("Y\u00f6n. md. 23/1-a; TKGM 2025/4, B\u00f6l\u00fcm 5"), t(" haz\u0131rlanan m\u00fclkiyet raporu \u00e7er\u00e7evesinde, 31 adet kadastro parselinin m\u00fclkiyet da\u011f\u0131l\u0131m\u0131:")], { indent: true }),
        emptyLine(),
        styledTable(["M\u00fclkiyet T\u00fcr\u00fc", "Parsel Adedi", "A\u00e7\u0131klama"], [
          ["Marmaraere\u011flisi OSB", "13", "OSB t\u00fczel ki\u015fili\u011fi ad\u0131na kay\u0131tl\u0131"],
          ["TE\u0130A\u015e", "5", "T\u00fcrkiye Elektrik \u0130letim A.\u015e."],
          ["TEDA\u015e Genel M\u00fcd\u00fcrl\u00fc\u011f\u00fc", "3", "T\u00fcrkiye Elektrik Da\u011f\u0131t\u0131m A.\u015e."],
          ["\u00d6zel M\u00fclkiyet", "10", "\u015eah\u0131s parselleri"],
          ["TOPLAM", "31", ""],
        ]),
        emptyLine(),
        heading2("4.2. Kadastral Parsel Envanteri"),
        richParagraph([t("Uygulamaya giren 31 adet kadastro parselinin detayl\u0131 listesi:")], { indent: true }),
        emptyLine(),
        styledTable(["Ada/Parsel", "Alan (m\u00b2)", "Cinsi"], [
          ["112/16", "39.578,32", "Tarla"], ["112/20", "42.813,48", "Tarla"], ["112/166", "6.149,43", "Tarla"], ["112/167", "16.518,81", "Tarla"],
          ["288/1", "37.481,98", "Arsa"], ["288/2", "3.075,76", "Arsa"], ["288/3", "3.137,18", "Arsa"], ["288/4", "3.428,86", "Arsa"], ["288/5", "3.750,78", "Arsa"],
          ["288/6", "71,63", "Pilon Yeri"], ["288/7", "529,21", "Pilon Yeri"], ["288/8", "49,39", "Yol"], ["288/9", "526,36", "Pilon Yeri"],
          ["288/10", "156,98", "Yol"], ["288/11", "354,47", "Yol"], ["288/12", "18,55", "Pilon Yeri"], ["288/13", "37,85", "Pilon Yeri"],
          ["288/14", "515.969,13", "Sanayi Bina ve Arsas\u0131"], ["288/28", "12,01", "Pilon Yeri"], ["288/29", "30,04", "Pilon Yeri"], ["288/30", "5,52", "Pilon Yeri"],
          ["288/31", "1.958,73", "Spor Alan\u0131"], ["288/32", "1.958,06", "Spor Alan\u0131"], ["288/33", "0,66", "Yol"],
          ["288/36", "2.515,76", "Park"], ["288/37", "1.042,42", "Yol"], ["288/38", "19,98", "Yol"], ["288/39", "401,21", "Yol"], ["288/40", "164,91", "Yol"],
          ["289/11", "3.697,74", "Teknik Altyap\u0131"], ["289/12", "14.020,86", "At\u0131ksu Ar\u0131tma Tesis Alan\u0131"],
        ]),
        emptyLine(),
        heading2("4.3. \u00d6nceki \u0130mar Uygulamalar\u0131"),
        richParagraph([t("Uygulama sahas\u0131ndaki parsellerin b\u00fcy\u00fck \u00e7o\u011funlu\u011fu daha \u00f6nce yap\u0131lan 3194 say\u0131l\u0131 Kanun\u2019un 18. madde uygulamas\u0131ndan \u00fcretilmi\u015ftir."), cite("3194 s.K. md. 18; Y\u00f6n. md. 23/1-a"), t(" Yap\u0131lan imar plan\u0131 de\u011fi\u015fikli\u011finde herhangi bir emsal art\u0131\u015f\u0131 veya yap\u0131la\u015fma ko\u015fullar\u0131nda iyile\u015ftirme anlam\u0131na gelecek bir de\u011fi\u015fiklik bulunmamaktad\u0131r. Dolay\u0131s\u0131yla bu parsellerden DOP kesintisi yap\u0131lmam\u0131\u015ft\u0131r.")], { indent: true }),
        richParagraph([t("Uygulama sahas\u0131nda ilk kez uygulamaya giren 4 adet parsel bulunmaktad\u0131r: ", { bold: true }), t("112/16, 112/20, 112/166 ve 112/167 numaral\u0131 parseller. Bu parsellerden DOP kesintisi yap\u0131lm\u0131\u015ft\u0131r."), cite("Y\u00f6n. md. 14")], { indent: true }),

        // 5. İMAR PLANI
        heading1("5. \u0130MAR PLANI B\u0130LG\u0130LER\u0130 VE PLAN KARARLARI"),
        heading2("5.1. Uygulama \u0130mar Plan\u0131 Bilgileri"),
        richParagraph([t("Y\u00f6netmeli\u011fin 23/1-b bendi gere\u011fince"), cite("Y\u00f6n. md. 23/1-b"), t(" plan bilgileri:")], { indent: true }),
        emptyLine(),
        infoBox([
          { label: "Plan T\u00fcr\u00fc", value: "1/1000 \u00d6l\u00e7ekli Uygulama \u0130mar Plan\u0131 De\u011fi\u015fikli\u011fi" },
          { label: "Onaylayan Kurum", value: "Sanayi ve Teknoloji Bakanl\u0131\u011f\u0131, Sanayi B\u00f6lgeleri Genel M\u00fcd\u00fcrl\u00fc\u011f\u00fc" },
          { label: "Onay Tarihi", value: "29.09.2025" },
          { label: "Karar No", value: "2025/25-156" },
          { label: "Plan \u00d6l\u00e7e\u011fi", value: "1/1000" },
          { label: "Yasal Dayanak", value: "4562 s. OSB Kanunu; 3194 s.K. md. 8/b" },
        ]),
        emptyLine(),
        richParagraph([t("1/1000 \u00d6l\u00e7ekli Uygulama \u0130mar Plan\u0131 ve Plan Notlar\u0131n\u0131n onayl\u0131 \u00f6rne\u011fi raporun ekindedir.")], { indent: true }),
        heading2("5.2. Plan De\u011fi\u015fikli\u011fi De\u011ferlendirmesi"),
        richParagraph([t("Yap\u0131lan imar plan\u0131 de\u011fi\u015fikli\u011finde herhangi bir emsal art\u0131\u015f\u0131 veya yap\u0131la\u015fma ko\u015fullar\u0131nda iyile\u015ftirme anlam\u0131na gelecek bir de\u011fi\u015fiklik bulunmamaktad\u0131r. De\u011fi\u015fiklik yaln\u0131zca parsel geometrisi ve fonksiyon alan\u0131 d\u00fczenlemelerine y\u00f6neliktir.")], { indent: true }),

        // 6. DOP
        heading1("6. D\u00dcZENLEME ORTAKLIK PAYI (DOP) HESABI"),
        heading2("6.1. DOP Hesaplama Y\u00f6ntemi"),
        richParagraph([t("DOP, 3194 s.K. md. 18/2"), cite("3194 s.K. md. 18/2"), t(" ve Y\u00f6netmelik md. 11-16"), cite("Y\u00f6n. md. 11-16"), t(" \u00e7er\u00e7evesinde hesaplanm\u0131\u015ft\u0131r. Parsellerin b\u00fcy\u00fck \u00e7o\u011funlu\u011fu \u00f6nceki 18. madde uygulamas\u0131ndan \u00fcretildi\u011finden, yaln\u0131zca ilk kez uygulamaya giren 4 parselden (112/16, 112/20, 112/166, 112/167) DOP kesilmi\u015ftir."), cite("Y\u00f6n. md. 14")], { indent: true }),
        emptyLine(),
        lawBox(["3194 Say\u0131l\u0131 \u0130mar Kanunu, Madde 18/2:", "\u00abD\u00fczenleme ortakl\u0131k pay\u0131 oran\u0131, d\u00fczenlemeye tabi tutulan arazi ve", "arsalar\u0131n d\u00fczenleme \u00f6ncesi toplam alan\u0131n\u0131n y\u00fczde k\u0131rk be\u015fini ge\u00e7emez.\u00bb"]),
        emptyLine(),
        heading2("6.2. DOP Hesap Tablosu"),
        emptyLine(),
        styledTable(["Kalem", "De\u011fer", "Yasal Dayanak"], [
          ["DOP Oran\u0131", "%8,82 (0.0882434)", "3194 s.K. md. 18/2"],
          ["DOP Al\u0131nan Toplam Alan", "93.822,48 m\u00b2", "Y\u00f6n. md. 11-16"],
          ["DOP Al\u0131nmayan Toplam Alan", "594.416,03 m\u00b2", "\u00d6nceki md. 18 uyg. (m\u00fcktesep hak)"],
          ["Toplam Kadastro Parsel Alan\u0131", "699.476,07 m\u00b2", "Y\u00f6n. md. 23/1-a"],
          ["Toplam \u0130mar Ada Alan\u0131", "679.959,30 m\u00b2", "-"],
        ]),
        emptyLine(),
        heading2("6.3. DOP Oran\u0131 De\u011ferlendirmesi"),
        richParagraph([t("Hesaplanan DOP oran\u0131 %8,82 olup, 3194 s.K.\u2019n\u0131n \u00f6ng\u00f6rd\u00fc\u011f\u00fc %45 \u00fcst s\u0131n\u0131r\u0131n\u0131n \u00e7ok alt\u0131ndad\u0131r."), cite("3194 s.K. md. 18/2"), t(" 11.09.2025 tarihli Y\u00f6netmelik de\u011fi\u015fikli\u011fi uyar\u0131nca kesinti + \u00f6nceki terk toplam\u0131 da %45\u2019i ge\u00e7memektedir."), cite("Y\u00f6n. de\u011f. R.G. 11.09.2025/33014")], { indent: true }),

        // 7. ALAN HESAPLARI
        heading1("7. ALAN HESAPLARI VE DENGELEME"),
        heading2("7.1. Dengeleme Tablosu"),
        richParagraph([t("Y\u00f6netmelik md. 22/1-b uyar\u0131nca"), cite("Y\u00f6n. md. 22/1-b"), t(" alan hesaplar\u0131:")], { indent: true }),
        emptyLine(),
        styledTable(["A\u00e7\u0131klama", "Alan (m\u00b2)"], [
          ["Toplam Kadastro Parsel Alan\u0131", "699.476,07"],
          ["DOP Al\u0131nan Parseller", "93.822,48"],
          ["DOP Al\u0131nmayan Parseller", "594.416,03"],
          ["Toplam \u0130mar Ada Alan\u0131", "679.959,30"],
        ]),

        // 8. DAĞITIM
        emptyLine(),
        heading1("8. DA\u011eITIM VE TAHS\u0130S \u0130\u015eLEMLER\u0130"),
        heading2("8.1. Da\u011f\u0131t\u0131m \u0130lkeleri"),
        richParagraph([t("Da\u011f\u0131t\u0131m, Y\u00f6netmelik md. 17 esas al\u0131narak yap\u0131lm\u0131\u015ft\u0131r."), cite("Y\u00f6n. md. 17"), t(" Temel kriterler:")], { indent: true }),
        richParagraph([t("\u2022  Kadastro parselinin bast\u0131\u011f\u0131 yerden tahsis (yerindelik)"), cite("Y\u00f6n. md. 19")]),
        richParagraph([t("\u2022  Mevcut yap\u0131la\u015fmalar\u0131n korunmas\u0131"), cite("Y\u00f6n. md. 20")]),
        bodyText("\u2022  M\u00fcstakil parsel olu\u015fturulabilecek hisselerin de\u011ferlendirilmesi"),
        bodyText("\u2022  Plan geometrisinin izin verdi\u011fi \u00f6l\u00e7\u00fcde kendi yerinden tahsis"),
        heading2("8.2. Uygulama Sonras\u0131 Olu\u015fan Durum"),
        richParagraph([t("Tescile konu 30 adet imar parseli olu\u015fmu\u015ftur:")], { indent: true }),
        emptyLine(),
        styledTable(["Fonksiyon Alan\u0131", "Parsel Adedi", "A\u00e7\u0131klama"], [
          ["Sanayi Alan\u0131", "4", "OSB sanayi parselleri"],
          ["Pilon Yeri", "5", "\u0130rtifak hakk\u0131 korunarak tescil"],
          ["Hizmet Destek Alan\u0131", "6", "OSB hizmet alanlar\u0131"],
          ["A\u00e7\u0131k Spor Tesis Alan\u0131", "1", "Spor tesisi"],
          ["Teknik Altyap\u0131 Alan\u0131", "1", "DOP\u2019tan kar\u015f\u0131lanarak tescil"],
          ["\u0130rtifak Hakl\u0131 Parseller", "10", "\u0130rtifak hakk\u0131 korunarak tescil"],
          ["Di\u011fer", "3", "Kalan parseller"],
          ["TOPLAM", "30", ""],
        ]),
        emptyLine(),
        heading2("8.3. Tescil Durumu"),
        richParagraph([t("30 imar parselinden 1 tanesi (Teknik Altyap\u0131) DOP\u2019tan kar\u015f\u0131lanarak tescil edilecektir."), cite("3194 s.K. md. 18/2"), t(" 10 parsel irtifak hakk\u0131n\u0131 korumak amac\u0131yla (pilon yerleri, enerji hatlar\u0131) tescil edilecektir."), cite("4721 s.K. md. 779 vd.")], { indent: true }),

        // 9. TEKNİK
        heading1("9. TEKN\u0130K A\u00c7IKLAMALAR"),
        heading2("9.1. Koordinat Sistemi"),
        richParagraph([t("T\u00fcm \u00f6l\u00e7\u00fcm ve hesaplamalar TKGM 2025/4 md. 6-7 uyar\u0131nca"), cite("TKGM 2025/4, md. 6-7"), t(" ITRF96/TUREF koordinat sisteminde ger\u00e7ekle\u015ftirilmi\u015ftir.")], { indent: true }),
        heading2("9.2. Ay\u0131rma \u00c7ap\u0131"),
        richParagraph([t("31 kadastro parselinden 30\u2019u tam olarak girmi\u015f, 112/16 (39.578,32 m\u00b2) i\u00e7in ay\u0131rma \u00e7ap\u0131 d\u00fczenlenmi\u015ftir."), cite("3194 s.K. md. 15-16; TKGM 2025/4")], { indent: true }),
        heading2("9.3. Pafta Bilgileri"),
        richParagraph([t("Parselasyon paftalar\u0131 TKGM 2025/4 B\u00f6l\u00fcm 4 standartlar\u0131na"), cite("TKGM 2025/4, B\u00f6l\u00fcm 4"), t(" ve Y\u00f6n. md. 22/1-c,d\u2019ye"), cite("Y\u00f6n. md. 22/1-c,d"), t(" uygun d\u00fczenlenmi\u015ftir.")], { indent: true }),

        // 10. KURUM GÖRÜŞLERİ
        heading1("10. \u0130LG\u0130L\u0130 KURUM G\u00d6R\u00dc\u015eLER\u0130"),
        richParagraph([t("Y\u00f6n. md. 22/1-e gere\u011fince"), cite("Y\u00f6n. md. 22/1-e"), t(" al\u0131nan g\u00f6r\u00fc\u015fler:")], { indent: true }),
        emptyLine(),
        styledTable(["Kurum", "Dayanak", "Konu"], [
          ["Tapu M\u00fcd\u00fcrl\u00fc\u011f\u00fc", "TKGM 2025/4", "Tescil kontrol\u00fc"],
          ["Kadastro M\u00fcd\u00fcrl\u00fc\u011f\u00fc", "3402 s.K.", "Kadastral veri do\u011frulama"],
          ["TE\u0130A\u015e", "6446 s.K.", "Enerji iletim hatlar\u0131 ve pilon yerleri"],
          ["TEDA\u015e", "6446 s.K.", "Enerji da\u011f\u0131t\u0131m hatlar\u0131"],
          ["Sanayi ve Tek. Bakanl\u0131\u011f\u0131", "4562 s. OSB K.", "\u0130mar plan\u0131 onay\u0131"],
        ]),

        // 11. SONUÇ
        emptyLine(),
        heading1("11. SONU\u00c7 VE DE\u011eERLEND\u0130RME"),
        richParagraph([t("\u0130\u015fbu parselasyon plan\u0131, 3194 s.K. md. 18"), cite("3194 s.K. md. 18"), t(", Arazi ve Arsa D\u00fcz. Hk. Y\u00f6netmelik"), cite("R.G. 22.02.2020/31047; de\u011f. R.G. 11.09.2025/33014"), t(" ve TKGM 2025/4 Genelgesi"), cite("TKGM, 12.05.2025"), t(" h\u00fck\u00fcmleri \u00e7er\u00e7evesinde, Sanayi B\u00f6lgeleri Genel M\u00fcd\u00fcrl\u00fc\u011f\u00fc\u2019n\u00fcn 29.09.2025 tarih ve 2025/25-156 say\u0131l\u0131 karar\u0131 ile onaylanan imar plan\u0131 de\u011fi\u015fikli\u011fine uygun olarak haz\u0131rlanm\u0131\u015ft\u0131r.")], { indent: true }),
        richParagraph([t("31 adet kadastro parselinden 30 adet imar parseli olu\u015fturulmu\u015f, DOP oran\u0131 %8,82 olarak hesaplanm\u0131\u015ft\u0131r (%45 \u00fcst s\u0131n\u0131r\u0131n \u00e7ok alt\u0131nda). DOP yaln\u0131zca ilk kez uygulamaya giren 4 parselden (112/16, 112/20, 112/166, 112/167) kesilmi\u015f olup, \u00f6nceki uygulama g\u00f6rm\u00fc\u015f parsellerden m\u00fcktesep hak ilkesi gere\u011fi kesinti yap\u0131lmam\u0131\u015ft\u0131r.")], { indent: true }),
        richParagraph([t("Parselasyon plan\u0131n\u0131n onaylanarak tescil edilmek \u00fczere Tapu M\u00fcd\u00fcrl\u00fc\u011f\u00fc\u2019ne g\u00f6nderilmesi hususunda gere\u011fini arz ederiz.")], { indent: true }),
        emptyLine(), emptyLine(), separatorLine(),
        new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 60 }, children: [t("Haz\u0131rlayan", { bold: true, color: COLORS.primary })] }),
        new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 60 }, children: [t("Ekin Harita M\u00fchendislik ve M\u00fc\u015favirlik B\u00fcrosu", { bold: true })] }),
        new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 60 }, children: [t("Alipa\u015fa Mah. \u00c7etin Eme\u00e7 Bulvar\u0131 No:140/3 \u2013 \u00c7ORLU", { color: COLORS.grayText })] }),
        new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 200 }, children: [t("\u0130mza / Ka\u015fe", { color: COLORS.grayText })] }),

        // 12. EKLER
        separatorLine(),
        heading1("12. EKLER"),
        bodyText("Ek-1: D\u00fczenleme S\u0131n\u0131r\u0131 Koordinat Listesi"),
        bodyText("Ek-2: Kadastro Parselleri Tapu Kay\u0131t \u00d6zetleri"),
        bodyText("Ek-3: M\u00fclkiyet Raporu (TKGM 2025/4 format\u0131)"),
        bodyText("Ek-4: 1/1000 \u00d6l\u00e7ekli Uygulama \u0130mar Plan\u0131 ve Plan Notlar\u0131 (Onayl\u0131)"),
        bodyText("Ek-5: Parselasyon Plan\u0131 Paftas\u0131"),
        bodyText("Ek-6: Ada R\u00f6leve Krokileri"),
        bodyText("Ek-7: Alan Hesap \u00c7izelgeleri ve Dengeleme Tablosu"),
        bodyText("Ek-8: DOP Hesap Detaylar\u0131"),
        bodyText("Ek-9: Da\u011f\u0131t\u0131m ve Tahsis Cetvelleri"),
        bodyText("Ek-10: Ask\u0131 Da\u011f\u0131t\u0131m Cetveli (PARS\u0130D format\u0131)"),
        bodyText("Ek-11: 112/16 Parseli Ay\u0131rma \u00c7ap\u0131"),
        bodyText("Ek-12: Uygulama Alan\u0131 Uydu G\u00f6r\u00fcnt\u00fcs\u00fc"),
        bodyText("Ek-13: Uygulama Alan\u0131 Plan G\u00f6r\u00fcnt\u00fcs\u00fc"),
        bodyText("Ek-14: Uygulama Sonras\u0131 G\u00f6r\u00fcn\u00fcm"),
        bodyText("Ek-15: \u0130lgili Kurum G\u00f6r\u00fc\u015f Yaz\u0131lar\u0131"),
        bodyText("Ek-16: Say\u0131sal Veriler (CD/USB)"),

        // 13. KAYNAKÇA
        emptyLine(),
        heading1("13. KAYNAK\u00c7A"),
        heading2("13.1. Kanunlar"),
        bodyText("\u2022  3194 Say\u0131l\u0131 \u0130mar Kanunu, R.G. 09.05.1985/18749"),
        bodyText("\u2022  4562 Say\u0131l\u0131 Organize Sanayi B\u00f6lgeleri Kanunu, R.G. 15.04.2000/24021"),
        bodyText("\u2022  2942 Say\u0131l\u0131 Kamula\u015ft\u0131rma Kanunu, R.G. 08.11.1983/18215"),
        bodyText("\u2022  4721 Say\u0131l\u0131 T\u00fcrk Medeni Kanunu, R.G. 08.12.2001/24607"),
        bodyText("\u2022  3402 Say\u0131l\u0131 Kadastro Kanunu, R.G. 09.07.1987/19512"),
        bodyText("\u2022  6446 Say\u0131l\u0131 Elektrik Piyasas\u0131 Kanunu, R.G. 30.03.2013/28603"),
        heading2("13.2. Y\u00f6netmelikler"),
        bodyText("\u2022  Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelik, R.G. 22.02.2020/31047"),
        bodyText("\u2022  Arazi ve Arsa D\u00fcz. Hk. Y\u00f6n. De\u011fi\u015fikli\u011fi, R.G. 11.09.2025/33014"),
        bodyText("\u2022  Mek\u00e2nsal Planlar Yap\u0131m Y\u00f6netmeli\u011fi, R.G. 14.06.2014/29030"),
        bodyText("\u2022  B\u00fcy\u00fck \u00d6l\u00e7ekli Harita ve Harita Bilgileri \u00dcretim Y\u00f6netmeli\u011fi, R.G. 15.07.2005/25876"),
        heading2("13.3. Genelgeler"),
        bodyText("\u2022  TKGM 2025/4 Tescile Konu Harita ve Planlar\u0131n Yap\u0131m\u0131 ve Kontrol\u00fc Genelgesi, 12.05.2025"),
        heading2("13.4. \u0130dari Kararlar"),
        bodyText("\u2022  Sanayi ve Teknoloji Bakanl\u0131\u011f\u0131, 19.03.2020, 37419779-453.01/59 \u2014 OSB t\u00fczel ki\u015fili\u011fi"),
        bodyText("\u2022  Sanayi ve Teknoloji Bakanl\u0131\u011f\u0131, 15.09.2022, 4024629 \u2014 Yer se\u00e7imi"),
        bodyText("\u2022  Sanayi B\u00f6lgeleri Genel M\u00fcd\u00fcrl\u00fc\u011f\u00fc, 29.09.2025, 2025/25-156 \u2014 \u0130mar plan\u0131 onay\u0131"),
      ],
    },
  ],
});

const outputPath = "C:/Users/STUUX/Desktop/Cloude_Deneme/Parselasyon_Plani_Aciklama_Raporu.docx";
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log("Marsab verili rapor olusturuldu: " + outputPath);
});
