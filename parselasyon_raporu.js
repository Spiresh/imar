const docx = require("docx");
const fs = require("fs");

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, HeadingLevel, BorderStyle, ShadingType,
  TableOfContents, PageBreak, ImageRun, Header, Footer, PageNumber,
  NumberFormat, Tab, TabStopPosition, TabStopType, convertInchesToTwip
} = docx;

const PAGE_WIDTH = 12240;
const MARGIN = 1440;
const CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN;

const COLORS = {
  primary: "1B3A5C",
  secondary: "2E75B6",
  accent: "4472C4",
  lightBg: "E8EFF7",
  mediumBg: "D6E4F0",
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
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: COLORS.secondary },
    },
    children: [
      new TextRun({
        text: text.toUpperCase(),
        bold: true,
        size: 28,
        color: COLORS.primary,
        font: "Calibri",
      }),
    ],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [
      new TextRun({
        text: "  ",
        size: 24,
      }),
      new TextRun({
        text,
        bold: true,
        size: 24,
        color: COLORS.secondary,
        font: "Calibri",
      }),
    ],
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 120 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 22,
        color: COLORS.accent,
        font: "Calibri",
        italics: true,
      }),
    ],
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
    children: [
      new TextRun({
        text,
        size: 22,
        color: COLORS.darkText,
        font: "Calibri",
        italics: opts.italic || false,
        bold: opts.bold || false,
      }),
    ],
  });
}

function cite(tag) {
  return new TextRun({
    text: " (" + tag + ")",
    size: 20,
    color: COLORS.citation,
    font: "Calibri",
    italics: true,
    bold: true,
  });
}

function t(text, opts = {}) {
  return new TextRun({
    text,
    size: opts.size || 22,
    color: opts.color || COLORS.darkText,
    font: "Calibri",
    italics: opts.italic || false,
    bold: opts.bold || false,
  });
}

function placeholder(text) {
  return new Paragraph({
    spacing: { after: 120, line: 276 },
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 720 },
    children: [
      new TextRun({
        text: "[" + text + "]",
        size: 22,
        color: COLORS.grayText,
        font: "Calibri",
        italics: true,
      }),
    ],
  });
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
    children: lines.map(
      (l) =>
        new Paragraph({
          spacing: { after: 60 },
          children: [
            new TextRun({
              text: l.label + ": ",
              bold: true,
              size: 20,
              color: COLORS.primary,
              font: "Calibri",
            }),
            new TextRun({
              text: l.value,
              size: 20,
              color: COLORS.darkText,
              font: "Calibri",
            }),
          ],
        })
    ),
  });

  return new Table({
    rows: [new TableRow({ children: [cell] })],
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  });
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
    children: lines.map(
      (l) =>
        new Paragraph({
          spacing: { after: 60 },
          children: [
            new TextRun({
              text: l,
              size: 20,
              color: "5C4A00",
              font: "Calibri",
              italics: true,
            }),
          ],
        })
    ),
  });

  return new Table({
    rows: [new TableRow({ children: [cell] })],
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  });
}

function styledTable(headers, rows) {
  const colCount = headers.length;
  const colWidth = Math.floor(CONTENT_WIDTH / colCount);

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map(
      (h) =>
        new TableCell({
          width: { size: colWidth, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: COLORS.headerBg },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
            left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
            right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
          },
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: h,
                  bold: true,
                  size: 20,
                  color: COLORS.white,
                  font: "Calibri",
                }),
              ],
            }),
          ],
        })
    ),
  });

  const dataRows = rows.map(
    (row, idx) =>
      new TableRow({
        children: row.map(
          (cell) =>
            new TableCell({
              width: { size: colWidth, type: WidthType.DXA },
              shading: {
                type: ShadingType.CLEAR,
                fill: idx % 2 === 0 ? COLORS.white : COLORS.rowAlt,
              },
              borders: {
                top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
                bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
                left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
                right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.border },
              },
              margins: { top: 40, bottom: 40, left: 100, right: 100 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cell,
                      size: 20,
                      color: COLORS.darkText,
                      font: "Calibri",
                    }),
                  ],
                }),
              ],
            })
        ),
      })
  );

  return new Table({
    rows: [headerRow, ...dataRows],
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: Array(colCount).fill(colWidth),
  });
}

function emptyLine() {
  return new Paragraph({ spacing: { after: 80 } });
}

function separatorLine() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 2, color: COLORS.border },
    },
    children: [new TextRun({ text: "" })],
  });
}

// ============= DOCUMENT =============
const doc = new Document({
  creator: "Parselasyon Plan\u0131 Otomasyonu",
  title: "Parselasyon Plan\u0131 A\u00e7\u0131klama Raporu",
  description: "3194 Say\u0131l\u0131 \u0130mar Kanunu Madde 18 Kapsam\u0131nda Haz\u0131rlanan Parselasyon Plan\u0131 A\u00e7\u0131klama Raporu",
  styles: {
    default: {
      document: {
        run: { font: "Calibri", size: 22, color: COLORS.darkText },
        paragraph: { spacing: { line: 276 } },
      },
      heading1: {
        run: { font: "Calibri", size: 28, bold: true, color: COLORS.primary },
        paragraph: { spacing: { before: 360, after: 200 } },
      },
      heading2: {
        run: { font: "Calibri", size: 24, bold: true, color: COLORS.secondary },
        paragraph: { spacing: { before: 280, after: 160 } },
      },
      heading3: {
        run: { font: "Calibri", size: 22, bold: true, color: COLORS.accent, italics: true },
        paragraph: { spacing: { before: 200, after: 120 } },
      },
    },
  },
  features: { updateFields: true },
  sections: [
    // ===== KAPAK SAYFASI =====
    {
      properties: {
        page: {
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
          pageNumbers: { start: 1 },
        },
      },
      children: [
        emptyLine(),
        emptyLine(),
        emptyLine(),
        emptyLine(),
        new Paragraph({
          spacing: { after: 400 },
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 12, color: COLORS.primary },
          },
          children: [new TextRun({ text: "" })],
        }),
        emptyLine(),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [
            new TextRun({
              text: "T.C.",
              bold: true,
              size: 32,
              color: COLORS.primary,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [
            new TextRun({
              text: "[Belediye / \u0130l \u00d6zel \u0130daresi Ad\u0131]",
              bold: true,
              size: 30,
              color: COLORS.secondary,
              font: "Calibri",
              italics: true,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [
            new TextRun({
              text: "\u0130mar ve \u015eehircilik M\u00fcd\u00fcrl\u00fc\u011f\u00fc",
              size: 26,
              color: COLORS.grayText,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 },
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 4, color: COLORS.secondary },
          },
          children: [new TextRun({ text: "" })],
        }),
        emptyLine(),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [
            new TextRun({
              text: "PARSELASYON PLANI",
              bold: true,
              size: 48,
              color: COLORS.primary,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [
            new TextRun({
              text: "A\u00c7IKLAMA RAPORU",
              bold: true,
              size: 44,
              color: COLORS.secondary,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [
            new TextRun({
              text: "3194 Say\u0131l\u0131 \u0130mar Kanunu Md. 18 | Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelik",
              size: 20,
              color: COLORS.grayText,
              font: "Calibri",
              italics: true,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 },
          children: [
            new TextRun({
              text: "TKGM 2025/4 Say\u0131l\u0131 Genelge Uyar\u0131nca Haz\u0131rlanm\u0131\u015ft\u0131r",
              size: 20,
              color: COLORS.citation,
              font: "Calibri",
              bold: true,
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 },
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 4, color: COLORS.accent },
          },
          children: [new TextRun({ text: "" })],
        }),
        emptyLine(),
        infoBox([
          { label: "\u0130li", value: "[.........]" },
          { label: "\u0130l\u00e7esi", value: "[.........]" },
          { label: "Mahalle/K\u00f6y", value: "[.........]" },
          { label: "Ada/Parsel", value: "[.........]" },
          { label: "D\u00fczenleme Alan\u0131", value: "[.........] m\u00b2" },
          { label: "Pafta No", value: "[.........]" },
        ]),
        emptyLine(),
        emptyLine(),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Haz\u0131rlayan: [Harita M\u00fchendisi / \u015eehir Planc\u0131s\u0131 Ad\u0131 Soyad\u0131]", size: 22, color: COLORS.grayText, font: "Calibri" }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [
            new TextRun({ text: "Oda Sicil No: [.........]", size: 22, color: COLORS.grayText, font: "Calibri" }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [
            new TextRun({ text: "Tarih: [GG.AA.YYYY]", size: 22, color: COLORS.grayText, font: "Calibri" }),
          ],
        }),
        new Paragraph({
          spacing: { before: 100 },
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 12, color: COLORS.primary },
          },
          children: [new TextRun({ text: "" })],
        }),
      ],
    },

    // ===== TABLE OF CONTENTS =====
    {
      properties: {
        page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
      },
      children: [
        heading1("\u0130\u00c7\u0130NDEK\u0130LER"),
        new TableOfContents("Table of Contents", {
          hyperlink: true,
          headingStyleRange: "1-3",
        }),
        new Paragraph({
          spacing: { before: 400 },
          children: [
            new TextRun({
              text: "Not: \u0130\u00e7indekiler tablosunu g\u00f6r\u00fcnt\u00fclemek i\u00e7in Word\u2019te sa\u011f t\u0131klay\u0131p \u201cAlan\u0131 G\u00fcncelle\u201d se\u00e7iniz.",
              size: 18, color: COLORS.grayText, font: "Calibri", italics: true,
            }),
          ],
        }),
      ],
    },

    // ===== KISALTMALAR =====
    {
      properties: {
        page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
      },
      children: [
        heading1("KISALTMALAR VE TANIMLAR"),
        emptyLine(),
        styledTable(
          ["K\u0131saltma", "A\u00e7\u0131klama"],
          [
            ["DOP", "D\u00fczenleme Ortakl\u0131k Pay\u0131 (3194 s.K. md. 18/2)"],
            ["KOP", "Kamu Ortakl\u0131k Pay\u0131 (3194 s.K. md. 18/3)"],
            ["TAKS", "Taban Alan\u0131 Kat Say\u0131s\u0131"],
            ["KAKS", "Kat Alan\u0131 Kat Say\u0131s\u0131 (Emsal)"],
            ["Hmaks", "Maksimum Bina Y\u00fcksekli\u011fi"],
            ["TKGM", "Tapu ve Kadastro Genel M\u00fcd\u00fcrl\u00fc\u011f\u00fc"],
            ["ITRF96", "International Terrestrial Reference Frame 1996"],
            ["TUREF", "T\u00fcrkiye Ulusal Referans \u00c7er\u00e7evesi"],
            ["UTM", "Universal Transverse Mercator Projeksiyonu"],
            ["PARS\u0130D", "Parselasyon Plan\u0131 \u0130zleme ve Denetim Sistemi (Y\u00f6n. de\u011f. 11.09.2025)"],
            ["e-Devlet", "T\u00fcrkiye.gov.tr Elektronik Devlet Kap\u0131s\u0131"],
          ]
        ),
      ],
    },

    // ===== ANA BÖLÜMLER =====
    {
      properties: {
        page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              border: {
                bottom: { style: BorderStyle.SINGLE, size: 2, color: COLORS.border },
              },
              children: [
                new TextRun({ text: "Parselasyon Plan\u0131 A\u00e7\u0131klama Raporu", size: 16, color: COLORS.grayText, font: "Calibri", italics: true }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              border: {
                top: { style: BorderStyle.SINGLE, size: 2, color: COLORS.border },
              },
              children: [
                new TextRun({ text: "Sayfa ", size: 16, color: COLORS.grayText, font: "Calibri" }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: COLORS.primary, font: "Calibri", bold: true }),
              ],
            }),
          ],
        }),
      },
      children: [
        // ===== 1. GİRİŞ VE AMAÇ =====
        heading1("1. G\u0130R\u0130\u015e VE AMA\u00c7"),

        heading2("1.1. Raporun Amac\u0131"),
        richParagraph([
          t("\u0130\u015fbu rapor, 3194 say\u0131l\u0131 \u0130mar Kanunu\u2019nun 18. maddesi"),
          cite("3194 s.K. md. 18"),
          t(" ve Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelik h\u00fck\u00fcmleri"),
          cite("R.G. 22.02.2020/31047; de\u011f. R.G. 11.09.2025/33014"),
          t(" \u00e7er\u00e7evesinde haz\u0131rlanan parselasyon plan\u0131n\u0131n teknik, hukuki ve idari boyutlar\u0131n\u0131 a\u00e7\u0131klamak amac\u0131yla kaleme al\u0131nm\u0131\u015ft\u0131r. Rapor; d\u00fczenleme sahas\u0131n\u0131n mevcut durumunu, uygulanan imar plan\u0131 kararlar\u0131n\u0131, d\u00fczenleme ortakl\u0131k pay\u0131 (DOP) hesaplamalar\u0131n\u0131 ve da\u011f\u0131t\u0131m-tahsis i\u015flemlerini sistematik bi\u00e7imde ortaya koymaktad\u0131r."),
        ], { indent: true }),

        richParagraph([
          t("Raporun haz\u0131rlanmas\u0131nda, tescile konu harita ve planlar\u0131n yap\u0131m\u0131 ile kontrol\u00fcne ili\u015fkin g\u00fcncel d\u00fczenleme olan TKGM 2025/4 say\u0131l\u0131 Genelge"),
          cite("TKGM, 12.05.2025, E-32381507-170.03.02-16435018"),
          t(" esas al\u0131nm\u0131\u015ft\u0131r. S\u00f6z konusu genelgenin be\u015finci b\u00f6l\u00fcm\u00fcnde d\u00fczenlenen parselasyon haritalar\u0131n\u0131n yap\u0131m esaslar\u0131 ve m\u00fclkiyet raporu gereklilikleri titizlikle uygulanm\u0131\u015ft\u0131r."),
        ], { indent: true }),

        heading2("1.2. Raporun Kapsam\u0131"),
        richParagraph([
          t("Rapor, Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmeli\u011fin 23. maddesinde belirtilen a\u00e7\u0131klama raporu esaslar\u0131na"),
          cite("Y\u00f6n. md. 23"),
          t(" uygun olarak; m\u00fclkiyet durumu, kadastral yap\u0131, imar plan\u0131 kararlar\u0131, DOP hesab\u0131, alan dengelemesi, da\u011f\u0131t\u0131m-tahsis i\u015flemleri ve teknik a\u00e7\u0131klamalar\u0131 kapsamaktad\u0131r. Ayr\u0131ca Y\u00f6netmeli\u011fin 22. maddesinde say\u0131lan hesap ve belgeler de bu raporla bir b\u00fct\u00fcn olarak sunulmaktad\u0131r."),
          cite("Y\u00f6n. md. 22"),
        ], { indent: true }),

        heading2("1.3. \u00c7al\u0131\u015fma Y\u00f6ntemi"),
        richParagraph([
          t("Parselasyon plan\u0131n\u0131n haz\u0131rlanmas\u0131nda; arazi \u00f6l\u00e7\u00fcm verileri, kadastro kay\u0131tlar\u0131, tapu sicil bilgileri, hali haz\u0131r harita verileri, uygulama imar plan\u0131 paftalar\u0131 ile ilgili kurum ve kurulu\u015f g\u00f6r\u00fc\u015fleri esas al\u0131nm\u0131\u015ft\u0131r. T\u00fcm \u00f6l\u00e7\u00fcm ve hesaplamalar B\u00fcy\u00fck \u00d6l\u00e7ekli Harita ve Harita Bilgileri \u00dcretim Y\u00f6netmeli\u011fi"),
          cite("R.G. 15.07.2005/25876"),
          t(" ile TKGM 2025/4 say\u0131l\u0131 Genelge\u2019nin teknik standartlar\u0131na"),
          cite("TKGM 2025/4, B\u00f6l\u00fcm 2-3"),
          t(" uygun olarak ITRF96/TUREF koordinat sisteminde ger\u00e7ekle\u015ftirilmi\u015ftir."),
        ], { indent: true }),
        placeholder("Kullan\u0131lan yaz\u0131l\u0131m, koordinat sistemi ve veri kaynaklar\u0131 detayland\u0131r\u0131lacakt\u0131r."),

        // ===== 2. YASAL DAYANAK =====
        heading1("2. YASAL DAYANAK VE MEVZUAT \u00c7ER\u00c7EVES\u0130"),

        heading2("2.1. Temel Mevzuat"),
        richParagraph([
          t("Parselasyon plan\u0131n\u0131n haz\u0131rlanmas\u0131nda a\u015fa\u011f\u0131da belirtilen yasal d\u00fczenlemeler esas al\u0131nm\u0131\u015ft\u0131r:"),
        ], { indent: true }),
        emptyLine(),

        styledTable(
          ["S\u0131ra", "Mevzuat", "R.G. Tarih/Say\u0131", "Konu"],
          [
            ["1", "3194 Say\u0131l\u0131 \u0130mar Kanunu (Md. 18)", "09.05.1985 / 18749", "Arazi ve arsa d\u00fczenlemesi"],
            ["2", "Arazi ve Arsa D\u00fcz. Hk. Y\u00f6netmelik", "22.02.2020 / 31047", "Uygulama usul ve esaslar\u0131"],
            ["3", "Y\u00f6netmelik De\u011fi\u015fikli\u011fi", "11.09.2025 / 33014", "PARS\u0130D, imar hakk\u0131 aktar\u0131m\u0131, ask\u0131 cetveli"],
            ["4", "TKGM 2025/4 Say\u0131l\u0131 Genelge", "12.05.2025", "Tescile konu harita/plan yap\u0131m ve kontrol\u00fc"],
            ["5", "3194 s.K. Md. 15-16", "09.05.1985 / 18749", "\u0130fraz, tevhit, yola terk"],
            ["6", "2942 Say\u0131l\u0131 Kamula\u015ft\u0131rma Kanunu", "08.11.1983 / 18215", "Kamula\u015ft\u0131rma i\u015flemleri"],
            ["7", "5393 Say\u0131l\u0131 Belediye Kanunu", "13.07.2005 / 25874", "\u0130dari yetki ve onay"],
            ["8", "4721 Say\u0131l\u0131 T\u00fcrk Medeni Kanunu", "08.12.2001 / 24607", "M\u00fclkiyet hakk\u0131"],
            ["9", "3402 Say\u0131l\u0131 Kadastro Kanunu", "09.07.1987 / 19512", "Kadastro tespiti"],
            ["10", "2863 Say\u0131l\u0131 K\u00fclt\u00fcr/Tabiat Varl. K.", "23.07.1983 / 18113", "Koruma alanlar\u0131"],
            ["11", "6831 Say\u0131l\u0131 Orman Kanunu", "08.09.1956 / 9402", "Orman alanlar\u0131"],
            ["12", "3621 Say\u0131l\u0131 K\u0131y\u0131 Kanunu", "17.04.1990 / 20495", "K\u0131y\u0131 alanlar\u0131"],
          ]
        ),

        emptyLine(),
        heading2("2.2. TKGM 2025/4 Say\u0131l\u0131 Genelge Kapsam\u0131"),
        richParagraph([
          t("Tapu ve Kadastro Genel M\u00fcd\u00fcrl\u00fc\u011f\u00fc\u2019n\u00fcn 12.05.2025 tarihli ve 2025/4 say\u0131l\u0131 \u201cTescile Konu Harita ve Planlar\u0131n Yap\u0131m\u0131 ve Kontrol\u00fc Genelgesi\u201d"),
          cite("TKGM 2025/4"),
          t(", parselasyon planlar\u0131n\u0131n teknik yap\u0131m esaslar\u0131n\u0131 ve kontrol s\u00fcre\u00e7lerini d\u00fczenlemektedir. Genelgenin be\u015finci b\u00f6l\u00fcm\u00fcnde parselasyon haritalar\u0131n\u0131n yap\u0131m\u0131na ili\u015fkin \u00f6zel h\u00fck\u00fcmler yer almakta olup, m\u00fclkiyet raporunun i\u00e7eri\u011fi ve tescil \u00f6ncesi kontrol kriterleri ayr\u0131nt\u0131l\u0131 bi\u00e7imde belirlenmektedir."),
        ], { indent: true }),

        emptyLine(),
        lawBox([
          "TKGM 2025/4, B\u00f6l\u00fcm 5 \u2014 Parselasyon Haritalar\u0131:",
          "\u00ab\u0130mar plan\u0131na ba\u011fl\u0131 ay\u0131rma ve/veya birle\u015ftirme, parselasyon, toplu-",
          "la\u015ft\u0131rma ile kamula\u015ft\u0131rma haritalar\u0131na ili\u015fkin i\u015fleme tabi parseller",
          "i\u00e7in m\u00fclkiyet raporu d\u00fczenlenir. Bu raporda uygulamaya al\u0131nacak",
          "parsellerin listesi, koordinatlar\u0131 ve y\u00fcz\u00f6l\u00e7\u00fcmleri ile d\u00fczenleme",
          "ortakl\u0131k pay\u0131 oran\u0131 yer al\u0131r.\u00bb",
        ]),

        emptyLine(),
        heading2("2.3. 11.09.2025 Tarihli Y\u00f6netmelik De\u011fi\u015fikli\u011fi"),
        richParagraph([
          t("Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelikte 11 Eyl\u00fcl 2025 tarihinde yap\u0131lan de\u011fi\u015fiklik"),
          cite("R.G. 11.09.2025/33014"),
          t(" ile parselasyon s\u00fcrecine \u00f6nemli yenilikler getirilmi\u015ftir:"),
        ], { indent: true }),

        bodyText("\u2022  PARS\u0130D (Parselasyon Plan\u0131 \u0130zleme ve Denetim Sistemi): Parselasyon planlar\u0131n\u0131n duyurusu, ask\u0131ya \u00e7\u0131kar\u0131lmas\u0131 ve itiraz s\u00fcre\u00e7leri e-Devlet entegrasyonu ile elektronik ortamda y\u00fcr\u00fct\u00fclecektir."),
        bodyText("\u2022  Ask\u0131 Da\u011f\u0131t\u0131m Cetveli: Ki\u015fisel verilerin korunmas\u0131 mevzuat\u0131na uygun olarak haz\u0131rlanan ve internet \u00fczerinden yay\u0131mlanan yeni belge t\u00fcr\u00fc tan\u0131mlanm\u0131\u015ft\u0131r."),
        bodyText("\u2022  \u0130mar Hakk\u0131 Aktar\u0131m\u0131: Kamu hizmet alan\u0131na isabet eden ve yap\u0131la\u015fma hakk\u0131n\u0131 kaybeden parsellerin (verici parsel) imar haklar\u0131n\u0131n ba\u015fka parsellere (al\u0131c\u0131 parsel) aktar\u0131labilmesine olanak sa\u011flayan d\u00fczenleme getirilmi\u015ftir."),
        bodyText("\u2022  DOP + \u00d6nceki Kesinti Toplam\u0131: Parselasyon plan\u0131nda yap\u0131lan kesinti miktar\u0131 ile \u00f6nceki uygulamalardan kaynaklanan kesinti ve terklerin toplam\u0131n\u0131n %45\u2019i ge\u00e7emeyece\u011fi h\u00fckm\u00fc peki\u015ftirilmi\u015ftir."),

        emptyLine(),
        heading2("2.4. \u0130lgili Yarg\u0131 Kararlar\u0131"),
        richParagraph([
          t("Parselasyon uygulamalar\u0131na ili\u015fkin Dan\u0131\u015ftay 6. Dairesi, \u0130dari Dava Daireleri Kurulu ve Anayasa Mahkemesi kararlar\u0131 da dikkate al\u0131nm\u0131\u015ft\u0131r. \u00d6zellikle DOP oran\u0131n\u0131n \u00fcst s\u0131n\u0131r\u0131na"),
          cite("AYM, 2018/53 E., 2021/31 K."),
          t(", tahsis adaletine"),
          cite("Dan. 6. D., 2019/12856 E."),
          t(" ve m\u00fclkiyet hakk\u0131n\u0131n korunmas\u0131na"),
          cite("AY md. 35; A\u0130HS Ek Protokol 1, md. 1"),
          t(" y\u00f6nelik i\u00e7tihatlar g\u00f6zetilmi\u015ftir."),
        ], { indent: true }),
        placeholder("G\u00fcncel emsal yarg\u0131 kararlar\u0131 ve i\u00e7tihatlar bu b\u00f6l\u00fcmde belirtilecektir."),

        // ===== 3. DÜZENLEME SAHASI =====
        heading1("3. D\u00dcZENLEME SAHASI B\u0130LG\u0130LER\u0130"),

        heading2("3.1. Co\u011frafi Konum ve S\u0131n\u0131rlar"),
        richParagraph([
          t("D\u00fczenleme sahas\u0131, [il/il\u00e7e/mahalle] s\u0131n\u0131rlar\u0131 dahilinde yer almakta olup, do\u011fusunda [...], bat\u0131s\u0131nda [...], kuzeyinde [...] ve g\u00fcneyinde [...] ile \u00e7evrilidir. Sahan\u0131n toplam alan\u0131 [....] m\u00b2 olup, TKGM 2025/4 say\u0131l\u0131 Genelge\u2019nin teknik standartlar\u0131na uygun olarak ITRF96 koordinat sisteminde tan\u0131mlanm\u0131\u015ft\u0131r."),
          cite("TKGM 2025/4, md. 6-7"),
        ], { indent: true }),
        placeholder("D\u00fczenleme s\u0131n\u0131r\u0131 koordinatlar\u0131, vaziyet plan\u0131 ve uydu g\u00f6r\u00fcnt\u00fcs\u00fc eklenecektir."),

        heading2("3.2. D\u00fczenleme S\u0131n\u0131r\u0131n\u0131n Belirlenmesi"),
        richParagraph([
          t("D\u00fczenleme s\u0131n\u0131r\u0131, Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmeli\u011fin 5. maddesi h\u00fck\u00fcmleri do\u011frultusunda"),
          cite("Y\u00f6n. md. 5"),
          t(", uygulama imar plan\u0131 s\u0131n\u0131rlar\u0131, m\u00fclkiyet yap\u0131s\u0131 ve arazi kullan\u0131m durumu g\u00f6zetilerek belirlenmi\u015ftir. S\u0131n\u0131r\u0131n belirlenmesinde b\u00fct\u00fcnl\u00fck ilkesi ve uygulanabilirlik kriterleri esas al\u0131nm\u0131\u015ft\u0131r. 3194 say\u0131l\u0131 Kanun\u2019un 18. maddesinin d\u00f6rd\u00fcnc\u00fc f\u0131kras\u0131 gere\u011fince d\u00fczenleme sahas\u0131n\u0131n bir m\u00fcstakil imar adas\u0131ndan k\u00fc\u00e7\u00fck olamayaca\u011f\u0131 hususu g\u00f6zetilmi\u015ftir."),
          cite("3194 s.K. md. 18/4"),
        ], { indent: true }),

        heading2("3.3. Topografya ve Arazi Yap\u0131s\u0131"),
        richParagraph([
          t("D\u00fczenleme sahas\u0131n\u0131n topografik yap\u0131s\u0131, e\u011fim analizi sonu\u00e7lar\u0131 ve jeolojik-jeoteknik verileri bu b\u00f6l\u00fcmde de\u011ferlendirilmektedir. Sahan\u0131n ortalama kotu [....] m, en d\u00fc\u015f\u00fck kotu [....] m ve en y\u00fcksek kotu [....] m olarak tespit edilmi\u015ftir. Zemin et\u00fcd raporu, Mekansal Planlar Yap\u0131m Y\u00f6netmeli\u011fi\u2019nin ilgili h\u00fck\u00fcmleri do\u011frultusunda de\u011ferlendirilmi\u015ftir."),
          cite("R.G. 14.06.2014/29030"),
        ], { indent: true }),
        placeholder("Topografik kesit, e\u011fim haritas\u0131 ve jeolojik rapor \u00f6zeti eklenecektir."),

        // ===== 4. MÜLK\u0130YET DURUMU =====
        heading1("4. M\u00dcLK\u0130YET DURUMU ANAL\u0130Z\u0130"),

        heading2("4.1. Kadastral Yap\u0131 ve Parsel Envanteri"),
        richParagraph([
          t("D\u00fczenleme sahas\u0131 i\u00e7erisinde toplam [....] adet kadastro parseli bulunmakta olup, bu parsellerin toplam alan\u0131 [....] m\u00b2\u2019dir. Y\u00f6netmeli\u011fin 23. maddesinin birinci f\u0131kras\u0131n\u0131n (a) bendi gere\u011fince d\u00fczenlemeye giren parsellerin say\u0131s\u0131, toplam alan\u0131 ve m\u00fclkiyet durumu a\u015fa\u011f\u0131daki tabloda sunulmaktad\u0131r."),
          cite("Y\u00f6n. md. 23/1-a"),
        ], { indent: true }),
        emptyLine(),

        styledTable(
          ["S\u0131ra No", "Ada No", "Parsel No", "Alan (m\u00b2)", "M\u00fclkiyet Durumu", "Tapu Cinsi"],
          [
            ["1", "[...]", "[...]", "[...]", "[Ger\u00e7ek/T\u00fczel Ki\u015fi]", "[Arsa/Tarla]"],
            ["2", "[...]", "[...]", "[...]", "[Ger\u00e7ek/T\u00fczel Ki\u015fi]", "[Arsa/Tarla]"],
            ["3", "[...]", "[...]", "[...]", "[Hazine/Belediye]", "[Arsa/Tarla]"],
            ["...", "...", "...", "...", "...", "..."],
          ]
        ),

        emptyLine(),
        heading2("4.2. M\u00fclkiyet T\u00fcr\u00fc Da\u011f\u0131l\u0131m\u0131"),
        richParagraph([
          t("TKGM 2025/4 say\u0131l\u0131 Genelge uyar\u0131nca haz\u0131rlanan m\u00fclkiyet raporunda, d\u00fczenleme sahas\u0131ndaki parsellerin m\u00fclkiyet t\u00fcr\u00fc itibar\u0131yla da\u011f\u0131l\u0131m\u0131 a\u015fa\u011f\u0131da sunulmaktad\u0131r:"),
          cite("TKGM 2025/4, B\u00f6l\u00fcm 5"),
        ], { indent: true }),
        emptyLine(),

        styledTable(
          ["M\u00fclkiyet T\u00fcr\u00fc", "Parsel Adedi", "Toplam Alan (m\u00b2)", "Oran (%)"],
          [
            ["\u00d6zel M\u00fclkiyet (\u015eah\u0131s)", "[...]", "[...]", "[...]"],
            ["Hazine", "[...]", "[...]", "[...]"],
            ["Belediye", "[...]", "[...]", "[...]"],
            ["Vak\u0131f", "[...]", "[...]", "[...]"],
            ["Di\u011fer Kamu", "[...]", "[...]", "[...]"],
            ["TOPLAM", "[...]", "[...]", "100"],
          ]
        ),

        emptyLine(),
        heading2("4.3. \u00d6zel Kanunlarla Korunan Alanlar"),
        richParagraph([
          t("D\u00fczenleme sahas\u0131 i\u00e7erisinde a\u015fa\u011f\u0131daki \u00f6zel kanunlar kapsam\u0131nda koruma alt\u0131nda bulunan alan ve parsellerin varl\u0131\u011f\u0131 incelenmi\u015ftir:"),
        ], { indent: true }),
        richParagraph([
          t("\u2022  2863 say\u0131l\u0131 K\u00fclt\u00fcr ve Tabiat Varl\u0131klar\u0131n\u0131 Koruma Kanunu"),
          cite("R.G. 23.07.1983/18113"),
        ]),
        richParagraph([
          t("\u2022  6831 say\u0131l\u0131 Orman Kanunu"),
          cite("R.G. 08.09.1956/9402"),
        ]),
        richParagraph([
          t("\u2022  3621 say\u0131l\u0131 K\u0131y\u0131 Kanunu"),
          cite("R.G. 17.04.1990/20495"),
        ]),
        richParagraph([
          t("\u2022  7269 say\u0131l\u0131 Umumi Hayata M\u00fcessir Afetler Kanunu"),
          cite("R.G. 25.05.1959/10213"),
        ]),
        placeholder("\u00d6zel kanunla korunan alan listesi, sit alan\u0131 kararlar\u0131 ve ilgili kurul g\u00f6r\u00fc\u015fleri eklenecektir."),

        heading2("4.4. \u00d6nceki \u0130mar Uygulamalar\u0131"),
        richParagraph([
          t("Y\u00f6netmeli\u011fin 23. maddesinin birinci f\u0131kras\u0131n\u0131n (a) bendi gere\u011fince, daha \u00f6nce 3194 say\u0131l\u0131 Kanun\u2019un 15., 16. veya 18. maddelerine g\u00f6re i\u015flem g\u00f6rm\u00fc\u015f parsellerin tespiti yap\u0131lm\u0131\u015ft\u0131r."),
          cite("Y\u00f6n. md. 23/1-a; 3194 s.K. md. 15-16-18"),
          t(" Bu parsellere ait terk oranlar\u0131 ve m\u00fcktesep haklar, TKGM 2025/4 say\u0131l\u0131 Genelge\u2019nin ilgili h\u00fck\u00fcmleri do\u011frultusunda de\u011ferlendirilmi\u015ftir. Genelge uyar\u0131nca, daha \u00f6nce imar uygulamas\u0131 g\u00f6rerek olu\u015fmu\u015f imar parsellerinin yeniden ifraz, tevhit, yola terk ve yoldan ihdas gibi i\u015flemlere tabi tutulmas\u0131 halinde ayr\u0131ca teknik rapor d\u00fczenlenmesi gerekmemektedir."),
          cite("TKGM 2025/4, md. 29"),
        ], { indent: true }),
        emptyLine(),

        styledTable(
          ["Ada/Parsel", "Uygulama T\u00fcr\u00fc", "Yasal Dayanak", "Tarih", "Terk Oran\u0131 (%)", "A\u00e7\u0131klama"],
          [
            ["[...]", "Md. 18", "3194 s.K.", "[...]", "[...]", "[...]"],
            ["[...]", "Md. 15-16", "3194 s.K.", "[...]", "[...]", "[...]"],
            ["[...]", "Ay\u0131rma-Birle\u015ftirme", "3194 s.K.", "[...]", "[...]", "[...]"],
          ]
        ),

        emptyLine(),
        heading2("4.5. Fiili Durum Tespiti"),
        richParagraph([
          t("D\u00fczenleme sahas\u0131ndaki yap\u0131la\u015fma durumu, mevcut kullan\u0131m bi\u00e7imleri ve fiili i\u015fgaller, Y\u00f6netmeli\u011fin 23. maddesi kapsam\u0131nda yerinde yap\u0131lan inceleme ile tespit edilmi\u015ftir."),
          cite("Y\u00f6n. md. 23/1-a"),
          t(" Muhdesat tespitleri 3402 say\u0131l\u0131 Kadastro Kanunu\u2019nun 19. maddesi h\u00fck\u00fcmleri \u00e7er\u00e7evesinde ger\u00e7ekle\u015ftirilmi\u015ftir."),
          cite("3402 s.K. md. 19"),
        ], { indent: true }),
        placeholder("Fiili durum foto\u011fraflar\u0131, muhdesat listesi ve encr\u00f6\u015fman analizi eklenecektir."),

        // ===== 5. İMAR PLANI BİLGİLERİ =====
        heading1("5. \u0130MAR PLANI B\u0130LG\u0130LER\u0130 VE PLAN KARARLARI"),

        heading2("5.1. Uygulama \u0130mar Plan\u0131 Bilgileri"),
        richParagraph([
          t("Parselasyon plan\u0131n\u0131n dayana\u011f\u0131 olan uygulama imar plan\u0131na ili\u015fkin bilgiler, Y\u00f6netmeli\u011fin 23. maddesinin birinci f\u0131kras\u0131n\u0131n (b) bendi gere\u011fince a\u015fa\u011f\u0131da sunulmaktad\u0131r:"),
          cite("Y\u00f6n. md. 23/1-b"),
        ], { indent: true }),
        emptyLine(),
        infoBox([
          { label: "Plan\u0131n Onay Tarihi", value: "[GG.AA.YYYY]" },
          { label: "Plan\u0131n Kesinle\u015fme Tarihi", value: "[GG.AA.YYYY]" },
          { label: "Plan \u00d6l\u00e7e\u011fi", value: "1/1000" },
          { label: "Onaylayan Kurum", value: "[.........]" },
          { label: "Meclis Karar No", value: "[.........]" },
          { label: "Plan Revizyonu", value: "[Var/Yok \u2013 Tarih]" },
          { label: "Yasal Dayanak", value: "3194 s.K. md. 8/b; Mek\u00e2nsal Planlar Yap\u0131m Y\u00f6n." },
        ]),

        emptyLine(),
        heading2("5.2. Arazi Kullan\u0131m Kararlar\u0131"),
        richParagraph([
          t("Uygulama imar plan\u0131nda d\u00fczenleme sahas\u0131 i\u00e7in \u00f6ng\u00f6r\u00fclen arazi kullan\u0131m kararlar\u0131, Mek\u00e2nsal Planlar Yap\u0131m Y\u00f6netmeli\u011fi\u2019nin plan g\u00f6sterim tekniklerine"),
          cite("R.G. 14.06.2014/29030, md. 7-8"),
          t(" uygun olarak a\u015fa\u011f\u0131da \u00f6zetlenmektedir:"),
        ], { indent: true }),
        emptyLine(),

        styledTable(
          ["Kullan\u0131m T\u00fcr\u00fc", "Plan Sembol\u00fc", "Alan (m\u00b2)", "Oran (%)", "Yasal Dayanak"],
          [
            ["Konut Alan\u0131", "K", "[...]", "[...]", "3194 s.K."],
            ["Ticaret Alan\u0131", "T", "[...]", "[...]", "3194 s.K."],
            ["Ye\u015fil Alan (Park)", "YA", "[...]", "[...]", "3194 s.K. md. 18/2"],
            ["E\u011fitim Tesisi", "E", "[...]", "[...]", "3194 s.K. md. 18/2"],
            ["Sa\u011fl\u0131k Tesisi", "S", "[...]", "[...]", "3194 s.K. md. 18/3 (KOP)"],
            ["Dini Tesis Alan\u0131", "DT", "[...]", "[...]", "3194 s.K. md. 18/2"],
            ["Yol", "Y", "[...]", "[...]", "3194 s.K. md. 18/2"],
            ["Otopark", "OP", "[...]", "[...]", "3194 s.K. md. 18/2"],
            ["Teknik Altyap\u0131", "TA", "[...]", "[...]", "3194 s.K. md. 18/3 (KOP)"],
            ["TOPLAM", "-", "[...]", "100", "-"],
          ]
        ),

        emptyLine(),
        heading2("5.3. Yap\u0131la\u015fma Ko\u015fullar\u0131"),
        richParagraph([
          t("\u0130mar plan\u0131nda \u00f6ng\u00f6r\u00fclen yap\u0131la\u015fma ko\u015fullar\u0131, Planl\u0131 Alanlar \u0130mar Y\u00f6netmeli\u011fi h\u00fck\u00fcmleri \u00e7er\u00e7evesinde"),
          cite("R.G. 03.07.2017/30113"),
          t(" a\u015fa\u011f\u0131daki tabloda sunulmaktad\u0131r:"),
        ], { indent: true }),
        emptyLine(),

        styledTable(
          ["Ada No", "Fonksiyon", "TAKS", "KAKS", "Hmaks (m)", "Kat", "\u00c7ekme (m)"],
          [
            ["[...]", "Konut", "[...]", "[...]", "[...]", "[...]", "[...]"],
            ["[...]", "Ticaret", "[...]", "[...]", "[...]", "[...]", "[...]"],
            ["[...]", "Karma", "[...]", "[...]", "[...]", "[...]", "[...]"],
          ]
        ),

        emptyLine(),
        heading2("5.4. Plan Notlar\u0131 ve \u00d6zel H\u00fck\u00fcmler"),
        richParagraph([
          t("Uygulama imar plan\u0131na ait genel ve \u00f6zel plan notlar\u0131, 3194 say\u0131l\u0131 Kanun\u2019un 8. maddesi ve Mek\u00e2nsal Planlar Yap\u0131m Y\u00f6netmeli\u011fi\u2019nin 26. maddesi uyar\u0131nca"),
          cite("3194 s.K. md. 8; R.G. 14.06.2014/29030, md. 26"),
          t(" bu b\u00f6l\u00fcmde detayl\u0131 olarak aktar\u0131lmaktad\u0131r. Parselasyon uygulamas\u0131n\u0131 do\u011frudan etkileyen h\u00fck\u00fcmler ayr\u0131ca de\u011ferlendirilmi\u015ftir."),
        ], { indent: true }),
        placeholder("Plan notlar\u0131n\u0131n tamam\u0131 ve \u00f6zel h\u00fck\u00fcmler bu b\u00f6l\u00fcme eklenecektir."),

        // ===== 6. DOP HESABI =====
        heading1("6. D\u00dcZENLEME ORTAKLIK PAYI (DOP) HESABI"),

        heading2("6.1. DOP Hesaplama Y\u00f6ntemi ve Esaslar\u0131"),
        richParagraph([
          t("D\u00fczenleme ortakl\u0131k pay\u0131, 3194 say\u0131l\u0131 Kanun\u2019un 18. maddesinin ikinci f\u0131kras\u0131"),
          cite("3194 s.K. md. 18/2"),
          t(" ve Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmeli\u011fin 11\u201316. maddeleri"),
          cite("Y\u00f6n. md. 11-16"),
          t(" \u00e7er\u00e7evesinde hesaplanm\u0131\u015ft\u0131r. DOP; yol, meydan, park, ye\u015fil alan, otopark, \u00e7ocuk bah\u00e7esi, cami, karakol ve Milli E\u011fitim Bakanl\u0131\u011f\u0131na ba\u011fl\u0131 ilk ve orta\u00f6\u011fretim kurumlar\u0131 i\u00e7in ayr\u0131lan alanlar\u0131 kar\u015f\u0131lamak \u00fczere hesaplanm\u0131\u015ft\u0131r."),
        ], { indent: true }),

        emptyLine(),
        lawBox([
          "3194 Say\u0131l\u0131 \u0130mar Kanunu, Madde 18/2:",
          "\u00abD\u00fczenleme ortakl\u0131k pay\u0131 oran\u0131, d\u00fczenlemeye tabi tutulan arazi ve",
          "arsalar\u0131n d\u00fczenleme \u00f6ncesi toplam alan\u0131n\u0131n y\u00fczde k\u0131rk be\u015fini ge\u00e7emez.\u00bb",
        ]),

        emptyLine(),
        heading2("6.2. DOP Hesap Tablosu"),
        emptyLine(),

        styledTable(
          ["Kalem", "De\u011fer", "Yasal Dayanak"],
          [
            ["D\u00fczenlemeye Giren Toplam Alan", "[.........] m\u00b2", "Y\u00f6n. md. 23/1-a"],
            ["Kamu Alanlar\u0131 \u0130htiyac\u0131 (Yol)", "[.........] m\u00b2", "3194 s.K. md. 18/2"],
            ["Kamu Alanlar\u0131 \u0130htiyac\u0131 (Park/Ye\u015fil)", "[.........] m\u00b2", "3194 s.K. md. 18/2"],
            ["Kamu Alanlar\u0131 \u0130htiyac\u0131 (E\u011fitim)", "[.........] m\u00b2", "3194 s.K. md. 18/2"],
            ["Kamu Alanlar\u0131 \u0130htiyac\u0131 (Dini Tesis)", "[.........] m\u00b2", "3194 s.K. md. 18/2"],
            ["Kamu Alanlar\u0131 \u0130htiyac\u0131 (Di\u011fer)", "[.........] m\u00b2", "3194 s.K. md. 18/2"],
            ["KOP Alanlar\u0131 (Sa\u011fl\u0131k, Teknik Altyap\u0131 vb.)", "[.........] m\u00b2", "3194 s.K. md. 18/3"],
            ["Toplam DOP Alan\u0131", "[.........] m\u00b2", "Y\u00f6n. md. 11-16"],
            ["DOP Oran\u0131", "% [...]", "3194 s.K. md. 18/2 (\u00fcst s\u0131n\u0131r %45)"],
          ]
        ),

        emptyLine(),
        heading2("6.3. DOP Oran\u0131 De\u011ferlendirmesi"),
        richParagraph([
          t("Hesaplanan DOP oran\u0131 %[....] olup, 3194 say\u0131l\u0131 Kanun\u2019un \u00f6ng\u00f6rd\u00fc\u011f\u00fc %45 \u00fcst s\u0131n\u0131r\u0131n\u0131n alt\u0131nda kalmaktad\u0131r."),
          cite("3194 s.K. md. 18/2"),
          t(" 11.09.2025 tarihli Y\u00f6netmelik de\u011fi\u015fikli\u011fi ile peki\u015ftirilen h\u00fck\u00fcm uyar\u0131nca, yap\u0131lan kesinti miktar\u0131 ile \u00f6nceki uygulamalardan kaynaklanan kesinti ve terklerin toplam\u0131 da %45\u2019i ge\u00e7memektedir."),
          cite("Y\u00f6n. de\u011f. R.G. 11.09.2025/33014"),
          t(" DOP ile kar\u015f\u0131lanamayan kamu hizmeti alanlar\u0131 i\u00e7in 2942 say\u0131l\u0131 Kamula\u015ft\u0131rma Kanunu h\u00fck\u00fcmlerine g\u00f6re kamula\u015ft\u0131rma yoluna gidilmesi gerekmektedir."),
          cite("2942 s.K."),
        ], { indent: true }),

        heading2("6.4. \u00d6nceki Uygulamalardan D\u00fc\u015f\u00fclecek Terkler"),
        richParagraph([
          t("Daha \u00f6nce 3194 say\u0131l\u0131 Kanun\u2019un 15/16 veya 18. maddeleri uyar\u0131nca uygulama g\u00f6rm\u00fc\u015f parsellerin mevcut terk oranlar\u0131, Y\u00f6netmeli\u011fin 14. maddesi gere\u011fince m\u00fcktesep hak olarak dikkate al\u0131nm\u0131\u015f ve DOP hesab\u0131nda ilgili oranlar d\u00fc\u015f\u00fclm\u00fc\u015ft\u00fcr."),
          cite("Y\u00f6n. md. 14; 3194 s.K. md. 18/2"),
        ], { indent: true }),

        // ===== 7. ALAN HESAPLARI VE DENGELEME =====
        heading1("7. ALAN HESAPLARI VE DENGELEME"),

        heading2("7.1. Kadastro-\u0130mar Alan Kar\u015f\u0131la\u015ft\u0131rmas\u0131"),
        richParagraph([
          t("Y\u00f6netmeli\u011fin 22. maddesinde belirtilen alan hesaplar\u0131 ve dengeleme tablosu"),
          cite("Y\u00f6n. md. 22/1-b"),
          t(" a\u015fa\u011f\u0131da sunulmaktad\u0131r:"),
        ], { indent: true }),
        emptyLine(),

        styledTable(
          ["A\u00e7\u0131klama", "Alan (m\u00b2)", "Yasal Dayanak"],
          [
            ["D\u00fczenlemeye Giren Toplam Kadastro Alan\u0131", "[...]", "Y\u00f6n. md. 23/1-a"],
            ["(-) DOP Olarak Ayr\u0131lan Alan", "[...]", "3194 s.K. md. 18/2"],
            ["(-) KOP Olarak Ayr\u0131lan Alan", "[...]", "3194 s.K. md. 18/3"],
            ["(-) Kamula\u015ft\u0131r\u0131lacak Alan", "[...]", "2942 s.K."],
            ["(+) Ba\u011f\u0131\u015f / Terk Alanlar\u0131", "[...]", "Y\u00f6n. md. 10"],
            ["(=) Da\u011f\u0131t\u0131lacak Toplam Alan", "[...]", "Y\u00f6n. md. 17"],
            ["Olu\u015fan \u0130mar Parselleri Toplam Alan\u0131", "[...]", "-"],
            ["Dengeleme Fark\u0131 (m\u00b2)", "[...]", "-"],
          ]
        ),

        emptyLine(),
        heading2("7.2. Alan Hesaplama Y\u00f6ntemi"),
        richParagraph([
          t("Alan hesaplamalar\u0131 koordinat y\u00f6ntemi ile yap\u0131lm\u0131\u015f olup, t\u00fcm hesaplar TKGM 2025/4 say\u0131l\u0131 Genelge\u2019nin teknik standartlar\u0131na"),
          cite("TKGM 2025/4, md. 6-8"),
          t(" ve B\u00fcy\u00fck \u00d6l\u00e7ekli Harita ve Harita Bilgileri \u00dcretim Y\u00f6netmeli\u011fi\u2019ne"),
          cite("R.G. 15.07.2005/25876"),
          t(" uygun olarak ITRF96 koordinat sisteminde ger\u00e7ekle\u015ftirilmi\u015ftir. Say\u0131sal verilerin do\u011fruluk d\u00fczeyi ve tolerans de\u011ferleri genelgede belirtilen s\u0131n\u0131rlar dahilindedir."),
        ], { indent: true }),

        // ===== 8. DAĞITIM VE TAHSİS =====
        heading1("8. DA\u011eITIM VE TAHS\u0130S \u0130\u015eLEMLER\u0130"),

        heading2("8.1. Da\u011f\u0131t\u0131m \u0130lkeleri"),
        richParagraph([
          t("Parselasyon plan\u0131nda imar parsellerinin da\u011f\u0131t\u0131m\u0131, Y\u00f6netmeli\u011fin 17\u201321. maddeleri"),
          cite("Y\u00f6n. md. 17-21"),
          t(" uyar\u0131nca a\u015fa\u011f\u0131daki ilkeler g\u00f6zetilerek yap\u0131lm\u0131\u015ft\u0131r:"),
        ], { indent: true }),
        richParagraph([
          t("\u2022  Parsel sahiplerinin m\u00fcmk\u00fcn mertebe eski yerlerinden tahsisi (yerindelik ilkesi)"),
          cite("Y\u00f6n. md. 19"),
        ]),
        richParagraph([
          t("\u2022  M\u00fclkiyet haklar\u0131n\u0131n korunmas\u0131 ve e\u015fit muamele"),
          cite("AY md. 35; 4721 s.K. md. 683"),
        ]),
        richParagraph([
          t("\u2022  Parsellerin imar plan\u0131na uygun, yap\u0131la\u015fmaya elveri\u015fli bi\u00e7imde olu\u015fturulmas\u0131"),
          cite("3194 s.K. md. 18/1"),
        ]),
        richParagraph([
          t("\u2022  Asgari parsel b\u00fcy\u00fckl\u00fcklerinin sa\u011flanmas\u0131"),
          cite("Planl\u0131 Alanlar \u0130mar Y\u00f6n., R.G. 03.07.2017/30113"),
        ]),
        richParagraph([
          t("\u2022  Hisseli tahsisin asgari d\u00fczeyde tutulmas\u0131"),
          cite("Y\u00f6n. md. 20"),
        ]),

        heading2("8.2. Da\u011f\u0131t\u0131m Cetveli"),
        richParagraph([
          t("A\u015fa\u011f\u0131daki tabloda, kadastro parsellerinden imar parsellerine yap\u0131lan tahsis i\u015flemleri \u00f6zetlenmektedir. 11.09.2025 tarihli Y\u00f6netmelik de\u011fi\u015fikli\u011fi ile getirilen ask\u0131 da\u011f\u0131t\u0131m cetveli format\u0131, PARS\u0130D \u00fczerinden e-Devlet entegrasyonuyla yay\u0131mlanacakt\u0131r."),
          cite("Y\u00f6n. de\u011f. R.G. 11.09.2025/33014"),
        ], { indent: true }),
        emptyLine(),

        styledTable(
          ["Eski Ada/Parsel", "Eski Alan (m\u00b2)", "DOP (m\u00b2)", "Yeni Ada/Parsel", "Yeni Alan (m\u00b2)"],
          [
            ["[...]", "[...]", "[...]", "[...]", "[...]"],
            ["[...]", "[...]", "[...]", "[...]", "[...]"],
            ["[...]", "[...]", "[...]", "[...]", "[...]"],
            ["[...]", "[...]", "[...]", "[...]", "[...]"],
          ]
        ),

        emptyLine(),
        heading2("8.3. \u0130mar Hakk\u0131 Aktar\u0131m\u0131 De\u011ferlendirmesi"),
        richParagraph([
          t("11.09.2025 tarihli Y\u00f6netmelik de\u011fi\u015fikli\u011fi ile getirilen imar hakk\u0131 aktar\u0131m\u0131 m\u00fcessesesi"),
          cite("Y\u00f6n. de\u011f. R.G. 11.09.2025/33014"),
          t(" kapsam\u0131nda, kamu hizmet alan\u0131na isabet ederek yap\u0131la\u015fma hakk\u0131n\u0131 kaybeden verici parsellerin imar haklar\u0131n\u0131n al\u0131c\u0131 parsellere aktar\u0131l\u0131p aktar\u0131lamayaca\u011f\u0131 de\u011ferlendirilmi\u015ftir."),
        ], { indent: true }),
        placeholder("Verici-al\u0131c\u0131 parsel analizi, imar hakk\u0131 aktar\u0131m cetveli ve de\u011fer dengelemesi eklenecektir."),

        heading2("8.4. Tahsis Hakkaniyeti De\u011ferlendirmesi"),
        richParagraph([
          t("Yap\u0131lan tahsislerin hakkaniyet ilkesine uygunlu\u011fu, Anayasa\u2019n\u0131n 35. maddesinde g\u00fcvence alt\u0131na al\u0131nan m\u00fclkiyet hakk\u0131"),
          cite("AY md. 35"),
          t(" ve Avrupa \u0130nsan Haklar\u0131 S\u00f6zle\u015fmesi Ek Protokol 1, Madde 1"),
          cite("A\u0130HS EP-1, md. 1"),
          t(" \u00e7er\u00e7evesinde, yerindelik oran\u0131 ve m\u00fclkiyet de\u011feri baz\u0131nda de\u011ferlendirilmi\u015ftir."),
        ], { indent: true }),
        placeholder("Art\u0131k parsel analizi, hisseli tahsis detaylar\u0131 ve de\u011fer dengelemesi eklenecektir."),

        // ===== 9. TEKNİK AÇIKLAMALAR =====
        heading1("9. TEKN\u0130K A\u00c7IKLAMALAR"),

        heading2("9.1. Koordinat Sistemi ve Datum"),
        richParagraph([
          t("T\u00fcm \u00f6l\u00e7\u00fcm ve hesaplamalar, TKGM 2025/4 say\u0131l\u0131 Genelge\u2019nin 6. ve 7. maddeleri uyar\u0131nca"),
          cite("TKGM 2025/4, md. 6-7"),
          t(" ITRF96 (2005.0 Epo\u011fu) datum\u0131nda, TUREF koordinat sisteminde ger\u00e7ekle\u015ftirilmi\u015ftir. Projeksiyon y\u00f6ntemi olarak UTM (Universal Transverse Mercator) kullan\u0131lm\u0131\u015f olup, \u00e7al\u0131\u015fma alan\u0131 [....] dilim i\u00e7erisindedir."),
        ], { indent: true }),

        heading2("9.2. \u00d6l\u00e7me Y\u00f6ntemi ve Cihazlar"),
        richParagraph([
          t("Arazi \u00f6l\u00e7\u00fcmleri, B\u00fcy\u00fck \u00d6l\u00e7ekli Harita ve Harita Bilgileri \u00dcretim Y\u00f6netmeli\u011fi"),
          cite("R.G. 15.07.2005/25876"),
          t(" ve TKGM 2025/4 say\u0131l\u0131 Genelge\u2019nin ilgili teknik h\u00fck\u00fcmlerine"),
          cite("TKGM 2025/4, B\u00f6l\u00fcm 2"),
          t(" uygun olarak [GNSS/Total Station/vb.] kullan\u0131larak ger\u00e7ekle\u015ftirilmi\u015ftir."),
        ], { indent: true }),
        emptyLine(),

        styledTable(
          ["Parametre", "De\u011fer", "Dayanak"],
          [
            ["Koordinat Sistemi", "ITRF96 / TUREF", "TKGM 2025/4, md. 6"],
            ["Projeksiyon", "UTM [....] Dilim", "TKGM 2025/4, md. 7"],
            ["\u00d6l\u00e7\u00fcm Y\u00f6ntemi", "[GNSS / Takeometrik]", "B\u00d6HHB\u00dcY md. 35-40"],
            ["Yaz\u0131l\u0131m", "[NetCAD / AutoCAD / vb.]", "-"],
            ["Hassasiyet", "[....] cm", "TKGM 2025/4, md. 8"],
          ]
        ),

        emptyLine(),
        heading2("9.3. Pafta ve Harita Bilgileri"),
        richParagraph([
          t("Parselasyon plan\u0131 paftalar\u0131 [1/1000] \u00f6l\u00e7ekli olarak haz\u0131rlanm\u0131\u015f olup, TKGM 2025/4 say\u0131l\u0131 Genelge\u2019nin pafta d\u00fczenleme standartlar\u0131na"),
          cite("TKGM 2025/4, B\u00f6l\u00fcm 4"),
          t(" ve Y\u00f6netmeli\u011fin 22. maddesinin (c) ve (d) bentlerinde belirtilen gerekliliklere"),
          cite("Y\u00f6n. md. 22/1-c,d"),
          t(" uygun olarak d\u00fczenlenmi\u015ftir."),
        ], { indent: true }),
        placeholder("Pafta indeksi, \u00f6l\u00e7ek bilgileri ve say\u0131sal veri formatlar\u0131 eklenecektir."),

        // ===== 10. İLGİLİ KURUM GÖRÜŞLERİ =====
        heading1("10. \u0130LG\u0130L\u0130 KURUM VE KURULU\u015e G\u00d6R\u00dc\u015eLER\u0130"),

        richParagraph([
          t("Parselasyon plan\u0131n\u0131n haz\u0131rlanmas\u0131 s\u00fcrecinde, Y\u00f6netmeli\u011fin 22. maddesinin birinci f\u0131kras\u0131n\u0131n (e) bendi"),
          cite("Y\u00f6n. md. 22/1-e"),
          t(" gere\u011fince gerekli g\u00f6r\u00fclen hallerde ilgili kurum ve kurulu\u015flardan al\u0131nan g\u00f6r\u00fc\u015fler a\u015fa\u011f\u0131da \u00f6zetlenmektedir:"),
        ], { indent: true }),
        emptyLine(),

        styledTable(
          ["Kurum/Kurulu\u015f", "G\u00f6r\u00fc\u015f Tarihi", "Say\u0131 No", "Yasal Dayanak", "\u00d6zet"],
          [
            ["Tapu M\u00fcd\u00fcrl\u00fc\u011f\u00fc", "[...]", "[...]", "TKGM 2025/4", "[...]"],
            ["Kadastro M\u00fcd\u00fcrl\u00fc\u011f\u00fc", "[...]", "[...]", "3402 s.K.", "[...]"],
            ["Koruma B\u00f6lge Kurulu", "[...]", "[...]", "2863 s.K.", "[...]"],
            ["DS\u0130", "[...]", "[...]", "6200 s.K.", "[...]"],
            ["TEDA\u015e/TE\u0130A\u015e", "[...]", "[...]", "6446 s.K.", "[...]"],
            ["Do\u011falgaz Da\u011f\u0131t\u0131m", "[...]", "[...]", "4646 s.K.", "[...]"],
            ["Orman \u0130\u015fl. M\u00fcd.", "[...]", "[...]", "6831 s.K.", "[...]"],
          ]
        ),

        // ===== 11. SONUÇ VE DEĞERLENDİRME =====
        emptyLine(),
        heading1("11. SONU\u00c7 VE DE\u011eERLEND\u0130RME"),

        richParagraph([
          t("\u0130\u015fbu rapor ile sunulan parselasyon plan\u0131, 3194 say\u0131l\u0131 \u0130mar Kanunu\u2019nun 18. maddesi"),
          cite("3194 s.K. md. 18"),
          t(", Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelik"),
          cite("R.G. 22.02.2020/31047; de\u011f. R.G. 11.09.2025/33014"),
          t(" ve TKGM 2025/4 say\u0131l\u0131 Tescile Konu Harita ve Planlar\u0131n Yap\u0131m\u0131 ve Kontrol\u00fc Genelgesi"),
          cite("TKGM, 12.05.2025"),
          t(" h\u00fck\u00fcmleri \u00e7er\u00e7evesinde, uygulama imar plan\u0131 kararlar\u0131na uygun olarak haz\u0131rlanm\u0131\u015ft\u0131r."),
        ], { indent: true }),

        richParagraph([
          t("D\u00fczenleme sahas\u0131ndaki toplam [....] adet kadastro parselinden [....] adet imar parseli olu\u015fturulmu\u015f olup, DOP oran\u0131 %[....] olarak hesaplanm\u0131\u015ft\u0131r. Hesaplanan DOP oran\u0131, 3194 say\u0131l\u0131 Kanun\u2019un \u00f6ng\u00f6rd\u00fc\u011f\u00fc %45 yasal \u00fcst s\u0131n\u0131r\u0131n\u0131n alt\u0131ndad\u0131r. T\u00fcm tahsis i\u015flemleri yerindelik, hakkaniyet ve yasal \u00e7er\u00e7eveye uygunluk ilkeleri do\u011frultusunda ger\u00e7ekle\u015ftirilmi\u015ftir."),
        ], { indent: true }),

        richParagraph([
          t("Parselasyon plan\u0131n\u0131n onaylanarak tescil edilmek \u00fczere Tapu M\u00fcd\u00fcrl\u00fc\u011f\u00fc\u2019ne g\u00f6nderilmesi hususunda gere\u011fini arz ederiz."),
        ], { indent: true }),

        emptyLine(),
        emptyLine(),
        separatorLine(),

        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [t("Haz\u0131rlayan", { bold: true, color: COLORS.primary })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [t("[Ad\u0131 Soyad\u0131]", { color: COLORS.grayText })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [t("[Unvan\u0131 / Oda Sicil No]", { color: COLORS.grayText })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 200 },
          children: [t("\u0130mza / Ka\u015fe", { color: COLORS.grayText })],
        }),

        separatorLine(),

        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [t("Onaylayan", { bold: true, color: COLORS.primary })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [t("[Yetkili Ad\u0131 Soyad\u0131]", { color: COLORS.grayText })],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 200 },
          children: [t("[Kurum / G\u00f6rev]", { color: COLORS.grayText })],
        }),

        // ===== 12. EKLER =====
        heading1("12. EKLER"),
        bodyText("Ek-1: D\u00fczenleme S\u0131n\u0131r\u0131 Koordinat Listesi"),
        bodyText("Ek-2: Kadastro Parselleri Tapu Kay\u0131t \u00d6zetleri"),
        bodyText("Ek-3: M\u00fclkiyet Raporu (TKGM 2025/4 format\u0131nda)"),
        bodyText("Ek-4: Uygulama \u0130mar Plan\u0131 Paftas\u0131 (Onayl\u0131)"),
        bodyText("Ek-5: Parselasyon Plan\u0131 Paftas\u0131"),
        bodyText("Ek-6: Ada R\u00f6leve Krokileri"),
        bodyText("Ek-7: Alan Hesap \u00c7izelgeleri ve Dengeleme Tablosu"),
        bodyText("Ek-8: DOP Hesap Detaylar\u0131"),
        bodyText("Ek-9: Da\u011f\u0131t\u0131m ve Tahsis Cetvelleri"),
        bodyText("Ek-10: Ask\u0131 Da\u011f\u0131t\u0131m Cetveli (PARS\u0130D format\u0131)"),
        bodyText("Ek-11: \u0130mar Hakk\u0131 Aktar\u0131m\u0131 De\u011ferlendirme Cetveli"),
        bodyText("Ek-12: \u0130lgili Kurum G\u00f6r\u00fc\u015f Yaz\u0131lar\u0131"),
        bodyText("Ek-13: Fiili Durum Foto\u011fraflar\u0131"),
        bodyText("Ek-14: \u00d6l\u00e7\u00fcm Cihaz\u0131 Kalibrasyon Belgeleri"),
        bodyText("Ek-15: Say\u0131sal Veriler (CD/USB)"),

        // ===== 13. KAYNAKÇA =====
        emptyLine(),
        heading1("13. KAYNAK\u00c7A"),

        heading2("13.1. Kanunlar"),
        bodyText("\u2022  3194 Say\u0131l\u0131 \u0130mar Kanunu, R.G. 09.05.1985/18749"),
        bodyText("\u2022  2942 Say\u0131l\u0131 Kamula\u015ft\u0131rma Kanunu, R.G. 08.11.1983/18215"),
        bodyText("\u2022  5393 Say\u0131l\u0131 Belediye Kanunu, R.G. 13.07.2005/25874"),
        bodyText("\u2022  4721 Say\u0131l\u0131 T\u00fcrk Medeni Kanunu, R.G. 08.12.2001/24607"),
        bodyText("\u2022  3402 Say\u0131l\u0131 Kadastro Kanunu, R.G. 09.07.1987/19512"),
        bodyText("\u2022  2863 Say\u0131l\u0131 K\u00fclt\u00fcr ve Tabiat Varl\u0131klar\u0131n\u0131 Koruma Kanunu, R.G. 23.07.1983/18113"),
        bodyText("\u2022  6831 Say\u0131l\u0131 Orman Kanunu, R.G. 08.09.1956/9402"),
        bodyText("\u2022  3621 Say\u0131l\u0131 K\u0131y\u0131 Kanunu, R.G. 17.04.1990/20495"),
        bodyText("\u2022  7269 Say\u0131l\u0131 Umumi Hayata M\u00fcessir Afetler Kanunu, R.G. 25.05.1959/10213"),

        heading2("13.2. Y\u00f6netmelikler"),
        bodyText("\u2022  Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelik, R.G. 22.02.2020/31047"),
        bodyText("\u2022  Arazi ve Arsa D\u00fcz. Hk. Y\u00f6n. De\u011fi\u015fikli\u011fi, R.G. 11.09.2025/33014"),
        bodyText("\u2022  Mek\u00e2nsal Planlar Yap\u0131m Y\u00f6netmeli\u011fi, R.G. 14.06.2014/29030"),
        bodyText("\u2022  Planl\u0131 Alanlar \u0130mar Y\u00f6netmeli\u011fi, R.G. 03.07.2017/30113"),
        bodyText("\u2022  B\u00fcy\u00fck \u00d6l\u00e7ekli Harita ve Harita Bilgileri \u00dcretim Y\u00f6netmeli\u011fi, R.G. 15.07.2005/25876"),

        heading2("13.3. Genelgeler"),
        bodyText("\u2022  TKGM 2025/4 Say\u0131l\u0131 Tescile Konu Harita ve Planlar\u0131n Yap\u0131m\u0131 ve Kontrol\u00fc Genelgesi, 12.05.2025, E-32381507-170.03.02-16435018"),
        bodyText("\u2022  TKGM 2023/4 Say\u0131l\u0131 Genelge (m\u00fclga)"),

        heading2("13.4. Uluslararas\u0131 S\u00f6zle\u015fmeler"),
        bodyText("\u2022  T\u00fcrkiye Cumhuriyeti Anayasas\u0131, Madde 35 (M\u00fclkiyet Hakk\u0131)"),
        bodyText("\u2022  Avrupa \u0130nsan Haklar\u0131 S\u00f6zle\u015fmesi, Ek Protokol 1, Madde 1 (M\u00fclkiyetin Korunmas\u0131)"),

        heading2("13.5. Yarg\u0131 Kararlar\u0131"),
        bodyText("\u2022  Anayasa Mahkemesi, 2018/53 E., 2021/31 K."),
        bodyText("\u2022  Dan\u0131\u015ftay 6. Dairesi, 2019/12856 E."),
        placeholder("G\u00fcncel emsal kararlar eklenecektir."),
      ],
    },
  ],
});

const outputPath = "C:/Users/STUUX/Desktop/Cloude_Deneme/Parselasyon_Plani_Aciklama_Raporu.docx";

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log("Atifli rapor olusturuldu: " + outputPath);
});
