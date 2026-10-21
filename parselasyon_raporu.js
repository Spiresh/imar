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
        // Top decorative bar
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
        // Accent line
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
        // Accent line below title
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 },
          border: {
            bottom: { style: BorderStyle.SINGLE, size: 4, color: COLORS.accent },
          },
          children: [new TextRun({ text: "" })],
        }),
        emptyLine(),
        // Info box on cover
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
            new TextRun({
              text: "Haz\u0131rlayan: [Harita M\u00fchendisi / \u015eehir Planc\u0131s\u0131 Ad\u0131 Soyad\u0131]",
              size: 22,
              color: COLORS.grayText,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [
            new TextRun({
              text: "Oda Sicil No: [.........]",
              size: 22,
              color: COLORS.grayText,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [
            new TextRun({
              text: "Tarih: [GG.AA.YYYY]",
              size: 22,
              color: COLORS.grayText,
              font: "Calibri",
            }),
          ],
        }),
        // Bottom bar
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
        page: {
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
        },
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
              size: 18,
              color: COLORS.grayText,
              font: "Calibri",
              italics: true,
            }),
          ],
        }),
      ],
    },

    // ===== ANA BÖLÜMLER =====
    {
      properties: {
        page: {
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
        },
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
                new TextRun({
                  text: "Parselasyon Plan\u0131 A\u00e7\u0131klama Raporu",
                  size: 16,
                  color: COLORS.grayText,
                  font: "Calibri",
                  italics: true,
                }),
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
                new TextRun({
                  text: "Sayfa ",
                  size: 16,
                  color: COLORS.grayText,
                  font: "Calibri",
                }),
                new TextRun({
                  children: [PageNumber.CURRENT],
                  size: 16,
                  color: COLORS.primary,
                  font: "Calibri",
                  bold: true,
                }),
              ],
            }),
          ],
        }),
      },
      children: [
        // ===== 1. GİRİŞ VE AMAÇ =====
        heading1("1. G\u0130R\u0130\u015e VE AMA\u00c7"),

        heading2("1.1. Raporun Amac\u0131"),
        bodyText(
          "\u0130\u015fbu rapor, 3194 say\u0131l\u0131 \u0130mar Kanunu\u2019nun 18. maddesi ve Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelik h\u00fck\u00fcmleri \u00e7er\u00e7evesinde haz\u0131rlanan parselasyon plan\u0131n\u0131n teknik, hukuki ve idari boyutlar\u0131n\u0131 a\u00e7\u0131klamak amac\u0131yla kaleme al\u0131nm\u0131\u015ft\u0131r. Rapor; d\u00fczenleme sahas\u0131n\u0131n mevcut durumunu, uygulanan imar plan\u0131 kararlar\u0131n\u0131, d\u00fczenleme ortakl\u0131k pay\u0131 (DOP) hesaplamalar\u0131n\u0131 ve da\u011f\u0131t\u0131m-tahsis i\u015flemlerini sistematik bi\u00e7imde ortaya koymaktad\u0131r.",
          { indent: true }
        ),

        heading2("1.2. Raporun Kapsam\u0131"),
        bodyText(
          "Rapor, d\u00fczenleme sahas\u0131 i\u00e7erisindeki t\u00fcm kadastro parsellerinin imar parsellerine d\u00f6n\u00fc\u015ft\u00fcr\u00fclmesi s\u00fcrecini kapsamaktad\u0131r. Bu \u00e7er\u00e7evede m\u00fclkiyet yap\u0131s\u0131, arazi kullan\u0131m durumu, altyap\u0131 verileri ve plan kararlar\u0131 bir b\u00fct\u00fcn olarak de\u011ferlendirilmi\u015ftir.",
          { indent: true }
        ),

        heading2("1.3. \u00c7al\u0131\u015fma Y\u00f6ntemi"),
        bodyText(
          "Parselasyon plan\u0131n\u0131n haz\u0131rlanmas\u0131nda; arazi \u00f6l\u00e7\u00fcm verileri, kadastro kay\u0131tlar\u0131, tapu sicil bilgileri, hali haz\u0131r harita verileri, uygulama imar plan\u0131 paftalar\u0131 ile ilgili kurum ve kurulu\u015f g\u00f6r\u00fc\u015fleri esas al\u0131nm\u0131\u015ft\u0131r. T\u00fcm hesaplamalar NetCAD / AutoCAD yaz\u0131l\u0131mlar\u0131 ile ITRF96/TUREF koordinat sisteminde ger\u00e7ekle\u015ftirilmi\u015ftir.",
          { indent: true }
        ),
        placeholder("Kullan\u0131lan yaz\u0131l\u0131m, koordinat sistemi ve veri kaynaklar\u0131 detayland\u0131r\u0131lacakt\u0131r."),

        // ===== 2. YASAL DAYANAK =====
        heading1("2. YASAL DAYANAK VE MEVZUAT \u00c7ER\u00c7EVES\u0130"),

        heading2("2.1. Temel Mevzuat"),
        bodyText(
          "Parselasyon plan\u0131n\u0131n haz\u0131rlanmas\u0131nda a\u015fa\u011f\u0131da belirtilen yasal d\u00fczenlemeler esas al\u0131nm\u0131\u015ft\u0131r:",
          { indent: true }
        ),
        emptyLine(),

        styledTable(
          ["S\u0131ra", "Mevzuat", "Konu"],
          [
            ["1", "3194 Say\u0131l\u0131 \u0130mar Kanunu (Madde 18)", "Arazi ve arsa d\u00fczenlemesi"],
            ["2", "Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelik (2020)", "Uygulama usul ve esaslar\u0131"],
            ["3", "3194 Say\u0131l\u0131 \u0130mar Kanunu (Madde 15-16)", "Parsellerin imar durumuna uygunlu\u011fu"],
            ["4", "2942 Say\u0131l\u0131 Kamula\u015ft\u0131rma Kanunu", "Kamula\u015ft\u0131rma i\u015flemleri"],
            ["5", "5393 Say\u0131l\u0131 Belediye Kanunu", "\u0130dari yetki ve onay"],
            ["6", "4721 Say\u0131l\u0131 T\u00fcrk Medeni Kanunu", "M\u00fclkiyet hakk\u0131"],
            ["7", "3402 Say\u0131l\u0131 Kadastro Kanunu", "Kadastro tespiti"],
            ["8", "TKGM Genelgeleri", "Tescil i\u015flemleri"],
          ]
        ),

        emptyLine(),
        heading2("2.2. \u0130lgili Yarg\u0131 Kararlar\u0131"),
        bodyText(
          "Parselasyon uygulamalar\u0131na ili\u015fkin Dan\u0131\u015ftay ve Anayasa Mahkemesi kararlar\u0131 da dikkate al\u0131nm\u0131\u015ft\u0131r. \u00d6zellikle DOP oran\u0131n\u0131n \u00fcst s\u0131n\u0131r\u0131na, tahsis adaletine ve m\u00fclkiyet hakk\u0131n\u0131n korunmas\u0131na y\u00f6nelik i\u00e7tihatlar g\u00f6zetilmi\u015ftir.",
          { indent: true }
        ),
        placeholder("Emsal yarg\u0131 kararlar\u0131 ve i\u00e7tihatlar bu b\u00f6l\u00fcmde belirtilecektir."),

        // ===== 3. DÜZENLEME SAHASI =====
        heading1("3. D\u00dcZENLEME SAHASI B\u0130LG\u0130LER\u0130"),

        heading2("3.1. Co\u011frafi Konum ve S\u0131n\u0131rlar"),
        bodyText(
          "D\u00fczenleme sahas\u0131, [il/il\u00e7e/mahalle] s\u0131n\u0131rlar\u0131 dahilinde yer almakta olup, do\u011fusunda [...], bat\u0131s\u0131nda [...], kuzeyinde [...] ve g\u00fcneyinde [...] ile \u00e7evrilidir. Saahan\u0131n toplam alan\u0131 [....] m\u00b2 olup, ITRF96 koordinat sisteminde tan\u0131mlanm\u0131\u015ft\u0131r.",
          { indent: true }
        ),
        placeholder("D\u00fczenleme s\u0131n\u0131r\u0131 koordinatlar\u0131, vaziyet plan\u0131 ve uydu g\u00f6r\u00fcnt\u00fcs\u00fc eklenecektir."),

        heading2("3.2. D\u00fczenleme S\u0131n\u0131r\u0131n\u0131n Belirlenmesi"),
        bodyText(
          "D\u00fczenleme s\u0131n\u0131r\u0131, Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmeli\u011fin 5. maddesi h\u00fck\u00fcmleri do\u011frultusunda, uygulama imar plan\u0131 s\u0131n\u0131rlar\u0131, m\u00fclkiyet yap\u0131s\u0131 ve arazi kullan\u0131m durumu g\u00f6zetilerek belirlenmi\u015ftir. S\u0131n\u0131r\u0131n belirlenmesinde b\u00fct\u00fcnl\u00fck ilkesi ve uygulanabilirlik kriterleri esas al\u0131nm\u0131\u015ft\u0131r.",
          { indent: true }
        ),

        heading2("3.3. Topografya ve Arazi Yap\u0131s\u0131"),
        bodyText(
          "D\u00fczenleme sahas\u0131n\u0131n topografik yap\u0131s\u0131, e\u011fim analizi sonu\u00e7lar\u0131 ve jeolojik-jeoteknik verileri bu b\u00f6l\u00fcmde de\u011ferlendirilmektedir. Saahan\u0131n ortalama kotu [....] m, en d\u00fc\u015f\u00fck kotu [....] m ve en y\u00fcksek kotu [....] m olarak tespit edilmi\u015ftir.",
          { indent: true }
        ),
        placeholder("Topografik kesit, e\u011fim haritas\u0131 ve jeolojik rapor \u00f6zeti eklenecektir."),

        // ===== 4. MÜLK\u0130YET DURUMU =====
        heading1("4. M\u00dcLK\u0130YET DURUMU ANAL\u0130Z\u0130"),

        heading2("4.1. Kadastral Yap\u0131 ve Parsel Envanteri"),
        bodyText(
          "D\u00fczenleme sahas\u0131 i\u00e7erisinde toplam [....] adet kadastro parseli bulunmakta olup, bu parsellerin toplam alan\u0131 [....] m\u00b2\u2019dir. Parsellerin ada ve parsel numaralar\u0131 ile y\u00fcz\u00f6l\u00e7\u00fcmleri a\u015fa\u011f\u0131daki tabloda sunulmaktad\u0131r.",
          { indent: true }
        ),
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
        bodyText(
          "D\u00fczenleme sahas\u0131ndaki parsellerin m\u00fclkiyet t\u00fcr\u00fc itibar\u0131yla da\u011f\u0131l\u0131m\u0131 a\u015fa\u011f\u0131da sunulmaktad\u0131r:",
          { indent: true }
        ),
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
        bodyText(
          "D\u00fczenleme sahas\u0131 i\u00e7erisinde 2863 say\u0131l\u0131 K\u00fclt\u00fcr ve Tabiat Varl\u0131klar\u0131n\u0131 Koruma Kanunu, 6831 say\u0131l\u0131 Orman Kanunu, 3621 say\u0131l\u0131 K\u0131y\u0131 Kanunu ve sair \u00f6zel kanunlar kapsam\u0131nda koruma alt\u0131nda bulunan alan ve parsellerin varl\u0131\u011f\u0131 incelenmi\u015ftir.",
          { indent: true }
        ),
        placeholder("\u00d6zel kanunla korunan alan listesi, sit alan\u0131 kararlar\u0131 ve ilgili kurul g\u00f6r\u00fc\u015fleri eklenecektir."),

        heading2("4.4. \u00d6nceki \u0130mar Uygulamalar\u0131"),
        bodyText(
          "D\u00fczenleme sahas\u0131 i\u00e7erisinde daha \u00f6nce 3194 say\u0131l\u0131 Kanun\u2019un 15., 16. veya 18. maddelerine g\u00f6re i\u015flem g\u00f6rm\u00fc\u015f parsellerin tespiti yap\u0131lm\u0131\u015ft\u0131r. Bu parsellere ait terk oranlar\u0131 ve m\u00fcktesep haklar de\u011ferlendirilmi\u015ftir.",
          { indent: true }
        ),
        emptyLine(),

        styledTable(
          ["Ada/Parsel", "Uygulama T\u00fcr\u00fc", "Tarih", "Terk Oran\u0131 (%)", "A\u00e7\u0131klama"],
          [
            ["[...]", "Md. 18", "[...]", "[...]", "[...]"],
            ["[...]", "Md. 15-16", "[...]", "[...]", "[...]"],
            ["[...]", "Ay\u0131rma-Birle\u015ftirme", "[...]", "[...]", "[...]"],
          ]
        ),

        emptyLine(),
        heading2("4.5. Fiili Durum Tespiti"),
        bodyText(
          "D\u00fczenleme sahas\u0131ndaki yap\u0131la\u015fma durumu, mevcut kullan\u0131m bi\u00e7imleri ve fiili i\u015fgaller yerinde yap\u0131lan inceleme ile tespit edilmi\u015ftir. Muhdesat tespitleri, encr\u00f6\u015fman durumlar\u0131 ve yap\u0131 envanter bilgileri bu b\u00f6l\u00fcmde sunulmaktad\u0131r.",
          { indent: true }
        ),
        placeholder("Fiili durum foto\u011fraflar\u0131, muhdesat listesi ve encr\u00f6\u015fman analizi eklenecektir."),

        // ===== 5. İMAR PLANI BİLGİLERİ =====
        heading1("5. \u0130MAR PLANI B\u0130LG\u0130LER\u0130 VE PLAN KARARLARI"),

        heading2("5.1. Uygulama \u0130mar Plan\u0131 Bilgileri"),
        emptyLine(),
        infoBox([
          { label: "Plan\u0131n Onay Tarihi", value: "[GG.AA.YYYY]" },
          { label: "Plan\u0131n Kesinle\u015fme Tarihi", value: "[GG.AA.YYYY]" },
          { label: "Plan \u00d6l\u00e7e\u011fi", value: "1/1000" },
          { label: "Onaylayan Kurum", value: "[.........]" },
          { label: "Meclis Karar No", value: "[.........]" },
          { label: "Plan Revizyonu", value: "[Var/Yok \u2013 Tarih]" },
        ]),
        emptyLine(),

        heading2("5.2. Arazi Kullan\u0131m Kararlar\u0131"),
        bodyText(
          "Uygulama imar plan\u0131nda d\u00fczenleme sahas\u0131 i\u00e7in \u00f6ng\u00f6r\u00fclen arazi kullan\u0131m kararlar\u0131 ve fonksiyon alanlar\u0131 a\u015fa\u011f\u0131da \u00f6zetlenmektedir:",
          { indent: true }
        ),
        emptyLine(),

        styledTable(
          ["Kullan\u0131m T\u00fcr\u00fc", "Plan Sembol\u00fc", "Alan (m\u00b2)", "Oran (%)"],
          [
            ["Konut Alan\u0131", "K", "[...]", "[...]"],
            ["Ticaret Alan\u0131", "T", "[...]", "[...]"],
            ["Ye\u015fil Alan (Park)", "YA", "[...]", "[...]"],
            ["E\u011fitim Tesisi", "E", "[...]", "[...]"],
            ["Sa\u011fl\u0131k Tesisi", "S", "[...]", "[...]"],
            ["Dini Tesis Alan\u0131", "DT", "[...]", "[...]"],
            ["Yol", "Y", "[...]", "[...]"],
            ["Otopark", "OP", "[...]", "[...]"],
            ["Teknik Altyap\u0131", "TA", "[...]", "[...]"],
            ["TOPLAM", "-", "[...]", "100"],
          ]
        ),

        emptyLine(),
        heading2("5.3. Yap\u0131la\u015fma Ko\u015fullar\u0131"),
        bodyText(
          "\u0130mar plan\u0131nda \u00f6ng\u00f6r\u00fclen yap\u0131la\u015fma ko\u015fullar\u0131 (TAKS, KAKS/emsal, Hmaks, \u00e7ekme mesafeleri, kat adedi) a\u015fa\u011f\u0131daki tabloda sunulmaktad\u0131r:",
          { indent: true }
        ),
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
        bodyText(
          "Uygulama imar plan\u0131na ait genel ve \u00f6zel plan notlar\u0131 bu b\u00f6l\u00fcmde detayl\u0131 olarak aktar\u0131lmaktad\u0131r. Parselasyon uygulamas\u0131n\u0131 do\u011frudan etkileyen h\u00fck\u00fcmler ayr\u0131ca de\u011ferlendirilmi\u015ftir.",
          { indent: true }
        ),
        placeholder("Plan notlar\u0131n\u0131n tamam\u0131 ve \u00f6zel h\u00fck\u00fcmler bu b\u00f6l\u00fcme eklenecektir."),

        // ===== 6. DOP HESABI =====
        heading1("6. D\u00dcZENLEME ORTAKLIK PAYI (DOP) HESABI"),

        heading2("6.1. DOP Hesaplama Y\u00f6ntemi ve Esaslar\u0131"),
        bodyText(
          "D\u00fczenleme ortakl\u0131k pay\u0131, 3194 say\u0131l\u0131 Kanun\u2019un 18. maddesi ve Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmeli\u011fin 11\u201316. maddeleri \u00e7er\u00e7evesinde hesaplanm\u0131\u015ft\u0131r. DOP; yol, meydan, park, ye\u015fil alan, otopark, \u00e7ocuk bah\u00e7esi, cami, karakol ve Milli E\u011fitim Bakanl\u0131\u011f\u0131na ba\u011fl\u0131 ilk ve orta\u00f6\u011fretim kurumlar\u0131 i\u00e7in ayr\u0131lan alanlar\u0131 kar\u015f\u0131lamak \u00fczere hesaplanm\u0131\u015ft\u0131r.",
          { indent: true }
        ),

        heading2("6.2. DOP Hesap Tablosu"),
        emptyLine(),

        styledTable(
          ["Kalem", "De\u011fer"],
          [
            ["D\u00fczenlemeye Giren Toplam Alan", "[.........] m\u00b2"],
            ["Kamu Alanlar\u0131 \u0130htiyac\u0131 (Yol)", "[.........] m\u00b2"],
            ["Kamu Alanlar\u0131 \u0130htiyac\u0131 (Park/Ye\u015fil)", "[.........] m\u00b2"],
            ["Kamu Alanlar\u0131 \u0130htiyac\u0131 (E\u011fitim)", "[.........] m\u00b2"],
            ["Kamu Alanlar\u0131 \u0130htiyac\u0131 (Dini Tesis)", "[.........] m\u00b2"],
            ["Kamu Alanlar\u0131 \u0130htiyac\u0131 (Di\u011fer)", "[.........] m\u00b2"],
            ["Toplam DOP Alan\u0131", "[.........] m\u00b2"],
            ["DOP Oran\u0131", "% [...]"],
          ]
        ),

        emptyLine(),
        heading2("6.3. DOP Oran\u0131 De\u011ferlendirmesi"),
        bodyText(
          "Hesaplanan DOP oran\u0131 %[....] olup, Kanun\u2019un \u00f6ng\u00f6rd\u00fc\u011f\u00fc %45 \u00fcst s\u0131n\u0131r\u0131n\u0131n alt\u0131nda kalmaktad\u0131r. DOP ile kar\u015f\u0131lanamayan kamu hizmeti alanlar\u0131 i\u00e7in kamula\u015ft\u0131rma yoluna gidilmesi gerekmektedir.",
          { indent: true }
        ),
        placeholder("DOP oran\u0131 yeterlili\u011fi, alternatif senaryolar ve kamula\u015ft\u0131rma ihtiyac\u0131 de\u011ferlendirilecektir."),

        heading2("6.4. \u00d6nceki Uygulamalardan D\u00fc\u015f\u00fclecek Terkler"),
        bodyText(
          "Daha \u00f6nce 15/16 veya 18. madde uygulamas\u0131 g\u00f6rm\u00fc\u015f parsellerin mevcut terk oranlar\u0131, m\u00fcktesep hak olarak dikkate al\u0131nm\u0131\u015f ve DOP hesab\u0131nda ilgili oranlar d\u00fc\u015f\u00fclm\u00fc\u015ft\u00fcr.",
          { indent: true }
        ),

        // ===== 7. ALAN HESAPLARI VE DENGELEME =====
        heading1("7. ALAN HESAPLARI VE DENGELEME"),

        heading2("7.1. Kadastro-\u0130mar Alan Kar\u015f\u0131la\u015ft\u0131rmas\u0131"),
        bodyText(
          "D\u00fczenleme sahas\u0131na giren kadastro parsellerinin toplam alanlar\u0131 ile olu\u015fturulan imar parsellerinin toplam alanlar\u0131 aras\u0131ndaki dengeleme tablosu a\u015fa\u011f\u0131da sunulmaktad\u0131r.",
          { indent: true }
        ),
        emptyLine(),

        styledTable(
          ["A\u00e7\u0131klama", "Alan (m\u00b2)"],
          [
            ["D\u00fczenlemeye Giren Toplam Kadastro Alan\u0131", "[...]"],
            ["(-) DOP Olarak Ayr\u0131lan Alan", "[...]"],
            ["(-) Kamula\u015ft\u0131r\u0131lacak Alan", "[...]"],
            ["(+) Ba\u011f\u0131\u015f / Terk Alanlar\u0131", "[...]"],
            ["(=) Da\u011f\u0131t\u0131lacak Toplam Alan", "[...]"],
            ["Olu\u015fan \u0130mar Parselleri Toplam Alan\u0131", "[...]"],
            ["Dengeleme Fark\u0131 (m\u00b2)", "[...]"],
          ]
        ),

        emptyLine(),
        heading2("7.2. Alan Hesaplama Y\u00f6ntemi"),
        bodyText(
          "Alan hesaplamalar\u0131 koordinat y\u00f6ntemi ile yap\u0131lm\u0131\u015f olup, t\u00fcm hesaplar ITRF96 koordinat sistemindedir. Say\u0131sal verilerin do\u011fruluk d\u00fczeyi ve tolerans de\u011ferleri B\u00fcy\u00fck \u00d6l\u00e7ekli Harita ve Harita Bilgileri \u00dcretim Y\u00f6netmeli\u011fi\u2019ne uygun olarak sa\u011flanm\u0131\u015ft\u0131r.",
          { indent: true }
        ),

        // ===== 8. DAĞITIM VE TAHSİS =====
        heading1("8. DA\u011eITIM VE TAHS\u0130S \u0130\u015eLEMLER\u0130"),

        heading2("8.1. Da\u011f\u0131t\u0131m \u0130lkeleri"),
        bodyText(
          "Parselasyon plan\u0131nda imar parsellerinin da\u011f\u0131t\u0131m\u0131, Y\u00f6netmeli\u011fin 17\u201321. maddeleri uyar\u0131nca a\u015fa\u011f\u0131daki ilkeler g\u00f6zetilerek yap\u0131lm\u0131\u015ft\u0131r:",
          { indent: true }
        ),
        bodyText("\u2022  Parsel sahiplerinin m\u00fcmk\u00fcn mertebe eski yerlerinden tahsisi (yerindelik ilkesi)"),
        bodyText("\u2022  M\u00fclkiyet haklar\u0131n\u0131n korunmas\u0131 ve e\u015fit muamele"),
        bodyText("\u2022  Parsellerin imar plan\u0131na uygun, yap\u0131la\u015fmaya elveri\u015fli bi\u00e7imde olu\u015fturulmas\u0131"),
        bodyText("\u2022  Asgari parsel b\u00fcy\u00fckl\u00fcklerinin sa\u011flanmas\u0131"),
        bodyText("\u2022  Hisseli tahsisin asgari d\u00fczeyde tutulmas\u0131"),

        heading2("8.2. Da\u011f\u0131t\u0131m Cetveli"),
        bodyText(
          "A\u015fa\u011f\u0131daki tabloda, kadastro parsellerinden imar parsellerine yap\u0131lan tahsis i\u015flemleri \u00f6zetlenmektedir:",
          { indent: true }
        ),
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
        heading2("8.3. Tahsis Hakkaniyeti De\u011ferlendirmesi"),
        bodyText(
          "Yap\u0131lan tahsislerin hakkaniyet ilkesine uygunlu\u011fu, yerindelik oran\u0131 ve m\u00fclkiyet de\u011feri baz\u0131nda de\u011ferlendirilmi\u015ftir. Tahsis i\u015flemlerinde ortaya \u00e7\u0131kan art\u0131k parseller ve \u00e7\u00f6z\u00fcm y\u00f6ntemleri a\u015fa\u011f\u0131da a\u00e7\u0131klanmaktad\u0131r.",
          { indent: true }
        ),
        placeholder("Art\u0131k parsel analizi, hisseli tahsis detaylar\u0131 ve de\u011fer dengelemesi eklenecektir."),

        // ===== 9. TEKNİK AÇIKLAMALAR =====
        heading1("9. TEKN\u0130K A\u00c7IKLAMALAR"),

        heading2("9.1. Koordinat Sistemi ve Datum"),
        bodyText(
          "T\u00fcm \u00f6l\u00e7\u00fcm ve hesaplamalar ITRF96 (2005.0 Epo\u011fu) datum\u0131nda, TUREF koordinat sisteminde ger\u00e7ekle\u015ftirilmi\u015ftir. Projeksiyon y\u00f6ntemi olarak UTM (Universal Transverse Mercator) kullan\u0131lm\u0131\u015f olup, \u00e7al\u0131\u015fma alan\u0131 [....] dilim i\u00e7erisindedir.",
          { indent: true }
        ),

        heading2("9.2. \u00d6l\u00e7me Y\u00f6ntemi ve Cihazlar"),
        bodyText(
          "Arazi \u00f6l\u00e7\u00fcmleri [GNSS/Total Station/vb.] kullan\u0131larak ger\u00e7ekle\u015ftirilmi\u015ftir. Kullan\u0131lan cihazlar\u0131n kalibrasyon belgeleri ekler aras\u0131nda sunulmaktad\u0131r.",
          { indent: true }
        ),
        emptyLine(),

        styledTable(
          ["Parametre", "De\u011fer"],
          [
            ["Koordinat Sistemi", "ITRF96 / TUREF"],
            ["Projeksiyon", "UTM [....] Dilim"],
            ["\u00d6l\u00e7\u00fcm Y\u00f6ntemi", "[GNSS / Takeometrik]"],
            ["Yaz\u0131l\u0131m", "[NetCAD / AutoCAD / vb.]"],
            ["Hassasiyet", "[....] cm"],
          ]
        ),

        emptyLine(),
        heading2("9.3. Pafta ve Harita Bilgileri"),
        bodyText(
          "Parselasyon plan\u0131 paftalar\u0131 [1/1000] \u00f6l\u00e7ekli olarak haz\u0131rlanm\u0131\u015f olup, TKGM standartlar\u0131na uygun \u015fekilde d\u00fczenlenen pafta indeks bilgileri a\u015fa\u011f\u0131dad\u0131r.",
          { indent: true }
        ),
        placeholder("Pafta indeksi, \u00f6l\u00e7ek bilgileri ve say\u0131sal veri formatlar\u0131 eklenecektir."),

        // ===== 10. İLGİLİ KURUM GÖRÜŞLERİ =====
        heading1("10. \u0130LG\u0130L\u0130 KURUM VE KURULU\u015e G\u00d6R\u00dc\u015eLER\u0130"),

        bodyText(
          "Parselasyon plan\u0131n\u0131n haz\u0131rlanmas\u0131 s\u00fcrecinde, Y\u00f6netmeli\u011fin 22. maddesi gere\u011fince ilgili kurum ve kurulu\u015flardan al\u0131nan g\u00f6r\u00fc\u015fler a\u015fa\u011f\u0131da \u00f6zetlenmektedir:",
          { indent: true }
        ),
        emptyLine(),

        styledTable(
          ["Kurum/Kurulu\u015f", "G\u00f6r\u00fc\u015f Tarihi", "Say\u0131 No", "\u00d6zet"],
          [
            ["Tapu M\u00fcd\u00fcrl\u00fc\u011f\u00fc", "[...]", "[...]", "[...]"],
            ["Kadastro M\u00fcd\u00fcrl\u00fc\u011f\u00fc", "[...]", "[...]", "[...]"],
            ["Koruma B\u00f6lge Kurulu", "[...]", "[...]", "[...]"],
            ["DSI", "[...]", "[...]", "[...]"],
            ["TEDA\u015e/TE\u0130A\u015e", "[...]", "[...]", "[...]"],
            ["Do\u011falgaz Da\u011f\u0131t\u0131m", "[...]", "[...]", "[...]"],
          ]
        ),

        // ===== 11. SONUÇ VE DEĞERLENDİRME =====
        emptyLine(),
        heading1("11. SONU\u00c7 VE DE\u011eERLEND\u0130RME"),

        bodyText(
          "\u0130\u015fbu rapor ile sunulan parselasyon plan\u0131, 3194 say\u0131l\u0131 \u0130mar Kanunu\u2019nun 18. maddesi ve Arazi ve Arsa D\u00fczenlemeleri Hakk\u0131nda Y\u00f6netmelik h\u00fck\u00fcmleri \u00e7er\u00e7evesinde, uygulama imar plan\u0131 kararlar\u0131na uygun olarak haz\u0131rlanm\u0131\u015ft\u0131r.",
          { indent: true }
        ),
        bodyText(
          "D\u00fczenleme sahas\u0131ndaki toplam [....] adet kadastro parselinden [....] adet imar parseli olu\u015fturulmu\u015f olup, DOP oran\u0131 %[....] olarak hesaplanm\u0131\u015ft\u0131r. T\u00fcm tahsis i\u015flemleri yerindelik, hakkaniyet ve yasal \u00e7er\u00e7eveye uygunluk ilkeleri do\u011frultusunda ger\u00e7ekle\u015ftirilmi\u015ftir.",
          { indent: true }
        ),
        bodyText(
          "Parselasyon plan\u0131n\u0131n onaylanarak tescil edilmek \u00fczere Tapu M\u00fcd\u00fcrl\u00fc\u011f\u00fc\u2019ne g\u00f6nderilmesi hususunda gere\u011fini arz ederiz.",
          { indent: true }
        ),

        emptyLine(),
        emptyLine(),
        separatorLine(),

        // İmza bloğu
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [
            new TextRun({
              text: "Haz\u0131rlayan",
              bold: true,
              size: 22,
              color: COLORS.primary,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [
            new TextRun({
              text: "[Ad\u0131 Soyad\u0131]",
              size: 22,
              color: COLORS.grayText,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [
            new TextRun({
              text: "[Unvan\u0131 / Oda Sicil No]",
              size: 22,
              color: COLORS.grayText,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 200 },
          children: [
            new TextRun({
              text: "\u0130mza / Ka\u015fe",
              size: 22,
              color: COLORS.grayText,
              font: "Calibri",
            }),
          ],
        }),

        separatorLine(),

        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [
            new TextRun({
              text: "Onaylayan",
              bold: true,
              size: 22,
              color: COLORS.primary,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 60 },
          children: [
            new TextRun({
              text: "[Yetkili Ad\u0131 Soyad\u0131]",
              size: 22,
              color: COLORS.grayText,
              font: "Calibri",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 200 },
          children: [
            new TextRun({
              text: "[Kurum / G\u00f6rev]",
              size: 22,
              color: COLORS.grayText,
              font: "Calibri",
            }),
          ],
        }),

        // ===== 12. EKLER =====
        heading1("12. EKLER"),
        bodyText("Ek-1: D\u00fczenleme S\u0131n\u0131r\u0131 Koordinat Listesi"),
        bodyText("Ek-2: Kadastro Parselleri Tapu Kay\u0131t \u00d6zetleri"),
        bodyText("Ek-3: Uygulama \u0130mar Plan\u0131 Paftas\u0131 (Onayl\u0131)"),
        bodyText("Ek-4: Parselasyon Plan\u0131 Paftas\u0131"),
        bodyText("Ek-5: Ada R\u00f6leve Krokileri"),
        bodyText("Ek-6: Alan Hesap \u00c7izelgeleri ve Dengeleme Tablosu"),
        bodyText("Ek-7: DOP Hesap Detaylar\u0131"),
        bodyText("Ek-8: Da\u011f\u0131t\u0131m ve Tahsis Cetvelleri"),
        bodyText("Ek-9: \u0130lgili Kurum G\u00f6r\u00fc\u015f Yaz\u0131lar\u0131"),
        bodyText("Ek-10: Fiili Durum Foto\u011fraflar\u0131"),
        bodyText("Ek-11: \u00d6l\u00e7\u00fcm Cihaz\u0131 Kalibrasyon Belgeleri"),
        bodyText("Ek-12: Say\u0131sal Veriler (CD/USB)"),
      ],
    },
  ],
});

const outputPath = "C:/Users/STUUX/Desktop/Cloude_Deneme/Parselasyon_Plani_Aciklama_Raporu.docx";

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log("Rapor olusturuldu: " + outputPath);
});
