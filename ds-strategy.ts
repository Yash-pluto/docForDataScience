import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  WidthType,
  ShadingType,
  LevelFormat,
} from "docx";
import fs from "fs";

const border = { style: BorderStyle.SINGLE, size: 2, color: "E0E0E0" };
const borders = { top: border, bottom: border, left: border, right: border };

const MAIN_FONT = "Calibri";

function heading1(text: string) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 120 },

    border: {
      bottom: {
        color: "2E7D32",
        space: 1,
        style: BorderStyle.SINGLE,
        size: 12,
      },
    },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 30,
        font: MAIN_FONT,
        color: "2E7D32",
      }),
    ],
  });
}

function heading2(text: string) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 80 },
    children: [
      new TextRun({
        text,
        bold: true,
        size: 24,
        font: MAIN_FONT,
        color: "37474F",
      }),
    ],
  });
}

function para(text: string, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [
      new TextRun({
        text,
        font: MAIN_FONT,
        size: 22,
        color: "424242",
        ...opts,
      }),
    ],
  });
}

function bullet(text: string, bold = false) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [
      new TextRun({ text, font: MAIN_FONT, size: 22, color: "424242", bold }),
    ],
  });
}

function emptyLine() {
  return new Paragraph({ children: [new TextRun("")] });
}

function makeTable(rows: string[][], colWidths: number[]) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: colWidths,
    rows: rows.map(
      (cells, ri) =>
        new TableRow({
          children: cells.map(
            (cell, ci) =>
              new TableCell({
                borders,
                width: { size: colWidths[ci], type: WidthType.DXA },
                shading:
                  ri === 0
                    ? { fill: "F5F5F5", type: ShadingType.CLEAR } // Softer header background
                    : { fill: "FFFFFF", type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 150, right: 150 }, // More breathing room
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: cell,
                        font: MAIN_FONT,
                        size: 21,
                        bold: ri === 0,
                        color: ri === 0 ? "212121" : "424242",
                      }),
                    ],
                  }),
                ],
              }),
          ),
        }),
    ),
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [
            new TextRun({
              text: "CSE435: Intro to Data Science",
              bold: true,
              font: MAIN_FONT,
              size: 40,
              color: "1A237E",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [
            new TextRun({
              text: "Exam Game Plan — Target: 35+ Marks",
              font: MAIN_FONT,
              size: 24,
              color: "757575",
              italics: true,
            }),
          ],
        }),

        heading1("1. The Setup"),
        para(
          "Total paper is 60 marks. We only need 35 to sit comfortably. Here is the math:",
        ),
        emptyLine(),

        makeTable(
          [
            ["Exam Component", "Weight", "Max Marks"],
            ["Internal Assessments", "40%", "40"],
            ["End Term Theory", "60%", "60"],
          ],
          [4000, 2680, 2680],
        ),
        emptyLine(),

        heading1("2. Where to Focus (Module Breakdown)"),
        para("Not all modules are created equal. Focus on III and IV first."),
        emptyLine(),

        makeTable(
          [
            ["Mod", "Topic", "Weight", "Action Plan"],
            ["I", "Basics of DS", "15%", "Skim it. Just know the definitions."],
            [
              "II",
              "Data Analytics & EDA",
              "20%",
              "Decent chunk of marks. Don't skip.",
            ],
            ["III", "Feature Selection", "25%", "URGENT — Master this."],
            ["IV", "Data Visualization", "25%", "URGENT — Master this."],
            ["V", "Ethics & Apps", "15%", "Read once the night before."],
          ],
          [800, 3200, 1500, 3860],
        ),

        emptyLine(),
        para(
          "Cheat code: Modules III & IV equal about 30 marks. Nail those, grab a few points in EDA, and you're already past 35.",
          { bold: true, color: "D84315" },
        ),
        emptyLine(),

        heading1("3. The Actual Hit List"),

        heading2("Module III: Feature Stuff (25%)"),
        bullet(
          "Filter vs. Wrapper vs. Embedded methods. Just memorize one example for each (e.g., Filter = Chi-square, Wrapper = RFE).",
        ),
        bullet(
          "Why do we even generate features? (Hint: domain knowledge makes models better).",
        ),
        bullet("Be ready to explain the customer retention example."),
        emptyLine(),

        heading2("Module IV: Graphs & Visuals (25%)"),
        bullet(
          "When to use what: Bar (categories), Line (time), Scatter (relationships), Histogram (distributions). Don't mix them up.",
        ),
        bullet(
          "Matplotlib syntax: plt.plot(), plt.bar(), etc. Keep it simple.",
        ),
        bullet(
          "If they ask an open-ended viz question, quickly sketch the chart. Graders eat that up.",
        ),
        emptyLine(),

        heading2("Module II: EDA (20%)"),
        bullet(
          "The 5-number summary and basic stats (mean, median, variance).",
        ),
        bullet("The pipeline: Collect → Clean → Explore → Model → Interpret."),
        bullet(
          "Be prepared to look at a sample dataset and write a 2-sentence conclusion.",
        ),
        emptyLine(),

        heading2("Modules I & V: The Fluff (30% combined)"),
        bullet("Know what Pandas, NumPy, and Scikit-learn actually do."),
        bullet(
          "Memorize a few ethical issues (bias, privacy) and real-world DS applications (finance, healthcare). Easy points.",
        ),
        emptyLine(),

        heading1("4. Last-Minute Reminders"),
        bullet("Do sections III and IV the second you open the paper."),
        bullet(
          "Never leave a question blank. BS a definition if you have to—partial credit is everything.",
        ),
        emptyLine(),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 400 },
          children: [
            new TextRun({
              text: "I built a full-stack portfolio in 5 hours. This exam is nothing. 🔥",
              font: MAIN_FONT,
              size: 22,
              italics: true,
              color: "9E9E9E",
            }),
          ],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("./DS_Exam_Strategy.docx", buffer);
  console.log("File saved successfully! Run it and check out the new design.");
});
