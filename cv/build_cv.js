// Builds Bright Akwasi Kodua's CV tailored for the LMI Holdings
// 2026/2027 National Service Recruitment. Run: node cv/build_cv.js
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, LevelFormat,
  TabStopType,
} = require("docx");

const ACCENT = "B3121B";   // deep red
const DARK = "222222";
const MUTED = "595959";
const FONT = "Calibri";

const PAGE_W = 11906;      // A4
const MARGIN = 850;        // ~0.59"
const CONTENT_W = PAGE_W - MARGIN * 2;

const run = (text, o = {}) => new TextRun({ text, font: FONT, size: 20, color: DARK, ...o });

const sectionHeading = (title) => new Paragraph({
  spacing: { before: 200, after: 80 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 2 } },
  children: [run(title.toUpperCase(), { bold: true, size: 22, color: ACCENT, characterSpacing: 30 })],
});

const para = (children, o = {}) => new Paragraph({ spacing: { after: 60, line: 264 }, children, ...o });

const bullet = (children) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 40, line: 252 },
  children: typeof children === "string" ? [run(children)] : children,
});

// Title on the left, meta (dates / tech / place) right-aligned on the same line.
const entryHeader = (title, right) => new Paragraph({
  spacing: { before: 100, after: 30 },
  tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
  children: [
    run(title, { bold: true, size: 21 }),
    ...(right ? [run("\t" + right, { italics: true, color: MUTED, size: 19 })] : []),
  ],
});

const subLine = (text) => para([run(text, { italics: true, color: MUTED, size: 19 })], { spacing: { after: 40 } });

const labelled = (label, value) => bullet([run(label + ": ", { bold: true }), run(value)]);

// Two-column skills grid, borderless.
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const cellBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
const half = Math.floor(CONTENT_W / 2);
const skillCell = (items) => new TableCell({
  width: { size: half, type: WidthType.DXA },
  borders: cellBorders,
  margins: { top: 0, bottom: 0, left: 0, right: 120 },
  children: items.map(([l, v]) => labelled(l, v)),
});
const skillsTable = (left, right) => new Table({
  width: { size: half * 2, type: WidthType.DXA },
  columnWidths: [half, half],
  borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideHorizontal: noBorder, insideVertical: noBorder },
  rows: [new TableRow({ children: [skillCell(left), skillCell(right)] })],
});

// Header band
const header = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 20 },
    children: [run("BRIGHT AKWASI KODUA", { bold: true, size: 40, color: DARK, characterSpacing: 40 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
    children: [run("Computer Science Graduate  |  Software Development  •  Cybersecurity  •  Networking", { size: 21, color: ACCENT, bold: true })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "F2F2F2" },
    children: [run("Accra, Ghana   |   +233 535 862 985   |   brightkoduaakwasi@gmail.com", { size: 19, color: MUTED })],
  }),
];

const body = [
  sectionHeading("Professional Summary"),
  para([run(
    "Tech-savvy Computer Science graduate from the University of Cape Coast with hands-on experience building full-stack " +
    "web applications, designing secure network infrastructure, and applying machine learning to real-world security " +
    "problems. Seeking a National Service placement with LMI Holdings for the 2026/2027 period, and eager to take part in " +
    "the Talent Development Programme, contribute to digital and IT operations across LMI's subsidiaries, and grow within a " +
    "diversified conglomerate."
  )], { alignment: AlignmentType.JUSTIFIED }),

  sectionHeading("Education"),
  entryHeader("BSc. Computer Science", "University of Cape Coast (UCC), Ghana"),
  para([run("Relevant Coursework:", { bold: true, size: 19 })], { spacing: { after: 30 } }),
  labelled("Software Engineering", "software development lifecycle, requirements analysis, system design"),
  labelled("Database Management", "relational database design, SQL, data integrity and security"),
  labelled("Cryptography & Network Security", "secure communication protocols"),
  labelled("Systems Administration", "business continuity / disaster recovery planning, backup strategies, network troubleshooting"),
  labelled("Operations Research", "decision modelling, optimisation, applied problem-solving"),

  sectionHeading("Technical Skills"),
  skillsTable(
    [
      ["Programming", "Python, JavaScript, Node.js"],
      ["Web Development", "React, Express, Flask"],
      ["Databases", "PostgreSQL, Supabase, SQL"],
    ],
    [
      ["Cybersecurity", "vulnerability scanning, file integrity monitoring, phishing detection"],
      ["Systems & Networking", "Unix/Linux administration, VLANs, VPN configuration, troubleshooting"],
      ["Tools & Design", "Git, Kali Linux, Adobe Photoshop"],
    ],
  ),

  sectionHeading("Key Projects"),
  entryHeader("Phishing URL Detection System", "Flask, Machine Learning"),
  subLine("Final-Year Dissertation Project"),
  bullet("Developing a machine-learning system that detects phishing URLs in real time, applying AI to a practical security problem."),
  bullet("Covers data collection, feature engineering, model training and evaluation, and deployment through a Flask web interface."),

  entryHeader("Lamont Motors: Full-Stack Web Application", "Node.js, Express, React, PostgreSQL, Supabase"),
  bullet("Built and deployed a complete car-dealership platform end-to-end as the sole developer."),
  bullet("Independently handled backend setup, environment configuration, and database migration to Supabase/PostgreSQL."),

  entryHeader("SmartPay Solutions: Network Infrastructure Design", "Academic Project"),
  bullet("Designed a full network security architecture for a fintech scenario, covering VLAN segmentation, IPsec VPN, and WPA3 wireless security."),

  sectionHeading("Work Experience"),
  entryHeader("Mobile Money Vendor", "Kumasi, Ghana"),
  bullet("Managed daily transactions, customer service, and end-of-day cash reconciliation with accuracy under pressure."),
  bullet("Built reliability, attention to detail, and trust with a diverse customer base."),

  entryHeader("Teacher", ""),
  bullet("Planned and delivered structured lessons and managed classroom operations."),
  bullet("Strengthened communication, presentation, and organisational skills."),

  sectionHeading("Certifications"),
  entryHeader("Cybersecurity Internship Certificate", "Thrive Ghana"),

  sectionHeading("Core Strengths"),
  para([run("Quick learner  •  Problem-solving  •  Teamwork  •  Communication  •  Attention to detail  •  Adaptability in fast-paced environments")]),

  sectionHeading("Languages"),
  para([run("English", { bold: true }), run(" (Fluent)     "), run("Twi", { bold: true }), run(" (Native)")]),

  sectionHeading("References"),
  para([run("Available upon request.", { italics: true, color: MUTED })]),
];

const doc = new Document({
  creator: "Bright Akwasi Kodua",
  title: "Bright Akwasi Kodua - CV",
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "▪", alignment: AlignmentType.LEFT,
        style: {
          paragraph: { indent: { left: 300, hanging: 220 } },
          run: { color: ACCENT },
        },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: 16838 },
        margin: { top: 720, bottom: 720, left: MARGIN, right: MARGIN },
      },
    },
    children: [...header, ...body],
  }],
});

const out = path.join(__dirname, "Bright_Akwasi_Kodua_CV_LMI_Holdings.docx");
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log("Wrote " + out);
});
