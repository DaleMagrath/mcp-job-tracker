/**
 * Render a resume .docx from a merged spec (resume_master.json facts + a
 * posting's tailored summary/key-qualifications), matching the 2-page,
 * Liberation Serif format the prior python-docx script produced.
 *
 * Ported from the former scripts/gen_resume.py. One layout note carried over
 * verbatim: role/date heading lines are a 2-column, borderless TABLE (title
 * left, dates right) rather than a tab stop across mixed bold/italic runs —
 * the table is what survived clean text-extraction in the original; there's
 * no reason to believe a tab stop behaves better here, so the table stays.
 */

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  ExternalHyperlink,
  BorderStyle,
  WidthType,
  AlignmentType,
  VerticalAlign,
} from "docx";

const HEADING_COLOR = "2E4057";
const CONTACT_COLOR = "444444";
const DATES_COLOR = "555555";
const LINK_COLOR = "1155CC";
const FONT = "Liberation Serif";

// Content width = page (12240 twips) - left (900) - right (900) = 10440 twips.
const CONTENT_TWIPS = 10440;
const LEFT_COL_TWIPS = Math.round(CONTENT_TWIPS * 0.7);
const RIGHT_COL_TWIPS = CONTENT_TWIPS - LEFT_COL_TWIPS;

const SINGLE_LINE = { line: 240, lineRule: "auto" as const };
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const ZERO_MARGINS = { top: 0, bottom: 0, left: 0, right: 0 };

export interface ResumeExperience {
  org?: string;
  title?: string;
  dates?: string;
  description?: string;
}
export interface ResumeProject {
  name?: string;
  dates?: string;
  description?: string;
}
export interface ResumeKeyQualifications {
  target?: string;
  bullets?: string[];
}
export interface ResumeSpec {
  name?: string;
  contact?: string;
  summary?: string;
  key_qualifications?: ResumeKeyQualifications;
  experience?: ResumeExperience[];
  projects?: ResumeProject[];
  education?: string[];
  certifications?: string[];
  skills?: string[];
  honors?: string[];
}

function heading(text: string): Paragraph {
  return new Paragraph({
    spacing: { before: 240, after: 100, ...SINGLE_LINE },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 2 } },
    children: [
      new TextRun({ text: text.toUpperCase(), bold: true, size: 22, color: HEADING_COLOR, font: FONT }),
    ],
  });
}

function role(title: string, dates: string): Table {
  return new Table({
    width: { size: CONTENT_TWIPS, type: WidthType.DXA },
    columnWidths: [LEFT_COL_TWIPS, RIGHT_COL_TWIPS],
    borders: {
      top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER,
      insideHorizontal: NO_BORDER, insideVertical: NO_BORDER,
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: LEFT_COL_TWIPS, type: WidthType.DXA },
            margins: ZERO_MARGINS,
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                spacing: { before: 160, after: 40, ...SINGLE_LINE },
                children: [new TextRun({ text: title, bold: true, size: 21, font: FONT })],
              }),
            ],
          }),
          new TableCell({
            width: { size: RIGHT_COL_TWIPS, type: WidthType.DXA },
            margins: ZERO_MARGINS,
            verticalAlign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                spacing: { before: 160, after: 40, ...SINGLE_LINE },
                children: [
                  new TextRun({ text: dates, italics: true, size: 20, color: DATES_COLOR, font: FONT }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function body(text: string): Paragraph {
  return new Paragraph({
    spacing: { after: 160, ...SINGLE_LINE },
    children: [new TextRun({ text, size: 21, font: FONT })],
  });
}

function bullet(text: string): Paragraph {
  return new Paragraph({
    spacing: { after: 100, ...SINGLE_LINE },
    indent: { left: 230, hanging: 230 },
    children: [new TextRun({ text: "• " + text, size: 20, font: FONT })],
  });
}

/** The header's contact line: "•"-separated tokens, with a linkedin.com
 *  token rendered as a real hyperlink — same rule gen_resume.py used. */
function contactLine(contact: string): Paragraph {
  const parts = contact ? contact.split("•").map((t) => t.trim()).filter(Boolean) : [];
  const children: (TextRun | ExternalHyperlink)[] = [];
  parts.forEach((tok, i) => {
    if (i > 0) {
      children.push(new TextRun({ text: "  •  ", size: 20, color: CONTACT_COLOR, font: FONT }));
    }
    if (tok.toLowerCase().includes("linkedin.com")) {
      const url = /^https?:\/\//i.test(tok) ? tok : "https://" + tok;
      children.push(
        new ExternalHyperlink({
          link: url,
          children: [
            new TextRun({ text: tok, size: 20, color: LINK_COLOR, underline: {}, font: FONT }),
          ],
        })
      );
    } else {
      children.push(new TextRun({ text: tok, size: 20, color: CONTACT_COLOR, font: FONT }));
    }
  });
  return new Paragraph({
    spacing: { after: 120, ...SINGLE_LINE },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "999999", space: 6 } },
    children,
  });
}

/** Build the resume document and return it as a .docx Buffer. */
export async function buildResumeDocx(spec: ResumeSpec): Promise<Buffer> {
  const children: (Paragraph | Table)[] = [];

  // ---- Header ----
  children.push(
    new Paragraph({
      spacing: { after: 0, ...SINGLE_LINE },
      children: [new TextRun({ text: spec.name ?? "", bold: true, size: 40, font: FONT })],
    })
  );
  children.push(contactLine(spec.contact ?? ""));

  // ---- Summary ----
  if (spec.summary) {
    children.push(heading("Summary"));
    children.push(body(spec.summary));
  }

  // ---- Key qualifications ----
  const bullets = spec.key_qualifications?.bullets ?? [];
  if (bullets.length) {
    const target = (spec.key_qualifications?.target ?? "").trim();
    const label = target ? `Key Qualifications — ${target} Alignment` : "Key Qualifications";
    children.push(heading(label));
    for (const b of bullets) children.push(bullet(b));
  }

  // ---- Experience / Projects ----
  const sections: [string, (ResumeExperience | ResumeProject)[], boolean][] = [
    ["Experience", spec.experience ?? [], true],
    ["Projects", spec.projects ?? [], false],
  ];
  for (const [section, items, isExp] of sections) {
    if (!items.length) continue;
    children.push(heading(section));
    for (const e of items) {
      let head: string;
      if (isExp) {
        const exp = e as ResumeExperience;
        const org = (exp.org ?? "").trim();
        const title = (exp.title ?? "").trim();
        head = org && title ? `${org}, ${title}` : org || title;
      } else {
        head = (e as ResumeProject).name ?? "";
      }
      children.push(role(head, e.dates ?? ""));
      if (e.description) children.push(body(e.description));
    }
  }

  // ---- Education ----
  if (spec.education?.length) {
    children.push(heading("Education"));
    for (const line of spec.education) children.push(body(line));
  }

  // ---- Certifications ----
  if (spec.certifications?.length) {
    children.push(heading("Licenses & Certifications"));
    for (const line of spec.certifications) children.push(body(line));
  }

  // ---- Skills ----
  if (spec.skills?.length) {
    children.push(heading("Skills"));
    children.push(body(spec.skills.join(" • ")));
  }

  // ---- Honors & Awards (optional, off by default for 2-page fit) ----
  if (spec.honors?.length) {
    children.push(heading("Honors & Awards"));
    for (const line of spec.honors) children.push(body(line));
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: FONT, size: 21 },
          paragraph: { spacing: { after: 0, ...SINGLE_LINE } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 720, bottom: 720, left: 900, right: 900 },
          },
        },
        children,
      },
    ],
  });

  return Packer.toBuffer(doc);
}
