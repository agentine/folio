import type { RichTextRun, Font } from './model/types.js';
import { parseSax } from './xml/parser.js';
import type { XmlAttribute } from './xml/parser.js';
import { el, xmlDeclaration } from './xml/builder.js';

export interface SharedStringItem {
  text: string;
  richText?: RichTextRun[];
}

export class SharedStrings {
  private _items: SharedStringItem[] = [];
  private _index: Map<string, number> = new Map(); // dedup key -> index
  private _count = 0; // total references (increments on every add, even dupes)

  /** Add a plain string. Returns the shared string index. */
  add(text: string): number {
    this._count++;
    const existing = this._index.get(text);
    if (existing !== undefined) return existing;
    const idx = this._items.length;
    this._items.push({ text });
    this._index.set(text, idx);
    return idx;
  }

  /** Add a rich text item. Returns the shared string index. */
  addRichText(runs: RichTextRun[]): number {
    this._count++;
    const key = `\0rt:${JSON.stringify(runs)}`;
    const existing = this._index.get(key);
    if (existing !== undefined) return existing;
    const text = runs.map((r) => r.text).join('');
    const idx = this._items.length;
    this._items.push({ text, richText: runs });
    this._index.set(key, idx);
    return idx;
  }

  /** Get a shared string item by index. */
  get(index: number): SharedStringItem | undefined {
    return this._items[index];
  }

  /** Get plain text by index. */
  getText(index: number): string {
    return this._items[index]?.text ?? '';
  }

  /** Total number of unique strings. */
  get uniqueCount(): number {
    return this._items.length;
  }

  /** Total number of string references (including duplicates). */
  get count(): number {
    return this._count;
  }

  /** All shared string items. */
  get items(): ReadonlyArray<SharedStringItem> {
    return this._items;
  }

  /** Generate sharedStrings.xml content. */
  toXml(): string {
    const sst = el('sst')
      .attr('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
      .attr('count', this._count)
      .attr('uniqueCount', this._items.length);

    for (const item of this._items) {
      const si = el('si');
      if (item.richText) {
        for (const run of item.richText) {
          const r = el('r');
          if (run.font) {
            r.child(fontToRPr(run.font));
          }
          r.child(el('t').text(run.text));
          si.child(r);
        }
      } else {
        const t = el('t');
        // Preserve leading/trailing whitespace
        if (item.text !== item.text.trim()) {
          t.attr('xml:space', 'preserve');
        }
        t.text(item.text);
        si.child(t);
      }
      sst.child(si);
    }

    return xmlDeclaration() + sst.toString();
  }

  /** Parse sharedStrings.xml content into this table. */
  static fromXml(xml: string): SharedStrings {
    const sst = new SharedStrings();
    let currentRuns: RichTextRun[] | null = null;
    let currentFont: Font | null = null;
    let inT = false;
    let inR = false;
    let textBuf = '';
    let totalCount = 0;

    // Track element nesting to know context
    const stack: string[] = [];

    parseSax(xml, {
      onOpenTag(name: string, attrs: XmlAttribute[]) {
        stack.push(name);
        if (name === 'sst') {
          for (const a of attrs) {
            if (a.name === 'count') totalCount = parseInt(a.value, 10);
          }
        } else if (name === 'si') {
          currentRuns = null;
          currentFont = null;
          textBuf = '';
        } else if (name === 'r') {
          inR = true;
          if (!currentRuns) currentRuns = [];
          currentFont = null;
        } else if (name === 'rPr') {
          currentFont = {};
        } else if (name === 't') {
          inT = true;
          textBuf = '';
        } else if (currentFont) {
          applyFontElement(currentFont, name, attrs);
        }
      },
      onText(text: string) {
        if (inT) {
          textBuf += text;
        }
      },
      onCloseTag(name: string) {
        stack.pop();
        if (name === 't') {
          inT = false;
          if (inR && currentRuns) {
            currentRuns.push({ text: textBuf, ...(currentFont ? { font: currentFont } : {}) });
            currentFont = null;
          }
        } else if (name === 'r') {
          inR = false;
        } else if (name === 'si') {
          if (currentRuns && currentRuns.length > 0) {
            // Rich text: add directly to items (bypass dedup key to preserve original order)
            const text = currentRuns.map((r) => r.text).join('');
            sst._items.push({ text, richText: currentRuns });
          } else {
            // Plain text
            sst._items.push({ text: textBuf });
          }
        }
      },
    });

    // Build index for lookup
    for (let i = 0; i < sst._items.length; i++) {
      const item = sst._items[i];
      if (item.richText) {
        sst._index.set(`\0rt:${JSON.stringify(item.richText)}`, i);
      } else {
        sst._index.set(item.text, i);
      }
    }
    sst._count = totalCount || sst._items.length;

    return sst;
  }
}

function fontToRPr(font: Font): ReturnType<typeof el> {
  const rPr = el('rPr');
  if (font.bold) rPr.child(el('b'));
  if (font.italic) rPr.child(el('i'));
  if (font.strike) rPr.child(el('strike'));
  if (font.underline) {
    const u = el('u');
    if (typeof font.underline === 'string') u.attr('val', font.underline);
    rPr.child(u);
  }
  if (font.vertAlign) rPr.child(el('vertAlign').attr('val', font.vertAlign));
  if (font.size) rPr.child(el('sz').attr('val', font.size));
  if (font.color) {
    const c = el('color');
    if (font.color.argb) c.attr('rgb', font.color.argb);
    if (font.color.theme !== undefined) c.attr('theme', font.color.theme);
    if (font.color.tint !== undefined) c.attr('tint', font.color.tint);
    rPr.child(c);
  }
  if (font.name) rPr.child(el('rFont').attr('val', font.name));
  if (font.family !== undefined) rPr.child(el('family').attr('val', font.family));
  return rPr;
}

function applyFontElement(font: Font, name: string, attrs: XmlAttribute[]): void {
  const val = attrs.find((a) => a.name === 'val')?.value;
  switch (name) {
    case 'b':
      font.bold = val !== '0' && val !== 'false';
      break;
    case 'i':
      font.italic = val !== '0' && val !== 'false';
      break;
    case 'strike':
      font.strike = val !== '0' && val !== 'false';
      break;
    case 'u':
      font.underline = (val as Font['underline']) || true;
      break;
    case 'vertAlign':
      if (val === 'superscript' || val === 'subscript') font.vertAlign = val;
      break;
    case 'sz':
      if (val) font.size = parseFloat(val);
      break;
    case 'color': {
      const rgb = attrs.find((a) => a.name === 'rgb')?.value;
      const theme = attrs.find((a) => a.name === 'theme')?.value;
      const tint = attrs.find((a) => a.name === 'tint')?.value;
      font.color = {};
      if (rgb) font.color.argb = rgb;
      if (theme) font.color.theme = parseInt(theme, 10);
      if (tint) font.color.tint = parseFloat(tint);
      break;
    }
    case 'rFont':
      if (val) font.name = val;
      break;
    case 'family':
      if (val) font.family = parseInt(val, 10);
      break;
  }
}
