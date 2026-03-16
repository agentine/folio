export interface XmlAttribute {
  name: string;
  value: string;
}

export interface SaxHandler {
  onOpenTag?(name: string, attributes: XmlAttribute[]): void;
  onCloseTag?(name: string): void;
  onText?(text: string): void;
  onCdata?(text: string): void;
}

const ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

function decodeEntities(text: string): string {
  return text.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#[0-9]+);/g, (_, ref: string) => {
    if (ref in ENTITIES) return ENTITIES[ref];
    if (ref.startsWith('#x')) return String.fromCodePoint(parseInt(ref.slice(2), 16));
    if (ref.startsWith('#')) return String.fromCodePoint(parseInt(ref.slice(1), 10));
    return _;
  });
}

function parseAttributes(attrString: string): XmlAttribute[] {
  const attrs: XmlAttribute[] = [];
  const re = /([a-zA-Z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(attrString)) !== null) {
    attrs.push({ name: m[1], value: decodeEntities(m[2] ?? m[3]) });
  }
  return attrs;
}

export function parseSax(xml: string | Buffer, handler: SaxHandler): void {
  const str = typeof xml === 'string' ? xml : xml.toString('utf8');
  let pos = 0;
  const len = str.length;

  while (pos < len) {
    const ltIdx = str.indexOf('<', pos);

    if (ltIdx === -1) {
      // Remaining text
      const text = str.slice(pos);
      if (text.trim() && handler.onText) {
        handler.onText(decodeEntities(text));
      }
      break;
    }

    // Text before tag
    if (ltIdx > pos) {
      const text = str.slice(pos, ltIdx);
      if (text.trim() && handler.onText) {
        handler.onText(decodeEntities(text));
      }
    }

    // XML declaration: <?xml ... ?>
    if (str.startsWith('<?', ltIdx)) {
      const end = str.indexOf('?>', ltIdx);
      if (end === -1) throw new Error('Unterminated processing instruction');
      pos = end + 2;
      continue;
    }

    // CDATA: <![CDATA[ ... ]]>
    if (str.startsWith('<![CDATA[', ltIdx)) {
      const end = str.indexOf(']]>', ltIdx + 9);
      if (end === -1) throw new Error('Unterminated CDATA section');
      if (handler.onCdata) {
        handler.onCdata(str.slice(ltIdx + 9, end));
      }
      pos = end + 3;
      continue;
    }

    // Comment: <!-- ... -->
    if (str.startsWith('<!--', ltIdx)) {
      const end = str.indexOf('-->', ltIdx + 4);
      if (end === -1) throw new Error('Unterminated comment');
      pos = end + 3;
      continue;
    }

    // DOCTYPE: <!DOCTYPE ... >
    if (str.startsWith('<!DOCTYPE', ltIdx) || str.startsWith('<!doctype', ltIdx)) {
      const end = str.indexOf('>', ltIdx);
      if (end === -1) throw new Error('Unterminated DOCTYPE');
      pos = end + 1;
      continue;
    }

    // Close tag: </name>
    if (str[ltIdx + 1] === '/') {
      const gtIdx = str.indexOf('>', ltIdx + 2);
      if (gtIdx === -1) throw new Error('Unterminated close tag');
      const name = str.slice(ltIdx + 2, gtIdx).trim();
      if (handler.onCloseTag) {
        handler.onCloseTag(name);
      }
      pos = gtIdx + 1;
      continue;
    }

    // Open tag or self-closing tag
    const gtIdx = str.indexOf('>', ltIdx + 1);
    if (gtIdx === -1) throw new Error('Unterminated open tag');

    const tagContent = str.slice(ltIdx + 1, gtIdx);
    const selfClosing = tagContent.endsWith('/');
    const content = selfClosing ? tagContent.slice(0, -1) : tagContent;

    // Split tag name from attributes
    const spaceIdx = content.search(/[\s/]/);
    let name: string;
    let attrStr: string;

    if (spaceIdx === -1) {
      name = content.trim();
      attrStr = '';
    } else {
      name = content.slice(0, spaceIdx);
      attrStr = content.slice(spaceIdx);
    }

    const attributes = parseAttributes(attrStr);

    if (handler.onOpenTag) {
      handler.onOpenTag(name, attributes);
    }
    if (selfClosing && handler.onCloseTag) {
      handler.onCloseTag(name);
    }

    pos = gtIdx + 1;
  }
}
