function escapeXmlText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeXmlAttr(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export class XmlElement {
  readonly name: string;
  private attributes: [string, string][] = [];
  private children: (XmlElement | string)[] = [];

  constructor(name: string) {
    this.name = name;
  }

  attr(name: string, value: string | number | boolean): this {
    this.attributes.push([name, String(value)]);
    return this;
  }

  child(element: XmlElement): this {
    this.children.push(element);
    return this;
  }

  text(content: string): this {
    this.children.push(content);
    return this;
  }

  toString(): string {
    const attrs = this.attributes
      .map(([k, v]) => ` ${k}="${escapeXmlAttr(v)}"`)
      .join('');

    if (this.children.length === 0) {
      return `<${this.name}${attrs}/>`;
    }

    const inner = this.children
      .map((c) => (typeof c === 'string' ? escapeXmlText(c) : c.toString()))
      .join('');

    return `<${this.name}${attrs}>${inner}</${this.name}>`;
  }
}

export function el(name: string): XmlElement {
  return new XmlElement(name);
}

export function xmlDeclaration(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
}
