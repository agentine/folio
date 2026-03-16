import { describe, it, expect } from 'vitest';
import { parseSax } from '../src/xml/parser.js';
import type { XmlAttribute } from '../src/xml/parser.js';

describe('SAX XML Parser', () => {
  it('parses a simple element with text', () => {
    const tags: string[] = [];
    const texts: string[] = [];

    parseSax('<root>hello</root>', {
      onOpenTag: (name) => tags.push(`open:${name}`),
      onCloseTag: (name) => tags.push(`close:${name}`),
      onText: (text) => texts.push(text),
    });

    expect(tags).toEqual(['open:root', 'close:root']);
    expect(texts).toEqual(['hello']);
  });

  it('parses attributes', () => {
    let attrs: XmlAttribute[] = [];

    parseSax('<item id="42" name="test" flag=\'true\'/>', {
      onOpenTag: (_name, a) => { attrs = a; },
    });

    expect(attrs).toEqual([
      { name: 'id', value: '42' },
      { name: 'name', value: 'test' },
      { name: 'flag', value: 'true' },
    ]);
  });

  it('handles self-closing tags', () => {
    const events: string[] = [];

    parseSax('<root><br/><hr /></root>', {
      onOpenTag: (name) => events.push(`open:${name}`),
      onCloseTag: (name) => events.push(`close:${name}`),
    });

    expect(events).toEqual([
      'open:root',
      'open:br', 'close:br',
      'open:hr', 'close:hr',
      'close:root',
    ]);
  });

  it('handles XML namespaces', () => {
    const xml = '<spreadsheetml:worksheet xmlns:spreadsheetml="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><spreadsheetml:row/></spreadsheetml:worksheet>';
    const tags: string[] = [];

    parseSax(xml, {
      onOpenTag: (name) => tags.push(name),
      onCloseTag: (name) => tags.push(`/${name}`),
    });

    expect(tags).toEqual([
      'spreadsheetml:worksheet',
      'spreadsheetml:row',
      '/spreadsheetml:row',
      '/spreadsheetml:worksheet',
    ]);
  });

  it('decodes XML entities in text', () => {
    const texts: string[] = [];

    parseSax('<root>&lt;a&gt; &amp; &quot;b&quot; &apos;c&apos;</root>', {
      onText: (text) => texts.push(text),
    });

    expect(texts).toEqual(['<a> & "b" \'c\'']);
  });

  it('decodes XML entities in attributes', () => {
    let attrs: XmlAttribute[] = [];

    parseSax('<item val="a&amp;b&lt;c"/>', {
      onOpenTag: (_, a) => { attrs = a; },
    });

    expect(attrs[0].value).toBe('a&b<c');
  });

  it('handles CDATA sections', () => {
    const cdata: string[] = [];

    parseSax('<root><![CDATA[<not & xml>]]></root>', {
      onCdata: (text) => cdata.push(text),
    });

    expect(cdata).toEqual(['<not & xml>']);
  });

  it('handles nested elements', () => {
    const events: string[] = [];

    parseSax('<a><b><c>text</c></b></a>', {
      onOpenTag: (name) => events.push(`+${name}`),
      onCloseTag: (name) => events.push(`-${name}`),
      onText: (text) => events.push(`"${text}"`),
    });

    expect(events).toEqual(['+a', '+b', '+c', '"text"', '-c', '-b', '-a']);
  });

  it('skips XML declaration', () => {
    const tags: string[] = [];

    parseSax('<?xml version="1.0" encoding="UTF-8"?><root/>', {
      onOpenTag: (name) => tags.push(name),
    });

    expect(tags).toEqual(['root']);
  });

  it('skips comments', () => {
    const tags: string[] = [];

    parseSax('<root><!-- comment --><item/></root>', {
      onOpenTag: (name) => tags.push(name),
    });

    expect(tags).toEqual(['root', 'item']);
  });

  it('handles numeric character references', () => {
    const texts: string[] = [];

    parseSax('<root>&#65;&#x42;</root>', {
      onText: (text) => texts.push(text),
    });

    expect(texts).toEqual(['AB']);
  });
});
