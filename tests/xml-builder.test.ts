import { describe, it, expect } from 'vitest';
import { XmlElement, el, xmlDeclaration } from '../src/xml/builder.js';

describe('XML Builder', () => {
  it('builds a simple self-closing element', () => {
    const xml = el('br').toString();
    expect(xml).toBe('<br/>');
  });

  it('builds an element with text', () => {
    const xml = el('name').text('Alice').toString();
    expect(xml).toBe('<name>Alice</name>');
  });

  it('builds an element with attributes', () => {
    const xml = el('item').attr('id', '1').attr('type', 'test').toString();
    expect(xml).toBe('<item id="1" type="test"/>');
  });

  it('builds nested elements', () => {
    const xml = el('root')
      .child(el('a').text('1'))
      .child(el('b').text('2'))
      .toString();
    expect(xml).toBe('<root><a>1</a><b>2</b></root>');
  });

  it('escapes text content', () => {
    const xml = el('data').text('a < b & c > d').toString();
    expect(xml).toBe('<data>a &lt; b &amp; c &gt; d</data>');
  });

  it('escapes attribute values', () => {
    const xml = el('item').attr('val', 'a"b\'c&d<e>f').toString();
    expect(xml).toBe('<item val="a&quot;b&apos;c&amp;d&lt;e&gt;f"/>');
  });

  it('handles namespace attributes', () => {
    const xml = el('worksheet')
      .attr('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
      .child(el('row'))
      .toString();
    expect(xml).toBe('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><row/></worksheet>');
  });

  it('generates XML declaration', () => {
    expect(xmlDeclaration()).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  });

  it('builds deeply nested structure', () => {
    const xml = el('a')
      .child(
        el('b').child(
          el('c').child(
            el('d').text('deep'),
          ),
        ),
      )
      .toString();
    expect(xml).toBe('<a><b><c><d>deep</d></c></b></a>');
  });

  it('accepts numeric and boolean attribute values', () => {
    const xml = el('item').attr('count', 5).attr('active', true).toString();
    expect(xml).toBe('<item count="5" active="true"/>');
  });
});
