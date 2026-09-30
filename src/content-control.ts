import type { Document, Element, Node } from '@xmldom/xmldom';
import type { ContentControlInfo, ContentControlKind } from './types.js';
import type { PartAccess } from './internal/context.js';
import { assertIndex } from './operations.js';
import { WORD_NS, assertText, children, childrenThroughTransparent, descendants, isValidXmlCharCode, setWordValue, wordElement, wordValue } from './xml.js';
import { appendText, bodyOf, compactDefined, mainParagraphElements } from './internal/elements.js';
import { visibleTextOf } from './revisions.js';

const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';

export interface ContentControlContext extends Pick<PartAccess, 'mainPath'> {
  getCachedPartDocument(path: string): Document;
  updatePartXmlInternal(path: string, update: (document: Document) => boolean | void): void;
}

export function textElements(element: Element): Element[] {
  const result: Element[] = [];
  function walk(node: Node, deletedDepth = 0): void {
    for (let child = node.firstChild; child; child = child.nextSibling) {
      if (child.nodeType !== 1) continue;
      const element = child as Element;
      if (element.namespaceURI === WORD_NS) {
        const localName = element.localName ?? '';
        if (element.localName === 'p') continue;
        const inDeleted = deletedDepth > 0 || localName === 'delText';
        if (!inDeleted && ['t', 'tab', 'br', 'cr', 'noBreakHyphen', 'softHyphen', 'sym'].includes(localName)) {
          result.push(element);
          continue;
        }
      }
      walk(element, deletedDepth + (element.namespaceURI === WORD_NS && ['del', 'moveFrom'].includes(element.localName ?? '') ? 1 : 0));
    }
  }
  walk(element);
  return result;
}

export function elementText(element: Element): string {
  if (['t', 'delText'].includes(element.localName ?? '')) return element.textContent ?? '';
  if (element.localName === 'tab') return '\t';
  if (element.localName === 'noBreakHyphen') return '\u2011';
  if (element.localName === 'softHyphen') return '\u00ad';
  if (element.localName === 'sym') {
    const value = element.getAttributeNS(WORD_NS, 'char') ?? element.getAttribute('w:char');
    if (!value || !/^[a-f0-9]{1,4}$/i.test(value)) return '';
    const code = Number.parseInt(value, 16);
    return Number.isFinite(code) && isValidXmlCharCode(code) ? String.fromCharCode(code) : '�';
  }
  return '\n';
}

export function textOf(element: Element): string {
  if (element.localName === 'r' && descendants(element, 'instrText').length > 0) return '';
  const text = visibleTextOf(element) || textElements(element).map(elementText).join('');
  return text;
}

// Edit only text-bearing nodes. Drawings, bookmarks, field codes and other XML survive.
export function replaceSpan(paragraph: Element, start: number, end: number, replacement: string): void {
  const elements = textElements(paragraph);
  if (elements.length === 0) {
    const run = children(paragraph, 'r')[0] ?? wordElement(paragraph.ownerDocument!, 'r');
    if (!run.parentNode) paragraph.appendChild(run);
    appendText(run, replacement);
    return;
  }
  let offset = 0;
  let inserted = false;
  for (const element of elements) {
    const oldText = elementText(element);
    const next = offset + oldText.length;
    const insertHere = !inserted && start >= offset && (start === end ? start <= next : start < next);
    const overlaps = offset < end && next > start;
    if (insertHere || overlaps) {
      if (insertHere && !overlaps && element.localName !== 't') {
        appendText(element.parentNode as Element, replacement, start === offset ? element : element.nextSibling);
        inserted = true;
        offset = next;
        continue;
      }
      const prefix = oldText.slice(0, Math.max(0, Math.min(oldText.length, start - offset)));
      const suffix = oldText.slice(Math.max(0, Math.min(oldText.length, end - offset)));
      const text = prefix + (insertHere ? replacement : '') + suffix;
      const parent = element.parentNode as Element;
      appendText(parent, text, element);
      parent.removeChild(element);
      if (insertHere) inserted = true;
    }
    offset = next;
  }
}

function contentControlText(content: Element): string {
  const paragraphs = descendants(content, 'p');
  return paragraphs.length ? paragraphs.map(textOf).join('\n') : textOf(content);
}

const SDT_PROPERTY_ORDER = [
  'alias', 'lock', 'placeholder', 'showingPlcHdr', 'dataBinding', 'temporary', 'id', 'tag',
  'docPartObj', 'docPartGallery', 'docPartUnique', 'comboBox', 'date', 'docPartList', 'dropDownList',
  'picture', 'richText', 'text', 'citation', 'group', 'bibliography', 'equation', 'ocx', 'entityPicker',
];

function sdtProperty(parent: Element, name: string): Element {
  const existing = children(parent, name)[0];
  if (existing) return existing;
  const element = wordElement(parent.ownerDocument!, name);
  const position = SDT_PROPERTY_ORDER.indexOf(name);
  const following = children(parent).find((child) => {
    const childPosition = SDT_PROPERTY_ORDER.indexOf(child.localName ?? '');
    return position !== -1 && childPosition > position;
  });
  parent.insertBefore(element, following ?? null);
  return element;
}

function sdtProperties(control: Element): Element {
  let properties = children(control, 'sdtPr')[0];
  if (!properties) {
    properties = wordElement(control.ownerDocument!, 'sdtPr');
    control.insertBefore(properties, children(control, 'sdtContent')[0] ?? control.firstChild);
  }
  return properties;
}

function contentControlKind(control: Element): ContentControlKind {
  const properties = children(control, 'sdtPr')[0] ?? control;
  if (children(properties, 'checkbox', W14_NS).length) return 'checkbox';
  const types: Array<[string, ContentControlKind]> = [
    ['text', 'text'], ['richText', 'richText'], ['dropDownList', 'dropDownList'],
    ['comboBox', 'comboBox'], ['date', 'date'], ['picture', 'picture'], ['group', 'group'],
  ];
  return types.find(([name]) => children(properties, name).length > 0)?.[1] ?? 'unknown';
}

function textSpanWithin(root: Element, scope: Element): { start: number; end: number; text: string } | undefined {
  const elements = textElements(root);
  const selected = elements.filter((element) => {
    let ancestor: Node | null = element;
    while (ancestor && ancestor !== scope) ancestor = ancestor.parentNode;
    return ancestor === scope;
  });
  if (!selected.length) return undefined;
  const firstIndex = elements.indexOf(selected[0]!);
  const start = elements.slice(0, firstIndex).reduce((length, element) => length + elementText(element).length, 0);
  const text = selected.map(elementText).join('');
  return { start, end: start + text.length, text };
}

function appendContentControlText(parent: Element, text: string, withinParagraph: boolean): void {
  const runs = descendants(parent, 'r').filter((run) => {
    let ancestor: Element | null = run.parentNode as Element | null;
    while (ancestor && ancestor !== parent) {
      if (ancestor.namespaceURI === WORD_NS && ['del', 'moveFrom'].includes(ancestor.localName ?? '')) return false;
      ancestor = ancestor.parentNode as Element | null;
    }
    return ancestor === parent;
  });
  const run = runs[0] ?? wordElement(parent.ownerDocument!, 'r');
  if (!run.parentNode) {
    if (withinParagraph) {
      const paragraphProperties = children(parent, 'pPr')[0];
      if (paragraphProperties) parent.insertBefore(run, paragraphProperties.nextSibling);
      else parent.insertBefore(run, parent.firstChild);
    } else {
      parent.appendChild(run);
    }
  }
  appendText(run, text);
}

function setSdtText(control: Element, text: string): void {
  const content = children(control, 'sdtContent')[0];
  if (!content) throw new Error('Content control has no w:sdtContent.');
  const paragraphs = descendants(content, 'p');
  if (!paragraphs.length) {
    let ancestor = control.parentNode as Element | null;
    while (ancestor && ancestor.localName !== 'p' && ancestor.localName !== 'body' && ancestor.localName !== 'tc') {
      ancestor = ancestor.parentNode as Element | null;
    }
    if (ancestor?.localName === 'p') {
      const span = textSpanWithin(ancestor, content);
      if (span) {
        const old = span.text;
        let prefix = 0;
        while (prefix < old.length && prefix < text.length && old[prefix] === text[prefix]) prefix++;
        let oldEnd = old.length;
        let newEnd = text.length;
        while (oldEnd > prefix && newEnd > prefix && old[oldEnd - 1] === text[newEnd - 1]) {
          oldEnd--;
          newEnd--;
        }
        if (prefix > 0 && /[\ud800-\udbff]/.test(old[prefix - 1]!)) prefix--;
        if (oldEnd < old.length && /[\udc00-\udfff]/.test(old[oldEnd]!)) { oldEnd++; newEnd++; }
        replaceSpan(ancestor, span.start + prefix, span.start + oldEnd, text.slice(prefix, newEnd));
      } else {
        appendContentControlText(content, text, false);
      }
      return;
    }
    const paragraph = wordElement(control.ownerDocument!, 'p');
    content.appendChild(paragraph);
    appendContentControlText(paragraph, text, true);
    return;
  }
  if (paragraphs.length === 1) {
    const paragraph = paragraphs[0]!;
    const span = textSpanWithin(paragraph, content);
    if (span) {
      let prefix = 0;
      while (prefix < span.text.length && prefix < text.length && span.text[prefix] === text[prefix]) prefix++;
      let oldEnd = span.text.length;
      let newEnd = text.length;
      while (oldEnd > prefix && newEnd > prefix && span.text[oldEnd - 1] === text[newEnd - 1]) {
        oldEnd--;
        newEnd--;
      }
      if (prefix > 0 && /[\ud800-\udbff]/.test(span.text[prefix - 1]!)) prefix--;
      if (oldEnd < span.text.length && /[\udc00-\udfff]/.test(span.text[oldEnd]!)) { oldEnd++; newEnd++; }
      replaceSpan(paragraph, span.start + prefix, span.start + oldEnd, text.slice(prefix, newEnd));
    } else {
      appendContentControlText(paragraph, text, true);
    }
    return;
  }
  for (const paragraph of paragraphs) {
    const span = textSpanWithin(paragraph, content);
    if (span) replaceSpan(paragraph, span.start, span.end, '');
  }
  appendContentControlText(paragraphs[0]!, text, true);
}

function ensureW14Namespace(control: Element): void {
  const root = control.ownerDocument?.documentElement;
  if (root && root.lookupNamespaceURI('w14') !== W14_NS) {
    root.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:w14', W14_NS);
  }
}

export function getContentControls(ctx: ContentControlContext): ContentControlInfo[] {
    const body = bodyOf(ctx.getCachedPartDocument(ctx.mainPath));
    const paragraphs = mainParagraphElements(body);
    const paragraphIndexes = new Map(paragraphs.map((paragraph, index) => [paragraph, index]));
    return descendants(body, 'sdt').map((control) => {
      const properties = children(control, 'sdtPr')[0];
      const controlContent = children(control, 'sdtContent')[0];
      const idValue = wordValue(children(properties ?? control, 'id')[0]);
      const parsedId = idValue !== undefined && /^\d+$/.test(idValue) ? Number(idValue) : undefined;
      const id = parsedId !== undefined && Number.isSafeInteger(parsedId) && parsedId >= 0 ? parsedId : undefined;
      const lockValue = wordValue(children(properties ?? control, 'lock')[0]);
      const lock: ContentControlInfo['lock'] = lockValue === 'sdtLocked' || lockValue === 'contentLocked' ||
        lockValue === 'sdtContentLocked' ? lockValue : 'unlocked';
      const checkbox = children(properties ?? control, 'checkbox', W14_NS)[0];
      const kind = contentControlKind(control);
      const checkboxChecked = checkbox && children(checkbox, 'checked', W14_NS)[0];
      const checkedValue = checkboxChecked?.getAttributeNS(W14_NS, 'val') ?? undefined;
      const list = children(properties ?? control, kind === 'dropDownList' ? 'dropDownList' : 'comboBox')[0];
      const items = list ? children(list, 'listItem').map((item) => ({
        displayText: item.getAttributeNS(WORD_NS, 'displayText') ?? '',
        value: item.getAttributeNS(WORD_NS, 'value') ?? '',
      })) : undefined;
      const date = children(properties ?? control, 'date')[0];
      const dateFormat = wordValue(children(date ?? control, 'dateFormat')[0]);
      const placeholder = children(properties ?? control, 'placeholder')[0];
      const binding = children(properties ?? control, 'dataBinding')[0];
      const bindingValue = binding ? compactDefined({
        prefixMappings: binding.getAttributeNS(WORD_NS, 'prefixMappings') ?? undefined,
        xpath: binding.getAttributeNS(WORD_NS, 'xpath') ?? undefined,
        storeItemId: binding.getAttributeNS(WORD_NS, 'storeItemID') ?? undefined,
      }) : undefined;
      const showingPlaceholder = children(properties ?? control, 'showingPlcHdr')[0];
      const placeholderValue = wordValue(showingPlaceholder);
      const controlParagraphs = controlContent ? descendants(controlContent, 'p') : [];
      if (controlContent && controlParagraphs.length === 0) {
        let ancestor = control.parentNode as Element | null;
        while (ancestor && ancestor !== body) {
          if (ancestor.namespaceURI === WORD_NS && ancestor.localName === 'p') {
            controlParagraphs.push(ancestor);
            break;
          }
          ancestor = ancestor.parentNode as Element | null;
        }
      }
      let ancestor = control.parentNode as Element | null;
      let nested = false;
      while (ancestor && ancestor !== body) {
        if (ancestor.namespaceURI === WORD_NS && ancestor.localName === 'sdt') {
          nested = true;
          break;
        }
        ancestor = ancestor.parentNode as Element | null;
      }
      return compactDefined({
        id,
        kind,
        alias: wordValue(children(properties ?? control, 'alias')[0]),
        tag: wordValue(children(properties ?? control, 'tag')[0]),
        lock,
        showingPlaceholder: !!showingPlaceholder && !['0', 'false', 'off'].includes((placeholderValue ?? '1').toLowerCase()),
        placeholderDocPart: wordValue(children(placeholder ?? control, 'docPart')[0]),
        items,
        checked: checkedValue === undefined || !['1', 'true', 'on', '0', 'false', 'off'].includes(checkedValue.toLowerCase())
          ? undefined : ['1', 'true', 'on'].includes(checkedValue.toLowerCase()),
        dateFormat,
        dataBinding: bindingValue,
        paragraphs: [...new Set(controlParagraphs.map((paragraph) => paragraphIndexes.get(paragraph)).filter((index): index is number => index !== undefined))],
        nested,
        text: controlContent ? contentControlText(controlContent) : '',
      });
    });
}

export function setContentControlText(ctx: ContentControlContext, id: number, text: string): void {
    assertIndex(id);
    assertText(text);
    ctx.updatePartXmlInternal(ctx.mainPath, (document) => {
      const body = bodyOf(document);
      const control = descendants(body, 'sdt').find((candidate) => {
        const properties = children(candidate, 'sdtPr')[0];
        const value = wordValue(children(properties ?? candidate, 'id')[0]);
        return value !== undefined && Number(value) === id;
      });
      if (!control) throw new Error(`Content control ${id} does not exist.`);
      const kind = contentControlKind(control);
      if (!['text', 'richText', 'dropDownList', 'comboBox', 'date'].includes(kind)) {
        throw new Error(`Content control ${id} does not accept text values.`);
      }
      const properties = children(control, 'sdtPr')[0];
      const lock = wordValue(children(properties ?? control, 'lock')[0]);
      if (lock === 'contentLocked' || lock === 'sdtContentLocked') {
        throw new Error(`Content control ${id} is locked against content changes.`);
      }
      const content = children(control, 'sdtContent')[0];
      if (!content) throw new Error('Content control has no w:sdtContent.');
      const showingPlaceholder = children(properties ?? control, 'showingPlcHdr');
      if (contentControlText(content) === text && showingPlaceholder.length === 0) return false;
      setSdtText(control, text);
      for (const showing of showingPlaceholder) properties?.removeChild(showing);
      return true;
    });
}

export function setContentControlChecked(ctx: ContentControlContext, id: number, checked: boolean): void {
    assertIndex(id);
    if (typeof checked !== 'boolean') throw new Error('checked must be boolean.');
    ctx.updatePartXmlInternal(ctx.mainPath, (document) => {
      const control = descendants(bodyOf(document), 'sdt').find((candidate) => {
        const properties = children(candidate, 'sdtPr')[0];
        const value = wordValue(children(properties ?? candidate, 'id')[0]);
        return value !== undefined && Number(value) === id;
      });
      if (!control) throw new Error(`Content control ${id} does not exist.`);
      if (contentControlKind(control) !== 'checkbox') {
        throw new Error(`Content control ${id} is not a checkbox.`);
      }
      ensureW14Namespace(control);
      const properties = sdtProperties(control);
      let checkbox = children(properties, 'checkbox', W14_NS)[0];
      if (!checkbox) {
        checkbox = document.createElementNS(W14_NS, 'w14:checkbox');
        properties.appendChild(checkbox);
      }
      const checkedElement = children(checkbox, 'checked', W14_NS)[0] ??
        (checkbox.appendChild(document.createElementNS(W14_NS, 'w14:checked')) as Element);
      const nextValue = checked ? '1' : '0';
      if (checkedElement.getAttributeNS(W14_NS, 'val') === nextValue) return false;
      checkedElement.setAttributeNS(W14_NS, 'w14:val', nextValue);
      return true;
    });
}

export function setContentControlProperties(
  ctx: ContentControlContext,
  id: number,
  patch: { alias?: string | null; tag?: string | null; lock?: ContentControlInfo['lock'] },
): void {
    assertIndex(id);
    if (!patch || typeof patch !== 'object' || Array.isArray(patch) ||
        Object.keys(patch).some((key) => !['alias', 'tag', 'lock'].includes(key))) {
      throw new Error('Content control properties patch must contain only alias, tag, or lock.');
    }
    if ('alias' in patch && patch.alias !== null && patch.alias !== undefined) assertText(patch.alias, 'alias');
    if ('tag' in patch && patch.tag !== null && patch.tag !== undefined) assertText(patch.tag, 'tag');
    if ('lock' in patch && patch.lock !== undefined &&
        !['sdtLocked', 'contentLocked', 'sdtContentLocked', 'unlocked'].includes(patch.lock)) {
      throw new Error('Invalid content control lock.');
    }
    ctx.updatePartXmlInternal(ctx.mainPath, (document) => {
      const control = descendants(bodyOf(document), 'sdt').find((candidate) => {
        const properties = children(candidate, 'sdtPr')[0];
        const value = wordValue(children(properties ?? candidate, 'id')[0]);
        return value !== undefined && Number(value) === id;
      });
      if (!control) throw new Error(`Content control ${id} does not exist.`);
      let properties = children(control, 'sdtPr')[0];
      const updateValue = (name: 'alias' | 'tag', value: string | null | undefined): boolean => {
        if (!(name in patch)) return false;
        const entries = children(properties ?? control, name);
        if (value === null) {
          for (const entry of entries) entry.parentNode?.removeChild(entry);
          return entries.length > 0;
        }
        if (value === undefined) return false;
        if (entries.length && wordValue(entries[0]) === value && entries.length === 1) return false;
        properties ??= sdtProperties(control);
        const propertyElement = entries[0] ?? sdtProperty(properties, name);
        setWordValue(propertyElement, value);
        for (const duplicate of entries.slice(1)) duplicate.parentNode?.removeChild(duplicate);
        return true;
      };
      let changed = updateValue('alias', patch.alias);
      changed = updateValue('tag', patch.tag) || changed;
      if ('lock' in patch && patch.lock !== undefined) {
        const locks = children(properties ?? control, 'lock');
        if (patch.lock === 'unlocked') {
          for (const entry of locks) entry.parentNode?.removeChild(entry);
          changed = locks.length > 0 || changed;
        } else if (locks.length === 1 && wordValue(locks[0]) === patch.lock) {
          return changed;
        } else {
          properties ??= sdtProperties(control);
          const lock = locks[0] ?? sdtProperty(properties, 'lock');
          setWordValue(lock, patch.lock);
          for (const duplicate of locks.slice(1)) duplicate.parentNode?.removeChild(duplicate);
          changed = true;
        }
      }
      return changed;
    });
}

export function removeContentControl(ctx: ContentControlContext, id: number, options: { keepContent?: boolean } = {}): void {
    assertIndex(id);
    if (!options || typeof options !== 'object' || Array.isArray(options) ||
        Object.keys(options).some((key) => key !== 'keepContent') ||
        ('keepContent' in options && typeof options.keepContent !== 'boolean')) {
      throw new Error('Invalid content control removal options.');
    }
    ctx.updatePartXmlInternal(ctx.mainPath, (document) => {
      const body = bodyOf(document);
      const control = descendants(body, 'sdt').find((candidate) => {
        const properties = children(candidate, 'sdtPr')[0];
        return wordValue(children(properties ?? candidate, 'id')[0]) === String(id);
      });
      if (!control) throw new Error(`Content control ${id} does not exist.`);
      const parent = control.parentNode;
      if (!parent) throw new Error('Content control is detached.');
      if (options.keepContent !== false) {
        const content = children(control, 'sdtContent')[0];
        if (content) {
          while (content.firstChild) parent.insertBefore(content.firstChild, control);
        }
      }
      parent.removeChild(control);
      if (options.keepContent === false) {
        let structuralContainer: Element | null = parent as Element;
        while (structuralContainer && structuralContainer !== body && structuralContainer.localName !== 'tc') {
          structuralContainer = structuralContainer.parentNode as Element | null;
        }
        if (structuralContainer?.localName === 'tc' && childrenThroughTransparent(structuralContainer, 'p').length === 0) {
          const paragraph = wordElement(document, 'p');
          const properties = children(structuralContainer, 'tcPr')[0];
          structuralContainer.insertBefore(paragraph, properties?.nextSibling ?? null);
        } else if (structuralContainer === body && childrenThroughTransparent(body, 'p').length === 0 &&
            childrenThroughTransparent(body, 'tbl').length === 0) {
          const paragraph = wordElement(document, 'p');
          body.insertBefore(paragraph, children(body, 'sectPr')[0] ?? null);
        }
      }
      return true;
    });
}
