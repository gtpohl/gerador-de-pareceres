'use strict';
const $ = id => document.getElementById(id);
const path = 'word/document.xml';
const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xmlSpace = 'http://www.w3.org/XML/1998/namespace';
const parser = new DOMParser();
const serializer = new XMLSerializer();
let model;
let lines = [];
let edits = [];
let selection;
let sequence = 0;

function parse(source) {
  const documentXml = parser.parseFromString(source, 'application/xml');
  if (documentXml.getElementsByTagName('parsererror').length) throw Error('XML inválido');
  return documentXml;
}

function readLines(doc) {
  const body = doc.getElementsByTagNameNS(ns, 'body')[0];
  if (!body) throw Error('Documento sem corpo');
  let offset = 0;
  return Array.from(body.getElementsByTagNameNS(ns, 'p'), paragraph => {
    const nodes = Array.from(paragraph.getElementsByTagNameNS(ns, 't'))
      .filter(node => nearestParagraph(node) === paragraph);
    const text = nodes.map(node => node.textContent).join('');
    const line = {nodes, text, offset};
    offset += text.length + 1;
    return line;
  });
}

function nearestParagraph(node) {
  for (let parent = node.parentElement; parent; parent = parent.parentElement) {
    if (parent.namespaceURI === ns && parent.localName === 'p') return parent;
  }
}

function drawText() {
  $('paper').replaceChildren(...lines.map(({text,offset}, index) => {
    const line = document.createElement('div');
    line.dataset.index = index;
    let cursor = 0;
    for (const edit of edits) {
      const start = Math.max(0, edit.start - offset);
      const end = Math.min(text.length, edit.end - offset);
      if (start >= end) continue;
      line.append(text.slice(cursor, start));
      const mark = document.createElement('mark');
      mark.textContent = text.slice(start, end);
      line.append(mark);
      cursor = end;
    }
    line.append(text.slice(cursor));
    return line;
  }));
}

function point(node, offset) {
  const paper = $('paper');
  if (node === paper) {
    if (offset < lines.length) return lines[offset].offset;
    const last = lines.at(-1);
    return last.offset + last.text.length;
  }
  const line = (node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement)?.closest('[data-index]');
  if (!line || line.parentElement !== paper) return null;
  const range = document.createRange();
  range.selectNodeContents(line);
  range.setEnd(node, offset);
  return lines[Number(line.dataset.index)].offset + range.toString().length;
}

document.addEventListener('selectionchange', () => {
  const selected = window.getSelection();
  if (!model || !selected.rangeCount || selected.isCollapsed) return;
  const range = selected.getRangeAt(0);
  const start = point(range.startContainer, range.startOffset);
  const end = point(range.endContainer, range.endOffset);
  selection = start !== null && end !== null && end > start ? {start,end} : null;
  $('mark').disabled = !selection;
});

function drawFields(focusId) {
  $('empty').hidden = Boolean(edits.length);
  $('download').disabled = !edits.length;
  $('fields').replaceChildren(...edits.map(edit => {
    const field = document.createElement('div');
    field.className = 'field';
    const excerpt = document.createElement('p');
    excerpt.className = 'excerpt';
    excerpt.textContent = edit.excerpt;
    const label = document.createElement('label');
    label.htmlFor = 'content-' + edit.id;
    label.textContent = 'Novo conteúdo';
    const input = document.createElement('textarea');
    input.id = label.htmlFor;
    input.rows = 3;
    input.value = edit.value;
    input.addEventListener('input', () => { edit.value = input.value; });
    const remove = document.createElement('button');
    remove.type = 'button';
    remove.className = 'secondary';
    remove.textContent = 'Desmarcar trecho';
    remove.addEventListener('click', () => {
      edits = edits.filter(item => item.id !== edit.id);
      drawText();
      drawFields();
      $('status').textContent = 'Trecho desmarcado.';
    });
    field.append(excerpt,label,input,remove);
    return field;
  }));
  if (focusId) $('content-' + focusId)?.focus();
}

$('mark').addEventListener('pointerdown', event => event.preventDefault());
$('mark').addEventListener('click', () => {
  if (!selection) return;
  const {start,end} = selection;
  selection = null;
  $('mark').disabled = true;
  if (edits.some(edit => start < edit.end && end > edit.start)) {
    $('status').textContent = 'Este trecho se sobrepõe a uma seleção anterior. Desmarque-a ou selecione outro trecho.';
    return;
  }
  const excerpt = lines.map(line => line.text).join('\n').slice(start,end);
  if (!excerpt.trim()) {
    $('status').textContent = 'Selecione um trecho com texto.';
    return;
  }
  const id = ++sequence;
  edits.push({id,start,end,excerpt,value:''});
  edits.sort((a,b) => a.start-b.start);
  window.getSelection().removeAllRanges();
  drawText();
  drawFields(id);
  $('status').textContent = 'Trecho marcado.';
});

$('upload').addEventListener('change', async () => {
  const file = $('upload').files[0];
  if (!file) return;
  $('loadStatus').textContent = 'Lendo o documento…';
  try {
    const bytes = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(bytes);
    const data = zip.file(path);
    if (!data) throw Error('Sem documento Word');
    const loaded = readLines(parse(await data.async('string')));
    if (!loaded.some(line => line.text.trim())) throw Error('Sem texto editável');
    if ($('upload').files[0] !== file) return;
    model = {bytes,name:file.name};
    lines = loaded;
    edits = [];
    selection = null;
    $('mark').disabled = true;
    $('editor').hidden = false;
    $('status').textContent = '';
    $('loadStatus').textContent = file.name + ' aberto. A prévia mostra o texto do corpo, incluindo tabelas.';
    drawText();
    drawFields();
  } catch (error) {
    if ($('upload').files[0] !== file) return;
    $('upload').value = '';
    $('loadStatus').textContent = 'Não foi possível abrir o arquivo. Selecione um .docx com texto editável.';
  }
});

function applyEdits(doc, changes) {
  const targetLines = readLines(doc);
  const replacements = new Map();
  for (const {start,end,value} of [...changes].reverse()) {
    let inserted = false;
    for (const {offset,nodes} of targetLines) {
      let position = offset;
      for (const node of nodes) {
        const length = node.textContent.length;
        const from = Math.max(0, start-position);
        const to = Math.min(length, end-position);
        if (from < to) {
          const text = replacements.get(node) ?? node.textContent;
          replacements.set(node, text.slice(0,from) + (inserted ? '' : value) + text.slice(to));
          inserted = true;
        }
        position += length;
      }
    }
  }
  for (const [node,text] of replacements) {
    const [first,...rest] = text.split(/\r\n?|\n/);
    node.setAttributeNS(xmlSpace, 'xml:space', 'preserve');
    node.textContent = first;
    let tail = node;
    for (const line of rest) {
      const prefix = node.prefix ? node.prefix + ':' : '';
      const br = doc.createElementNS(ns,prefix+'br');
      const next = doc.createElementNS(ns,prefix+'t');
      next.setAttributeNS(xmlSpace,'xml:space','preserve');
      next.textContent = line;
      tail.after(br,next);
      tail = next;
    }
  }
}

$('download').addEventListener('click', async () => {
  if (!model || !edits.length) return;
  const {bytes,name} = model;
  const changes = edits.map(edit => ({...edit}));
  $('download').disabled = true;
  try {
    const zip = await JSZip.loadAsync(bytes);
    const doc = parse(await zip.file(path).async('string'));
    applyEdits(doc, changes);
    zip.file(path, serializer.serializeToString(doc));
    const blob = await zip.generateAsync({type:'blob',compression:'DEFLATE',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
    const href = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = href;
    link.download = name.replace(/\.docx$/i,'') + ' - preenchido.docx';
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(href),10000);
    $('status').textContent = 'Documento gerado.';
  } catch (error) {
    $('status').textContent = 'Não foi possível gerar o documento. Tente novamente.';
  } finally {
    $('download').disabled = !edits.length;
  }
});
