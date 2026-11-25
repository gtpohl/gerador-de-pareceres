'use strict';
const $ = id => document.getElementById(id);
const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const fields = ['professor', 'disciplina', 'conteudo', 'data'];
const agreement = {
  M: {sujeito:'O aluno', referencia:'o aluno', pronome:'ele', aprovacao:'pudesse ser aprovado'},
  F: {sujeito:'A aluna', referencia:'a aluna', pronome:'ela', aprovacao:'pudesse ser aprovada'}
};
let students = [];
let teachers = new Map();
const parser = new DOMParser();
const serializer = new XMLSerializer();
const elements = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
let roster = [];
const schoolListName = 'Campus Vila Andrade';
const classLists = new Map();
const classListStorage = 'parecer-retencao-class-lists-v1';
const signatures = {};
const signers = [
  {key:'director', input:'directorSignature', preview:'directorSignaturePreview', clear:'clearDirectorSignature', label:'Diretor', id:7901, tab:1400},
  {key:'teacher', input:'signature', preview:'signaturePreview', clear:'clearSignature', label:'Professor', id:7900, tab:4896}
];
let studentSequence = 0;
const batch = new Map();
let paragraphs = [];
const reportModels = {
  '2026': {word:'modelo.docx', pdf:'cabecalho.pdf', text:'texto.json', logo:'logo.png', pdfTop:214},
  '2025': {word:'modelo-2025.docx', pdf:'cabecalho-2025.pdf', text:'texto-2025.json', logo:'logo-2025.png', pdfTop:229}
};
const modelTexts = new Map();

function getModelText(model) {
  if (!modelTexts.has(model)) {
    const pending = fetch(model.text).then(response => {
      if (!response.ok) throw new Error('Texto do modelo indisponível.');
      return response.json();
    }).catch(error => {
      modelTexts.delete(model);
      throw error;
    });
    modelTexts.set(model, pending);
  }
  return modelTexts.get(model);
}

async function updateReportModel() {
  const year = $('reportModel').value;
  const model = reportModels[year];
  $('modelYear').textContent = year;
  $('modelLogo').src = model.logo;
  $('modelStatus').textContent = 'Carregando o modelo…';
  paragraphs = [];
  render();
  try {
    const texts = await getModelText(model);
    if ($('reportModel').value !== year) return;
    paragraphs = texts;
    $('modelStatus').textContent = '';
    render();
  } catch (error) {
    if ($('reportModel').value === year) $('modelStatus').textContent = 'Prévia indisponível. Recarregue a página ou troque de modelo.';
  }
}

function parse(source) {
  const doc = parser.parseFromString(source, 'application/xml');
  if (elements(doc, 'parsererror').length) throw new Error('XML inválido.');
  return doc;
}

async function readXml(zip, path) {
  const entry = zip.file(path);
  if (!entry) throw new Error('Arquivo interno ausente: ' + path);
  return parse(await entry.async('string'));
}

function values() {
  const result = Object.fromEntries(fields.map(key => [key, $(key).value.trim()]));
  if (teachers.size) {
    result.professor = $('teacherSelect').value;
    result.disciplina = $('subjectSelect').value;
  }
  result.data = result.data.split('-').reverse().join('/');
  if (result.conteudo && !/[.!?…]$/.test(result.conteudo)) result.conteudo += '.';
  return result;
}

function interpolate(text, data) {
  return text.replace(/\{\{(\w+)\}\}/g, (_, key) => data[key] || ({conteudo:'[conteúdos da disciplina]', aluno:'[nome do aluno]', sujeito:'O(A) aluno(a)', referencia:'o(a) aluno(a)', pronome:'ele(a)', aprovacao:'pudesse obter aprovação'})[key] || '—');
}

function studentData(student) {
  const treatment = student?.override || student?.gender || inferGender(student?.name);
  const wording = agreement[treatment] || {
    sujeito:'Estudante', referencia:'quem precisava', pronome:student?.name || 'a pessoa estudante', aprovacao:'pudesse obter aprovação'
  };
  return {...values(), treatment, aluno:student?.name, classe:student?.className || (!roster.length ? $('classe').value.trim() : ''), ...wording};
}

function manualStudent() {
  return {name:$('aluno').value.trim(), override:$('gender').value};
}

function render() {
  const student = roster.length ? batch.get($('previewStudent').value) : manualStudent();
  const data = studentData(student);
  document.querySelectorAll('[data-field]').forEach(node => {
    node.textContent = data[node.dataset.field] || '—';
  });
  $('paragraphs').replaceChildren(...paragraphs.map(text => {
    const p = document.createElement('p');
    p.textContent = interpolate(text, data);
    return p;
  }));
}

function compactName(text) {
  return text.normalize('NFD').replace(/\p{Diacritic}/gu, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

const nameLabels = /^(nome|nomes|aluno|aluna|alunos|alunas|estudante|estudantes|nome completo|nome d[oa] alun[oa]|nome do a aluno a|nome d[oa] estudante)$/;

async function readWorkbook(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const book = await readXml(zip, 'xl/workbook.xml');
  const connections = await readXml(zip, 'xl/_rels/workbook.xml.rels');
  const links = new Map(elements(connections, 'Relationship').map(item => [item.getAttribute('Id'), item.getAttribute('Target')]));
  const shared = zip.file('xl/sharedStrings.xml')
    ? elements(await readXml(zip, 'xl/sharedStrings.xml'), 'si').map(item => elements(item, 't').map(t => t.textContent).join(''))
    : [];
  const tabs = [];
  for (const item of elements(book, 'sheet')) {
    const target = links.get(item.getAttributeNS(rel, 'id') || item.getAttribute('r:id'));
    if (!target || !target.includes('worksheets/')) continue;
    const path = new URL(target, 'https://xlsx.invalid/xl/workbook.xml').pathname.slice(1);
    const sheet = await readXml(zip, path);
    const rows = elements(sheet, 'row').map(row => {
      const cells = new Map();
      for (const cell of elements(row, 'c')) {
        const column = (cell.getAttribute('r') || '').match(/^[A-Z]+/i)?.[0];
        if (!column) continue;
        const raw = elements(cell, 'v')[0]?.textContent ?? '';
        const type = cell.getAttribute('t');
        const value = type === 's' ? shared[Number(raw)] : type === 'inlineStr'
          ? elements(cell, 't').map(t => t.textContent).join('') : raw;
        if (value?.trim()) cells.set(column, value.trim());
      }
      return cells;
    }).filter(row => row.size);
    tabs.push({name: item.getAttribute('name'), rows});
  }
  if (!tabs.some(tab => tab.rows.length)) throw new Error('Nenhuma aba contém dados.');
  return tabs;
}

function selectOptions(select, entries, placeholder) {
  select.replaceChildren(...(placeholder ? [new Option(placeholder, '')] : []), ...entries.map(([label, value]) => new Option(label, value)));
}

function inferGender(name) {
  const first = compactName(name || '').split(' ')[0];
  const female = givenNames.F.has(first);
  const male = givenNames.M.has(first);
  if (female === male) return '';
  return female ? 'F' : 'M';
}

function genderCode(value) {
  const code = compactName(value || '');
  if (/^(m|masculino|aluno|homem)$/.test(code)) return 'M';
  if (/^(f|feminino|aluna|mulher)$/.test(code)) return 'F';
  return '';
}

function setSelected(student, selected) {
  if (selected) batch.set(student.id, student);
  else batch.delete(student.id);
  if (student.checkbox) student.checkbox.checked = selected;
}

function updateSelection(preferred) {
  const preview = $('previewStudent');
  const current = preferred || preview.value;
  const options = Array.from(batch, ([id, student]) => [student.name + ' — ' + student.className, id]);
  selectOptions(preview, options, options.length ? '' : 'Selecione alunos nas turmas');
  if (batch.has(current)) preview.value = current;
  preview.disabled = $('removeFromBatch').disabled = !batch.size;
  $('batchCount').textContent = batch.size + (batch.size === 1 ? ' aluno no lote' : ' alunos no lote');
  const count = students.filter(student => batch.has(student.id)).length;
  $('selectionCount').textContent = count + ' de ' + students.length + ' selecionados nesta turma';
  render();
}

function updateStudents() {
  const tab = roster[Number($('sheet').value)];
  const column = $('column').value;
  const hasHeader = $('header').checked;
  const genderColumn = hasHeader && tab.rows[0]
    ? [...tab.rows[0]].find(([, label]) => /^(sexo|genero|tratamento)$/.test(compactName(label)))?.[0] : null;
  if (!tab.students || tab.column !== column || tab.hasHeader !== hasHeader) {
    for (const student of tab.students || []) batch.delete(student.id);
    tab.column = column;
    tab.hasHeader = hasHeader;
    tab.students = tab.rows.slice(hasHeader ? 1 : 0).flatMap(row => {
      const name = row.get(column);
      return name ? [{id:String(++studentSequence), name, className:tab.name, gender:genderCode(row.get(genderColumn))}] : [];
    });
  }
  students = tab.students;
  $('students').replaceChildren(...students.map(student => {
    const row = document.createElement('div');
    row.className = 'student-row';
    const label = document.createElement('label');
    label.className = 'check';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = batch.has(student.id);
    checkbox.addEventListener('change', () => {
      setSelected(student, checkbox.checked);
      updateSelection(checkbox.checked ? student.id : undefined);
    });
    const name = document.createElement('strong');
    name.textContent = student.name;
    label.append(checkbox, name);
    const gender = document.createElement('select');
    gender.setAttribute('aria-label', 'Tratamento de ' + student.name);
    const automatic = student.gender || inferGender(student.name);
    const labelText = {F:'Aluna', M:'Aluno'}[automatic] || 'Texto neutro';
    selectOptions(gender, [['Aluno','M'],['Aluna','F']], 'Auto: ' + labelText);
    gender.value = student.override || '';
    gender.addEventListener('change', () => {
      student.override = gender.value;
      render();
    });
    student.checkbox = checkbox;
    row.append(label, gender);
    return row;
  }));
  updateSelection();
}

function updateSheet() {
  const tab = roster[Number($('sheet').value)];
  const columns = [...new Set(tab.rows.flatMap(row => [...row.keys()]))].sort((a,b) => a.length - b.length || a.localeCompare(b));
  const first = tab.rows[0] || new Map();
  $('removeList').hidden = tab.workbook === schoolListName;
  const identified = columns.find(col => nameLabels.test(compactName(first.get(col) || '')));
  const probableName = columns.find(col => tab.rows.some(row => /\p{L}+\s+\p{L}+/u.test(row.get(col) || '')));
  $('header').checked = Boolean(identified || [...first.values()].some(value => /^(ra|matricula|chamada|sexo|genero|tratamento)$/.test(compactName(value))));
  selectOptions($('column'), columns.map(col => [first.has(col) ? col + ' — ' + first.get(col).slice(0, 50) : col, col]));
  $('column').value = tab.column ?? (identified || probableName || columns[0] || '');
  if (tab.hasHeader !== undefined) $('header').checked = tab.hasHeader;
  updateStudents();
}

function saveClassLists() {
  try {
    const saved = Array.from(classLists)
      .filter(([, workbook]) => workbook.remember)
      .map(([name, workbook]) => [name, workbook.tabs.map(tab => ({
        name:tab.name, rows:tab.rows.map(row => [...row]),
        column:tab.column, hasHeader:tab.hasHeader
      }))]);
    if (saved.length) localStorage.setItem(classListStorage, JSON.stringify(saved));
    else localStorage.removeItem(classListStorage);
    return true;
  } catch (error) {
    return false;
  }
}

function refreshClassLists(preferred) {
  const previous = preferred || roster[Number($('sheet').value)];
  roster = [...classLists.values()].flatMap(workbook => workbook.tabs);
  $('pickers').hidden = !roster.length;
  $('previewBatch').hidden = !roster.length;
  $('manualStudent').hidden = Boolean(roster.length);
  for (const id of ['aluno','classe','gender']) $(id).disabled = Boolean(roster.length);
  if (!roster.length) {
    students = [];
    updateSelection();
    return;
  }
  const multiple = classLists.size > 1;
  selectOptions($('sheet'), roster.map((tab,index) => [
    multiple ? tab.name + ' — ' + tab.workbook : tab.name, String(index)
  ]));
  $('sheet').value = String(Math.max(0, roster.indexOf(previous)));
  updateSheet();
}

function removeClassWorkbook(name) {
  const workbook = classLists.get(name);
  for (const tab of workbook?.tabs || []) {
    for (const student of tab.students || []) batch.delete(student.id);
  }
  classLists.delete(name);
}

function restoreClassLists() {
  try {
    const saved = JSON.parse(localStorage.getItem(classListStorage) || '[]');
    for (const [name, tabs] of saved) {
      classLists.set(name, {remember:true, tabs:tabs.map(tab => ({
        ...tab, workbook:name, rows:tab.rows.map(row => new Map(row))
      }))});
    }
  } catch (error) {
    $('importStatus').textContent = 'As listas salvas não puderam ser recuperadas. Você pode importá-las novamente.';
  }
  if (classLists.size) refreshClassLists();
}

let schoolLoad;
$('classSource').addEventListener('change', async () => {
  const source = $('classSource');
  const school = source.value === 'school';
  $('schoolList').hidden = !school;
  $('uploadLists').hidden = school;
  if (!school) return;

  const status = $('schoolStatus');
  status.textContent = 'Carregando as turmas…';
  try {
    schoolLoad ??= (async () => {
      const response = await fetch('turmas-vla.xlsx');
      if (!response.ok) throw new Error('Planilha indisponível');
      const tabs = await readWorkbook(await response.blob());
      for (const tab of tabs) tab.workbook = schoolListName;
      return tabs;
    })().catch(error => {
      schoolLoad = undefined;
      throw error;
    });
    const tabs = await schoolLoad;
    if (source.value !== 'school') return;
    classLists.set(schoolListName, {remember:false,tabs});
    refreshClassLists(tabs[0]);
    status.textContent = tabs.length + ' turmas carregadas. Escolha uma no menu abaixo.';
  } catch (error) {
    // Return to the import option so the user can select the school list again.
    source.value = 'upload';
    source.dispatchEvent(new Event('change'));
    $('importStatus').textContent = 'Não foi possível carregar a lista da escola. Selecione-a novamente ou importe um Excel.';
  }
});

$('file').addEventListener('change', async event => {
  const input = event.target;
  const files = [...input.files];
  if (!files.length) return;
  const remember = $('rememberLists').checked;
  input.disabled = true;
  $('importStatus').textContent = 'Lendo as planilhas…';
  const imported = [];
  const failed = [];
  try {
    for (const file of files) {
      try {
        const tabs = (await readWorkbook(file)).filter(tab => tab.rows.length);
        for (const tab of tabs) tab.workbook = file.name;
        removeClassWorkbook(file.name);
        classLists.set(file.name, {remember, tabs});
        imported.push(tabs[0]);
      } catch (error) {
        failed.push(file.name);
      }
    }
    if (imported.length) refreshClassLists(imported[0]);
    const saved = !imported.length || saveClassLists();
    $('importStatus').textContent = [
      imported.length ? imported.length + (imported.length === 1 ? ' planilha adicionada.' : ' planilhas adicionadas.') : '',
      failed.length ? 'Não foi possível importar: ' + failed.join(', ') + '.' : '',
      saved ? '' : 'O navegador não conseguiu salvar as listas. Elas podem ser usadas nesta sessão.'
    ].filter(Boolean).join(' ');
  } finally {
    input.value = '';
    input.disabled = false;
  }
});

$('removeList').addEventListener('click', () => {
  const tab = roster[Number($('sheet').value)];
  if (!tab || !window.confirm('Remover a planilha “' + tab.workbook + '” e suas turmas deste navegador?')) return;
  removeClassWorkbook(tab.workbook);
  refreshClassLists();
  $('importStatus').textContent = saveClassLists()
    ? 'Planilha removida.' : 'Planilha removida desta sessão. Não foi possível atualizar as listas salvas.';
});

$('sheet').addEventListener('change', updateSheet);
for (const id of ['column','header']) {
  $(id).addEventListener('change', () => {
    updateStudents();
    if (!saveClassLists()) $('importStatus').textContent = 'Não foi possível salvar o ajuste no navegador.';
  });
}

$('previewStudent').addEventListener('change', render);
$('removeFromBatch').addEventListener('click', () => {
  const student = batch.get($('previewStudent').value);
  if (student) setSelected(student, false);
  updateSelection();
});
for (const [id, selected] of [['selectAll', true], ['selectNone', false]]) {
  $(id).addEventListener('click', () => {
    for (const student of students) setSelected(student, selected);
    updateSelection();
  });
}

function updateSubjects() {
  const subjects = teachers.get($('teacherSelect').value);
  selectOptions($('subjectSelect'), Array.from(subjects || [], subject => [subject, subject]), 'Selecione a disciplina');
  if (subjects?.size === 1) $('subjectSelect').selectedIndex = 1;
  render();
}

$('teacherSelect').addEventListener('change', updateSubjects);
$('subjectSelect').addEventListener('change', render);
$('teachersFile').addEventListener('change', async event => {
  const file = event.target.files[0];
  if (!file) return;
  $('teachersStatus').textContent = 'Lendo a lista…';
  try {
    const tabs = await readWorkbook(file);
    const lists = new Map();
    for (const {rows} of tabs) {
      const headerIndex = rows.findIndex(row => {
        const labels = Array.from(row.values(), compactName);
        return labels.some(label => /professor|docente/.test(label)) && labels.some(label => /disciplina|componente curricular/.test(label));
      });
      if (headerIndex < 0) continue;
      const columns = [...rows[headerIndex]];
      const teacherColumn = columns.find(([,label]) => /professor|docente/.test(compactName(label)))[0];
      const subjectColumn = columns.find(([,label]) => /disciplina|componente curricular/.test(compactName(label)))[0];
      for (const row of rows.slice(headerIndex + 1)) {
        const teacher = row.get(teacherColumn);
        const subject = row.get(subjectColumn);
        if (!teacher || !subject) continue;
        if (!lists.has(teacher)) lists.set(teacher, new Set());
        lists.get(teacher).add(subject);
      }
    }
    if (event.target.files[0] !== file) return;
    if (!lists.size) throw new Error('Lista vazia.');
    teachers = lists;
    $('teachersFilename').textContent = file.name;
    $('teachersStatus').textContent = lists.size + (lists.size === 1 ? ' professor cadastrado.' : ' professores cadastrados.');
    $('manualTeacher').hidden = true;
    $('teacherPickers').hidden = false;
    for (const id of ['professor','disciplina']) $(id).disabled = true;
    for (const id of ['teacherSelect','subjectSelect']) $(id).disabled = false;
    selectOptions($('teacherSelect'), [...lists.keys()].sort((a,b) => a.localeCompare(b, 'pt-BR')).map(name => [name,name]), 'Selecione o professor');
    updateSubjects();
  } catch (error) {
    if (event.target.files[0] !== file) return;
    event.target.value = '';
    $('teachersStatus').textContent = 'Use um .xlsx com as colunas Professor e Disciplina preenchidas.';
  }
});

function setSignature(signer, image) {
  signatures[signer.key] = image;
  $(signer.preview).hidden = !image;
  $(signer.clear).hidden = !image;
  if (image) $(signer.preview).src = image.url;
  else $(signer.preview).removeAttribute('src');
}

for (const signer of signers) {
  $(signer.input).addEventListener('change', async event => {
    const file = event.target.files[0];
    if (!file) return;
    $('exportStatus').textContent = '';
    try {
      if (!['image/png', 'image/jpeg'].includes(file.type)) throw new Error('Formato inválido');
      const image = await createImageBitmap(file);
      if (event.target.files[0] !== file) {
        image.close();
        return;
      }
      const canvas = document.createElement('canvas');
      const scale = Math.min(1, 650 / image.width, 260 / image.height);
      canvas.width = Math.max(1, Math.round(image.width * scale));
      canvas.height = Math.max(1, Math.round(image.height * scale));
      canvas.getContext('2d').drawImage(image, 0, 0, canvas.width, canvas.height);
      image.close();
      setSignature(signer, {url: canvas.toDataURL('image/png'), ratio: canvas.width / canvas.height});
    } catch (error) {
      if (event.target.files[0] !== file) return;
      event.target.value = '';
      setSignature(signer, null);
      $('exportStatus').textContent = 'Selecione uma imagem PNG ou JPG válida para a assinatura.';
    }
  });
  $(signer.clear).addEventListener('click', () => {
    $(signer.input).value = '';
    setSignature(signer, null);
  });
}

async function addSignature(zip, doc, image, signer, paragraph) {
  const relationships = await readXml(zip, 'word/_rels/document.xml.rels');
  const root = relationships.documentElement;
  const link = relationships.createElementNS(root.namespaceURI, 'Relationship');
  const id = 'rIdSignature' + signer.id;
  const filename = 'assinatura-' + signer.key + '.png';
  link.setAttribute('Id', id);
  link.setAttribute('Type', rel + '/image');
  link.setAttribute('Target', 'media/' + filename);
  root.append(link);
  zip.file('word/_rels/document.xml.rels', serializer.serializeToString(relationships));
  zip.file('word/media/' + filename, image.url.split(',')[1], {base64:true});
  const width = Math.max(1, Math.round(Math.min(1100000, image.ratio * 255000)));
  const height = Math.max(1, Math.round(width / image.ratio));
  const run = parse(`<w:r xmlns:w="${word}" xmlns:r="${rel}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:drawing><wp:inline><wp:extent cx="${width}" cy="${height}"/><wp:docPr id="${signer.id}" name="Assinatura do ${signer.label}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${signer.id}" name="${filename}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`).documentElement;
  const tab = doc.createElementNS(word, 'w:r');
  tab.append(doc.createElementNS(word, 'w:tab'));
  paragraph.append(tab);
  paragraph.append(doc.importNode(run, true));
}

async function generateWord(data, images, template) {
  const zip = await JSZip.loadAsync(template);
  const doc = await readXml(zip, 'word/document.xml');
  for (const node of elements(doc, 't')) {
    if (!node.textContent.includes('{{')) continue;
    node.textContent = interpolate(node.textContent, data).replace(/[\r\n]+/g, ' ');
  }
  const attached = signers.filter(signer => images[signer.key]);
  if (attached.length) {
    const paragraph = elements(doc, 'p').find(p => p.getAttributeNS(word, 'rsidR') === 'A11A5EED');
    const props = elements(paragraph, 'pPr')[0];
    const align = elements(props, 'jc')[0];
    align.setAttributeNS(word, 'w:val', 'left');
    const tabs = doc.createElementNS(word, 'w:tabs');
    props.insertBefore(tabs, align);
    for (const signer of attached) {
      const stop = doc.createElementNS(word, 'w:tab');
      stop.setAttributeNS(word, 'w:val', 'center');
      stop.setAttributeNS(word, 'w:pos', String(signer.tab));
      tabs.append(stop);
      await addSignature(zip, doc, images[signer.key], signer, paragraph);
    }
  }
  zip.file('word/document.xml', serializer.serializeToString(doc));
  return zip.generateAsync({type:'blob', compression:'DEFLATE', mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}

async function generatePdf(data, images, template, modelText, pdfTop) {
  if (!modelText.length) throw new Error('Texto do modelo indisponível.');
  const {PDFDocument, StandardFonts, rgb} = PDFLib;
  const pdf = await PDFDocument.load(template);
  pdf.setTitle('Parecer de retenção');
  pdf.setSubject('Conselho de classe');
  let page = pdf.getPage(0);
  const {width, height} = page.getSize();
  const margin = 36;
  const contentWidth = width - margin * 2;
  const regular = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const supported = new Set(regular.getCharacterSet());
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');
  let y = height - pdfTop;

  const supports = text => Array.from(text).every(char => supported.has(char.codePointAt(0)));
  function measure(text, size, font) {
    if (supports(text)) return font.widthOfTextAtSize(text, size);
    context.font = (font === bold ? 'bold ' : '') + size + 'px Arial';
    return context.measureText(text).width;
  }
  async function line(text, x, size, font) {
    if (supports(text)) {
      page.drawText(text, {x, y, size, font, color:rgb(0,0,0)});
      return;
    }
    const scale = 3;
    const measured = measure(text, size, font);
    canvas.width = Math.ceil((measured + 4) * scale);
    canvas.height = Math.ceil(size * 1.6 * scale);
    context.scale(scale, scale);
    context.font = (font === bold ? 'bold ' : '') + size + 'px Arial';
    context.fillText(text, 0, size);
    const png = await pdf.embedPng(canvas.toDataURL('image/png'));
    page.drawImage(png, {x, y:y + size - canvas.height / scale, width:canvas.width / scale, height:canvas.height / scale});
  }
  function ensureSpace(required) {
    if (y - required >= 42) return;
    page = pdf.addPage([width,height]);
    y = height - 54;
  }
  function wrap(text, size, font) {
    const lines = [];
    let current = '';
    for (const word of text.replace(/\s+/g, ' ').trim().split(' ')) {
      const combined = current ? current + ' ' + word : word;
      if (measure(combined, size, font) <= contentWidth) {
        current = combined;
        continue;
      }
      if (current) lines.push(current);
      current = '';
      for (const char of word) {
        if (current && measure(current + char, size, font) > contentWidth) {
          lines.push(current);
          current = '';
        }
        current += char;
      }
    }
    if (current) lines.push(current);
    return lines;
  }
  async function block(text, {size=11, font=regular, gap=0, center=false} = {}) {
    for (const value of wrap(text, size, font)) {
      ensureSpace(15);
      const x = center ? (width - measure(value, size, font)) / 2 : margin;
      await line(value, x, size, font);
      y -= 15;
    }
    y -= gap;
  }

  await block('Nome ' + ({F:'da aluna', M:'do aluno'}[data.treatment] || 'completo') + ': ' + data.aluno, {font:bold});
  await block('Classe: ' + data.classe);
  await block('Curso: Ensino Médio', {gap:10});
  await block('Nome do(a) Professor(a): ' + data.professor);
  await block('Disciplina: ' + data.disciplina, {gap:15});
  await block('PARECER SOBRE A RETENÇÃO', {font:bold, size:12, center:true, gap:10});
  for (const text of modelText) await block(interpolate(text, data), {gap:8});
  ensureSpace(88);
  await block('Data: ' + data.data, {gap:22});
  const labels = ['Assinatura do Diretor', 'Assinatura do Professor', 'Assinatura do Responsável'];
  for (const [index, label] of labels.entries()) {
    const center = margin + contentWidth * (index + 0.5) / labels.length;
    const image = images[signers[index]?.key];
    if (image) {
      const png = await pdf.embedPng(image.url);
      const signatureWidth = Math.min(110, 32 * image.ratio);
      page.drawImage(png, {x:center-signatureWidth/2, y:y+2, width:signatureWidth, height:signatureWidth/image.ratio});
    }
    page.drawLine({start:{x:center-65,y}, end:{x:center+65,y}, thickness:.6});
    page.drawText(label, {x:center-measure(label, 9, regular)/2, y:y-13, size:9, font:regular});
  }
  return new Blob([await pdf.save()], {type:'application/pdf'});
}

$('form').addEventListener('input', render);
$('form').addEventListener('submit', async event => {
  event.preventDefault();
  const selected = roster.length ? [...batch.values()] : [manualStudent()];
  if (!selected.length) {
    $('exportStatus').textContent = 'Selecione pelo menos um aluno para o lote.';
    $('selectAll').focus();
    return;
  }
  const data = selected.map(studentData);
  if ([...fields, 'aluno', 'classe'].some(field => data.some(item => !item[field]))) {
    $('exportStatus').textContent = 'Preencha todos os campos obrigatórios.';
    return;
  }
  const format = event.submitter?.value === 'pdf' ? 'pdf' : 'docx';
  const year = $('reportModel').value;
  const model = reportModels[year];
  const images = {...signatures};
  const buttons = [$('download'), $('downloadPdf')];
  buttons.forEach(button => button.disabled = true);
  $('exportStatus').textContent = 'Gerando ' + data.length + (data.length === 1 ? ' parecer…' : ' pareceres…');
  try {
    const [response, modelText] = await Promise.all([
      fetch(format === 'pdf' ? model.pdf : model.word),
      format === 'pdf' ? getModelText(model) : Promise.resolve(null)
    ]);
    if (!response.ok) throw new Error('Modelo indisponível.');
    const template = await response.arrayBuffer();
    const packageZip = data.length > 1 ? new JSZip() : null;
    let blob;
    let name;
    for (const [index, entry] of data.entries()) {
      const safeName = entry.aluno.replace(/[<>:"/\\|?*\u0000-\u001F]/g, '-').slice(0,100);
      name = 'Parecer de retenção ' + year + ' - ' + safeName + '.' + format;
      blob = format === 'pdf' ? await generatePdf(entry, images, template, modelText, model.pdfTop) : await generateWord(entry, images, template);
      if (packageZip) packageZip.file(String(index + 1).padStart(2, '0') + ' - ' + name, await blob.arrayBuffer());
    }
    if (packageZip) {
      name = 'Pareceres de retenção ' + year + ' - ' + format.toUpperCase() + '.zip';
      blob = await packageZip.generateAsync({type:'blob'});
    }
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = name;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    $('exportStatus').textContent = data.length === 1 ? 'Parecer gerado.' : data.length + ' pareceres gerados.';
  } catch (error) {
    console.error(error);
    $('exportStatus').textContent = 'Não foi possível gerar os pareceres. Tente novamente.';
  } finally {
    buttons.forEach(button => button.disabled = false);
  }
});

$('data').value = new Intl.DateTimeFormat('sv-SE', {timeZone:'America/Sao_Paulo', year:'numeric', month:'2-digit', day:'2-digit'}).format(new Date());
$('reportModel').addEventListener('change', updateReportModel);
restoreClassLists();
updateReportModel();
