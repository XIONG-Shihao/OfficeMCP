/** Read-only interpretation of complex content omitted by the SDK model. */
const nodes = (root, name) => root ? Array.from(root.getElementsByTagNameNS('*', name)) : [];
const texts = (root, name = 't') => nodes(root, name).map(node => node.textContent ?? '').join('');
const attribute = (node, name) => Array.from(node?.attributes ?? []).find(a => a.localName === name)?.value;
const bounded = value => ({ text: value.slice(0, 4000), textLength: value.length, truncated: value.length > 4000 });
export function complexObjectInventory(parsed, partNames) {
  const items = [];
  const add = item => {
    const editable = ['textbox', 'equation', 'chart'].includes(item.kind) && !/^word\/comments/i.test(item.partName);
    const index = ['textbox', 'equation'].includes(item.kind) ? Number(item.id.match(/-(\d+)$/)?.[1]) - 1 : undefined;
    items.push({ ...item, readable: item.readable ?? true, mutationTargetAvailable: editable, contentSource: 'read-only-ooxml-package',
    limitations: [...(item.limitations ?? []), 'Guarded object adapters have explicit shape limits; a target is not universal editing support.'],
    ...(editable ? { nativeTarget: { kind: 'opc_object', objectKind: item.kind, objectId: item.id, partName: item.partName, ...(index !== undefined ? { index } : {}) } } : {}) });
  };
  for (const [partName, document] of parsed) {
    if (document.documentElement.localName === 'chartSpace') {
      const title = nodes(document, 'title')[0], chart = nodes(document, 'plotArea')[0];
      const chartType = Array.from(chart?.childNodes ?? []).find(n => n.nodeType === 1 && n.localName.endsWith('Chart'))?.localName;
      add({ id: 'chart:' + partName, kind: 'chart', partName, name: texts(title), ...bounded(texts(title)), chartType,
        series: nodes(chart, 'ser').slice(0, 100).map(series => ({ name: texts(nodes(series, 'tx')[0], 'v') || texts(nodes(series, 'tx')[0]),
          categories: nodes(nodes(series, 'cat')[0], 'v').slice(0, 100).map(n => n.textContent),
          values: nodes(nodes(series, 'val')[0], 'v').slice(0, 100).map(n => n.textContent),
          cacheOnly: true })), dataCoverage: 'Bounded cached chart data; embedded-workbook formulas and external sources are not evaluated.' });
    }
    for (const [index, box] of nodes(document, 'txbxContent').entries()) add({ id: `${partName}#textbox-${index + 1}`, kind: 'textbox', partName, ...bounded(texts(box)) });
    for (const [index, equation] of nodes(document, 'oMath').entries()) add({ id: `${partName}#equation-${index + 1}`, kind: 'equation', partName,
      ...bounded(texts(equation)), expressionXml: equation.toString().slice(0, 8000), expressionTruncated: equation.toString().length > 8000, evaluated: false });
    if (partName.startsWith('word/diagrams/') && document.documentElement.localName === 'dataModel') add({ id: 'smartArt:' + partName, kind: 'smartArt', partName,
      ...bounded(texts(document)), nodes: nodes(document, 'pt').slice(0, 100).map(point => ({ modelId: attribute(point, 'modelId'), ...bounded(texts(point)) })), layoutEditable: false });
    for (const [index, shape] of nodes(document, 'sp').entries()) {
      const geometry = nodes(shape, 'prstGeom')[0];
      if (geometry) add({ id: `${partName}#shape-${index + 1}`, kind: 'shape', partName, preset: attribute(geometry, 'prst'), ...bounded(texts(shape)) });
    }
  }
  for (const partName of partNames.filter(name => name.startsWith('word/embeddings/') && !name.endsWith('/'))) add({ id: 'embedded:' + partName,
    kind: 'embeddedObject', partName, readable: false, text: '', limitations: ['Embedded payload metadata only; its application data is not interpreted or executed.'] });
  return items;
}
