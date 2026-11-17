"use strict";
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

test('blob download keeps the generated Unicode filename returned by the server', async () => {
  let clickedName;
  const control = () => ({ addEventListener() {}, append() {}, classList: { toggle() {} },
    click() { clickedName = this.download; }, remove() {} });
  const nodes = new Map();
  const context = vm.createContext({
    structuredClone, console, setTimeout() {},
    URL: { createObjectURL: () => 'blob:generated-pptx', revokeObjectURL() {} },
    document: { body: { append() {} }, createElement: control,
      getElementById(id) { if (!nodes.has(id)) nodes.set(id, control()); return nodes.get(id); } },
    fetch: async () => ({ ok: true, json: async () => ({ images: [], layouts: [] }),
      blob: async () => ({ size: 1024 }) }),
  });
  const source = fs.readFileSync(path.join(__dirname, '../../src/superslide/static/app.js'), 'utf8');
  vm.runInContext(source + '\nglobalThis.downloadForTest = downloadPresentation;', context);
  await context.downloadForTest('/api/downloads/job/presentation.pptx', 'Шаттловые системы 2D и 3D.pptx');
  assert.equal(clickedName, 'Шаттловые системы 2D и 3D.pptx');
});
