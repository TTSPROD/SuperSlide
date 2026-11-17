"use strict";
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

function editor() {
  // Execute the shipped script; only the initial page controls and bootstrap
  // request need stand-ins to test the editor's state transitions without a DOM.
  const control = () => ({ addEventListener() {}, append() {}, classList: { toggle() {} } });
  const nodes = new Map();
  const context = vm.createContext({
    structuredClone, setTimeout, console,
    document: {
      getElementById(id) { if (!nodes.has(id)) nodes.set(id, control()); return nodes.get(id); },
      createElement: control,
    },
    fetch: async () => ({ ok: true, json: async () => ({ images: [], layouts: [] }) }),
  });
  const source = fs.readFileSync(path.join(__dirname, '../../src/superslide/static/app.js'), 'utf8');
  vm.runInContext(source + '\nglobalThis.editor = { sourceTitle, writeSourceTitle, remapSourceUpdates };', context);
  return context.editor;
}
function variant(id, paragraphs, mixed = false, heading = null) {
  return {
    title_field_id: id, title_paragraph_only: mixed, title_paragraph_index: heading,
    editable_field_ids: [id], photo_fields: [],
    fields: { [id]: { kind: 'text', scope: 'slide', paragraphs, paragraph_count: paragraphs.length } },
  };
}
const clone = value => JSON.parse(JSON.stringify(value));

test('an author in a mixed title field cannot become part of a plain heading', () => {
  const api = editor();
  const previous = variant('slide_text_26', ['Цитата', 'Имя Фамилия'], true, 0);
  const next = variant('slide_text_31', ['LOREM', 'IPSUM']);
  const slide = { updates: [] };
  api.writeSourceTitle(slide, previous, 'Новый заголовок');
  const before = structuredClone(slide);
  assert.throws(() => api.remapSourceUpdates(slide, previous, next), /нельзя сохранить основной текст/);
  assert.deepEqual(clone(slide), clone(before));
});

test('a plain heading enters just the heading paragraph of a mixed field', () => {
  const api = editor();
  const previous = variant('slide_text_31', ['Заголовок']);
  const next = variant('slide_text_26', ['Цитата', 'Автор и описание'], true, 0);
  const slide = { updates: [{ field_id: 'slide_text_31', text: 'Результат проекта' }] };
  assert.deepEqual(clone(api.remapSourceUpdates(slide, previous, next)), [
    { field_id: 'slide_text_26', paragraphs: ['Результат проекта', 'Автор и описание'] },
  ]);
});

test('compatible mixed fields keep their heading and body with native soft breaks', () => {
  const api = editor();
  const previous = variant('old', ['Заголовок', 'Тело'], true, 0);
  const next = variant('new', ['Образец', 'Описание'], true, 0);
  const slide = { updates: [{ field_id: 'old', paragraphs: ['Итоги', 'Первая\vвторая строка'] }] };
  assert.deepEqual(clone(api.remapSourceUpdates(slide, previous, next)), [
    { field_id: 'new', paragraphs: ['Итоги', 'Первая\vвторая строка'] },
  ]);
});

test('a mixed field with a different paragraph structure is rejected without changing the slide', () => {
  const api = editor();
  const previous = variant('old', ['Заголовок', 'Тело'], true, 0);
  for (const next of [variant('new', ['Тело', 'Заголовок'], true, 1), variant('new', ['Заголовок', 'Тело', 'Подпись'], true, 0)]) {
    const slide = { updates: [{ field_id: 'old', paragraphs: ['Итоги', 'Описание'] }] };
    const before = structuredClone(slide);
    assert.throws(() => api.remapSourceUpdates(slide, previous, next), /нельзя сохранить основной текст/);
    assert.deepEqual(clone(slide), clone(before));
  }
});

test('plain multi-paragraph headings retain their content when moving to a plain field', () => {
  const api = editor();
  const previous = variant('old', ['Первая', 'Вторая']);
  const next = variant('new', ['Другая', 'Композиция']);
  const slide = { updates: [{ field_id: 'old', paragraphs: ['Итоги', 'Проекта'] }] };
  assert.deepEqual(clone(api.remapSourceUpdates(slide, previous, next)), [
    { field_id: 'new', paragraphs: ['Итоги', 'Проекта'] },
  ]);
});
