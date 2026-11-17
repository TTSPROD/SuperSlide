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
  vm.runInContext(source + '\nglobalThis.editor = { sourceTitle, writeSourceTitle };', context);
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

test('editing an inherited mixed heading preserves body soft breaks and the selected layout', () => {
  const api = editor();
  const layout = variant('slide_text_15', ['Заголовок', 'Первая\vвторая строка', 'Автор проекта'], true, 0);
  const original = clone(layout);
  const slide = { variant_id: 'slide_010', updates: [] };
  api.writeSourceTitle(slide, layout, 'Результат проекта');
  assert.deepEqual(clone(slide), {
    variant_id: 'slide_010',
    updates: [{ field_id: 'slide_text_15', paragraphs: ['Результат проекта', 'Первая\vвторая строка', 'Автор проекта'] }],
  });
  assert.equal(api.sourceTitle(slide, layout), 'Результат проекта');
  assert.deepEqual(clone(layout), original);
});

test('editing a noninitial heading preserves existing generated body paragraphs', () => {
  const api = editor();
  const layout = variant('slide_text_26', ['Образец', 'Заголовок', 'Подпись'], true, 1);
  const slide = { variant_id: 'slide_011', updates: [
    { field_id: 'slide_text_26', paragraphs: ['Содержание\vс переносом', 'Старый заголовок', 'Анна Соколова'] },
  ] };
  api.writeSourceTitle(slide, layout, 'Новый заголовок');
  assert.deepEqual(clone(slide.updates), [
    { field_id: 'slide_text_26', paragraphs: ['Содержание\vс переносом', 'Новый заголовок', 'Анна Соколова'] },
  ]);
  assert.equal(slide.variant_id, 'slide_011');
});

test('editing a plain heading preserves other updates, notes and source references', () => {
  const api = editor();
  const layout = variant('slide_text_32', ['Заголовок']);
  const slide = { layout_id: 'title', variant_id: 'slide_002', notes: 'Заметка докладчику', source_refs: ['chunk_1'], updates: [
    { field_id: 'slide_picture_16', photo_id: 'worker_portrait_orange_helmet' },
    { field_id: 'slide_text_32', paragraphs: ['Прежний заголовок'] },
    { field_id: 'slide_text_10', text: 'Основное содержание' },
  ] };
  const expected = structuredClone(slide);
  expected.updates[1] = { field_id: 'slide_text_32', text: 'Итоги проекта' };
  api.writeSourceTitle(slide, layout, 'Итоги проекта');
  assert.deepEqual(clone(slide), expected);
  assert.equal(api.sourceTitle(slide, layout), 'Итоги проекта');
});

test('repeated title edits replace the same update without duplicating or resetting body text', () => {
  const api = editor();
  const layout = variant('slide_text_15', ['Заголовок', 'Исходный образец'], true, 0);
  const slide = { variant_id: 'slide_010', updates: [
    { field_id: 'slide_text_15', paragraphs: ['Первая версия', 'Подтверждённое содержание\vвторой строки'] },
  ] };
  for (const title of ['Вторая версия', '', 'Окончательный заголовок']) api.writeSourceTitle(slide, layout, title);
  assert.deepEqual(clone(slide.updates), [
    { field_id: 'slide_text_15', paragraphs: ['Окончательный заголовок', 'Подтверждённое содержание\vвторой строки'] },
  ]);
  assert.equal(api.sourceTitle(slide, layout), 'Окончательный заголовок');
  assert.equal(slide.variant_id, 'slide_010');
});
