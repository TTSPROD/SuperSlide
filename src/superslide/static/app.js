"use strict";

const $ = (id) => document.getElementById(id);
const state = { bootstrap: null, draft: null, sourceJob: null, uploads: [], images: [], busy: false };
const requestHeaders = { "X-SuperSlide-Request": "1" };

async function api(path, options = {}) {
  const response = await fetch(path, options);
  let data;
  try { data = await response.json(); } catch { throw new Error("Не удалось прочитать ответ приложения."); }
  if (!response.ok) throw new Error(data.error?.message || "Не удалось выполнить запрос.");
  return data;
}
function jsonOptions(value) {
  return { method: "POST", headers: { ...requestHeaders, "Content-Type": "application/json" }, body: JSON.stringify(value) };
}
function status(message, error = false) {
  $("status").textContent = message;
  $("status").classList.toggle("error", error);
}
function setBusy(busy) {
  state.busy = busy;
  for (const id of ["create", "demo", "import", "build", "files", "task", "slide-count"]) $(id).disabled = busy;
  document.querySelectorAll(".slide-title, .variant-select, .remove-file, .image-cell, .image-upload").forEach((element) => { element.disabled = busy; });
}
function element(tag, className, value) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (value !== undefined) node.textContent = value;
  return node;
}
function showWarnings(items) {
  const area = $("warnings");
  area.replaceChildren();
  for (const warning of items || []) area.append(element("p", "", warning));
}
function invalidateDownload() { $("downloads").hidden = true; }
function renderFiles() {
  const list = $("file-list");
  list.replaceChildren();
  for (const file of state.uploads) {
    const row = element("li");
    row.append(element("span", "", `${file.name} · ${(file.size / 1048576).toFixed(1)} МБ`));
    const remove = element("button", "remove-file", "Удалить");
    remove.type = "button";
    remove.disabled = state.busy;
    remove.setAttribute("aria-label", `Удалить ${file.name}`);
    remove.addEventListener("click", async () => {
      setBusy(true);
      try {
        await api(`/api/uploads/${file.id}`, { method: "DELETE", headers: requestHeaders });
        state.uploads = state.uploads.filter((item) => item.id !== file.id);
        renderFiles();
        status("Документ удалён.");
      } catch (error) { status(error.message, true); }
      finally { setBusy(false); }
    });
    row.append(remove); list.append(row);
  }
}
async function uploadFiles(files) {
  if (state.busy || !files.length) return;
  setBusy(true);
  try {
    for (const file of files) {
      status(`Загрузка: ${file.name}`);
      const uploaded = await api(`/api/uploads?filename=${encodeURIComponent(file.name)}`, {
        method: "POST", headers: { ...requestHeaders, "Content-Type": "application/octet-stream" }, body: file
      });
      state.uploads.push(uploaded);
      renderFiles();
    }
    status("Документы загружены.");
  } catch (error) { status(error.message, true); }
  finally { $("files").value = ""; setBusy(false); }
}
async function poll(id) {
  for (;;) {
    const job = await api(`/api/jobs/${id}`);
    const elapsed = Number(job.elapsed_seconds || 0);
    status(job.message + (job.status === "running" && elapsed >= 1 ? ` · ${Math.floor(elapsed)} с` : ""));
    if (job.status === "succeeded") return job;
    if (job.status === "failed") {
      if (job.partial_download) {
        const recovery = element("p");
        const link = element("a", "", "Скачать готовые слайды (JSON)");
        link.href = job.partial_download; recovery.append(link); $("warnings").append(recovery);
      }
      throw new Error(job.error || "Не удалось выполнить задание.");
    }
    await new Promise((resolve) => setTimeout(resolve, 1000));
  }
}
function layoutFor(id) { return state.bootstrap.layouts.find((layout) => layout.id === id); }
function applyImage(slide, image, uploaded) {
  if (image) {
    image.asset_id = uploaded.id;
    image.alt_text = uploaded.description.slice(0, 180);
  } else {
    slide.image = { asset_id: uploaded.id, fit: "cover", alt_text: uploaded.description.slice(0, 180) };
    slide.variant_id = "photo";
  }
  renderSlides();
  invalidateDownload();
}
function imagePicker(label, image, slide, allowNone) {
  const box = element("div", "image-picker");
  box.append(element("div", "image-kind", label));
  const strip = element("div", "image-strip");
  if (allowNone) {
    const none = element("button", "image-cell none" + (image ? "" : " selected"), "Без фото");
    none.type = "button";
    none.addEventListener("click", () => {
      slide.image = null;
      if (slide.variant_id === "photo") slide.variant_id = "default";
      renderSlides();
      invalidateDownload();
    });
    strip.append(none);
  }
  for (const item of state.images) {
    const cell = element("button", "image-cell" + (image && image.asset_id === item.id ? " selected" : ""));
    cell.type = "button";
    const pic = element("img");
    pic.src = item.url;
    pic.alt = "";
    cell.append(pic);
    cell.title = item.description;
    cell.setAttribute("aria-label", `Выбрать фото: ${item.description}`);
    cell.addEventListener("click", () => {
      applyImage(slide, image, item);
      status("Фото выбрано. Не забудьте про подпись и уместность.");
    });
    strip.append(cell);
  }
  const upload = element("button", "image-upload", "Загрузить своё…");
  upload.type = "button";
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/png,image/jpeg";
  input.hidden = true;
  upload.addEventListener("click", () => input.click());
  input.addEventListener("change", async () => {
    const file = input.files[0];
    input.value = "";
    if (!file) return;
    setBusy(true);
    try {
      status(`Загрузка фото: ${file.name}`);
      const uploaded = await api(`/api/images?filename=${encodeURIComponent(file.name)}`, {
        method: "POST", headers: { ...requestHeaders, "Content-Type": "application/octet-stream" }, body: file
      });
      state.images.push(uploaded);
      applyImage(slide, image, uploaded);
      status("Фото загружено и выбрано. Проверьте подпись и уместность.");
    } catch (error) { status(error.message, true); }
    finally { setBusy(false); }
  });
  box.append(strip, upload, input);
  return box;
}
function showPreview(slide) {
  const layout = layoutFor(slide.layout_id);
  const variant = layout?.variants.find((item) => item.id === (slide.variant_id || "default"));
  if (variant?.preview_url) {
    $("preview").src = variant.preview_url;
    $("preview").alt = `Пример оформления: ${layout.name}, ${variant.name}`;
    $("preview-caption").textContent = `Пример композиции: ${layout.name} · ${variant.name}`;
  }
}
function renderSlides() {
  if (state.draft.template_version === "2.0") { renderSourceSlides(); return; }
  const list = $("slide-list");
  list.replaceChildren();
  state.draft.slides.forEach((slide, index) => {
    const layout = layoutFor(slide.layout_id);
    const row = element("li", "slide-row");
    const content = element("div");
    content.append(element("div", "slide-kind", layout?.name || slide.layout_id));
    const quote = slide.layout_id === "quote";
    const title = document.createElement(quote ? "textarea" : "input");
    title.className = quote ? "slide-title quote" : "slide-title";
    title.value = quote ? slide.quote : slide.title;
    title.maxLength = quote ? 220 : 100;
    title.setAttribute("aria-label", `${quote ? "Цитата" : "Заголовок"} слайда ${index + 1}`);
    title.addEventListener("input", () => { slide[quote ? "quote" : "title"] = title.value; invalidateDownload(); });
    title.addEventListener("focus", () => showPreview(slide));
    content.append(title);
    if (layout && layout.variants.length > 1) {
      const wrap = element("div", "variant-wrap");
      const label = element("label", "", "Оформление");
      const select = element("select", "variant-select");
      select.id = `variant-${index}`; label.htmlFor = select.id;
      for (const variant of layout.variants) {
        const option = element("option", "", variant.name);
        option.value = variant.id;
        // Photo cover requires an image; other cover variants prohibit one.
        if (slide.layout_id === "title_slide") option.disabled = (variant.id === "photo") !== Boolean(slide.image);
        select.append(option);
      }
      select.value = slide.variant_id || "default";
      select.addEventListener("change", () => {
        slide.variant_id = select.value;
        // The contract allows a cover photo only in the photo variant.
        if (slide.layout_id === "title_slide" && slide.variant_id !== "photo" && slide.image) {
          slide.image = null;
          renderSlides();
        } else {
          showPreview(slide);
        }
        invalidateDownload();
      });
      select.addEventListener("focus", () => showPreview(slide));
      wrap.append(label, select); content.append(wrap);
    }
    if (slide.layout_id === "title_slide") {
      content.append(imagePicker("Фото обложки (вариант «Фотообложка»)", slide.image, slide, true));
    } else if (slide.layout_id === "photo_text" || slide.layout_id === "image_spotlight") {
      content.append(imagePicker("Фотография слайда", slide.image, slide, false));
    } else if (slide.layout_id === "gallery") {
      slide.images.forEach((item, index) => content.append(imagePicker(`Фото ${index + 1}`, item, slide, false)));
    }
    row.append(content); list.append(row);
  });
  $("slides-subtitle").textContent = `${state.draft.slides.length} слайдов · заголовки и оформление можно изменить`;
  $("empty-state").hidden = true;
  $("slides-section").hidden = false;
  invalidateDownload();
  if (state.draft.slides.length) showPreview(state.draft.slides[0]);
}
function renderSourceSlides() {
  const list = $("slide-list");
  list.replaceChildren();
  state.draft.slides.forEach((slide, index) => {
    const layout = layoutFor(slide.layout_id);
    const variant = layout?.variants.find((item) => item.id === slide.variant_id);
    const row = element("li", "slide-row");
    const content = element("div");
    content.append(element("div", "slide-kind", `${index + 1}. ${layout?.name || slide.layout_id}`));
    if (layout) {
      const label = element("label", "", "Исходный вариант");
      const select = element("select", "variant-select");
      select.id = `variant-${index}`; label.htmlFor = select.id;
      for (const choice of layout.variants) {
        const option = element("option", "", choice.name);
        option.value = choice.id; select.append(option);
      }
      select.value = slide.variant_id;
      select.addEventListener("change", () => {
        slide.variant_id = select.value;
        slide.updates = [];
        renderSlides(); invalidateDownload();
        status("Выбран исходный вариант. Заполните его текстовые поля.");
      });
      content.append(label, select);
    }
    if (variant?.description) content.append(element("p", "", variant.description));
    for (const [fieldId, field] of Object.entries(variant?.fields || {})) {
      if (field.kind !== "text") continue;
      const update = slide.updates.find((item) => item.field_id === fieldId);
      const label = element("label", "", field.name || fieldId);
      const input = document.createElement("textarea");
      input.className = "slide-title";
      input.id = `field-${index}-${fieldId}`; label.htmlFor = input.id;
      input.value = update?.paragraphs ? update.paragraphs.join("\n") : update?.text ?? field.sample_text ?? "";
      if (field.max_characters) input.maxLength = field.max_characters;
      input.rows = Math.min(6, Math.max(2, input.value.split("\n").length));
      input.addEventListener("input", () => {
        let record = slide.updates.find((item) => item.field_id === fieldId);
        if (!record) { record = { field_id: fieldId }; slide.updates.push(record); }
        delete record.paragraphs;
        record.text = input.value;
        invalidateDownload();
      });
      input.addEventListener("focus", () => showPreview(slide));
      content.append(label, input);
    }
    row.append(content); list.append(row);
  });
  $("slides-subtitle").textContent = `${state.draft.slides.length} слайдов · исходные варианты и текстовые поля`;
  $("empty-state").hidden = true;
  $("slides-section").hidden = false;
  invalidateDownload();
  if (state.draft.slides.length) showPreview(state.draft.slides[0]);
}
function setDraft(draft, sourceJob, warnings, demo = false) {
  state.draft = draft; state.sourceJob = sourceJob;
  renderSlides(); showWarnings(warnings);
  status(demo ? "Открыт демонстрационный пример." : "Структура готова. Проверьте содержание перед сборкой.");
}

$("draft-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  const task = $("task").value.trim();
  if (!task) { status("Опишите задачу для презентации.", true); $("task").focus(); return; }
  setBusy(true);
  try {
    status("Подготовка запроса к корпоративной LLM…");
    const started = await api("/api/drafts", jsonOptions({ task, document_ids: state.uploads.map((item) => item.id), slide_count: Number($("slide-count").value) }));
    const job = await poll(started.job_id);
    setDraft(job.draft, started.job_id, job.warnings);
  } catch (error) { status(error.message, true); }
  finally { setBusy(false); }
});
$("demo").addEventListener("click", async () => {
  setBusy(true);
  try { const example = await api("/api/example"); setDraft(example.draft, null, example.warnings, true); }
  catch (error) { status(error.message, true); }
  finally { setBusy(false); }
});
$("import").addEventListener("click", () => $("import-file").click());
$("import-file").addEventListener("change", async (event) => {
  const file = event.target.files[0];
  event.target.value = "";
  if (!file) return;
  setBusy(true);
  try {
    if (file.size > 2 * 1024 * 1024) throw new Error("Файл структуры превышает 2 МБ.");
    let draft;
    try { draft = JSON.parse(await file.text()); } catch { throw new Error("Не удалось прочитать JSON-файл."); }
    const imported = await api("/api/import", jsonOptions({ draft }));
    setDraft(imported.draft, null, imported.warnings);
    status("Черновик загружен из файла. Проверьте содержание перед сборкой.");
  } catch (error) { status(error.message, true); }
  finally { setBusy(false); }
});
$("build").addEventListener("click", async () => {
  if (!state.draft) return;
  if (state.draft.template_version !== "2.0" && state.draft.slides.some((slide) => !(slide.layout_id === "quote" ? slide.quote : slide.title)?.trim())) {
    status("Заполните все заголовки и цитаты.", true); return;
  }
  setBusy(true); invalidateDownload();
  try {
    status("Сборка презентации…");
    const started = await api("/api/build", jsonOptions({ draft: state.draft, source_job_id: state.sourceJob }));
    const job = await poll(started.job_id);
    $("download-pptx").href = job.downloads.pptx;
    $("download-json").href = job.downloads.json;
    $("downloads").hidden = false;
    showWarnings(job.warnings);
    status("Презентация готова к скачиванию.");
  } catch (error) { status(error.message, true); }
  finally { setBusy(false); }
});
$("files").addEventListener("change", (event) => uploadFiles([...event.target.files]));
for (const name of ["dragenter", "dragover"]) $("dropzone").addEventListener(name, (event) => { event.preventDefault(); $("dropzone").classList.add("dragging"); });
for (const name of ["dragleave", "drop"]) $("dropzone").addEventListener(name, (event) => { event.preventDefault(); $("dropzone").classList.remove("dragging"); });
$("dropzone").addEventListener("drop", (event) => uploadFiles([...event.dataTransfer.files]));
$("settings-open").addEventListener("click", () => { $("settings-status").textContent = ""; $("settings-dialog").showModal(); });
$("settings-close").addEventListener("click", () => { $("api-key").value = ""; $("settings-dialog").close(); });
$("settings-dialog").addEventListener("close", () => { $("api-key").value = ""; });
$("settings-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  const area = $("settings-status");
  try {
    await api("/api/settings", jsonOptions({ api_key: $("api-key").value }));
    $("api-key").value = ""; area.classList.remove("error"); area.textContent = "Ключ сохранён на время работы приложения.";
    if (state.bootstrap) { state.bootstrap.key_configured = true; updateConnection(); }
  } catch (error) { area.classList.add("error"); area.textContent = error.message; }
});
$("connection-check").addEventListener("click", async () => {
  const area = $("settings-status");
  const button = $("connection-check");
  button.disabled = true; area.classList.remove("error");
  try {
    if ($("api-key").value.trim()) {
      await api("/api/settings", jsonOptions({ api_key: $("api-key").value }));
      $("api-key").value = "";
      state.bootstrap.key_configured = true; updateConnection();
    }
    const started = await api("/api/connection-check", jsonOptions({}));
    for (;;) {
      const job = await api(`/api/jobs/${started.job_id}`);
      area.textContent = job.status === "queued" ? "Проверка в очереди…" : `Ожидание ответа модели · ${Math.floor(job.elapsed_seconds || 0)} с`;
      if (job.status === "failed") throw new Error(job.error);
      if (job.status === "succeeded") {
        area.textContent = `Модель ${job.model} ответила. Проверка заняла ${job.elapsed_seconds} с.`;
        break;
      }
      await new Promise((resolve) => setTimeout(resolve, 1000));
    }
  } catch (error) { area.classList.add("error"); area.textContent = error.message; }
  finally { button.disabled = false; }
});
function updateConnection() {
  const info = state.bootstrap;
  $("connection-info").textContent = info.llm_configured ? `Модель: ${info.model} · ${info.key_configured ? "ключ задан" : "ключ не задан"}` : "Модель не настроена. Проверьте локальную конфигурацию.";
  $("preview-caption").textContent = `${info.design_count} композиция в фирменном стиле`;
}
async function init() {
  for (let count = 3; count <= 20; count++) {
    const option = element("option", "", String(count)); option.value = String(count); option.selected = count === 8; $("slide-count").append(option);
  }
  try {
    const [bootstrap, images] = await Promise.all([api("/api/bootstrap"), api("/api/images")]);
    state.bootstrap = bootstrap;
    state.images = images.images;
    updateConnection();
  }
  catch (error) { status(error.message, true); $("create").disabled = true; $("demo").disabled = true; }
}
init();
