"use strict";

const $ = (id) => document.getElementById(id);
const state = { bootstrap: null, draft: null, sourceJob: null, uploads: [], images: [], busy: false };
const requestHeaders = { "X-SuperSlide-Request": "1" };
let photoTarget = null;

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
  for (const id of ["create", "download-pptx", "files", "task", "slide-count"]) $(id).disabled = busy;
  document.querySelectorAll(".slide-title, .remove-file, .image-cell, .image-upload, .image-choose, .catalog-image, .image-search").forEach((element) => { element.disabled = busy; });
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
function downloadStatus(message, error = false) {
  $("download-status").textContent = message;
  $("download-status").classList.toggle("error", error);
}
function invalidateDownload() { downloadStatus(""); }
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
async function poll(id, progress = status) {
  for (;;) {
    const job = await api(`/api/jobs/${id}`);
    const elapsed = Number(job.elapsed_seconds || 0);
    progress(job.message + (job.status === "running" && elapsed >= 1 ? ` · ${Math.floor(elapsed)} с` : ""));
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
function slideCountLabel(count) {
  const last = count % 10, pair = count % 100;
  return `${count} ${last === 1 && pair !== 11 ? "слайд" : last >= 2 && last <= 4 && (pair < 12 || pair > 14) ? "слайда" : "слайдов"}`;
}
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
    content.append(title);
    if (slide.layout_id === "title_slide") {
      content.append(imagePicker("Фото обложки (вариант «Фотообложка»)", slide.image, slide, true));
    } else if (slide.layout_id === "photo_text" || slide.layout_id === "image_spotlight") {
      content.append(imagePicker("Фотография слайда", slide.image, slide, false));
    } else if (slide.layout_id === "gallery") {
      slide.images.forEach((item, index) => content.append(imagePicker(`Фото ${index + 1}`, item, slide, false)));
    }
    row.append(content); list.append(row);
  });
  $("slides-subtitle").textContent = `${slideCountLabel(state.draft.slides.length)} · заголовки и изображения`;
  $("slides-section").hidden = false;
  invalidateDownload();
}
function sourceParagraphs(slide, fieldId, field) {
  const update = slide.updates.find((item) => item.field_id === fieldId);
  return update?.paragraphs?.slice() || (update?.text !== undefined ? update.text.split("\n") : field?.paragraphs?.slice() || (field?.sample_text || "").split("\n"));
}
function sourceTitle(slide, variant) {
  const id = variant?.title_field_id;
  const paragraphs = sourceParagraphs(slide, id, variant?.fields?.[id]);
  return (variant?.title_paragraph_only ? paragraphs[variant.title_paragraph_index || 0] || "" : paragraphs.join(" ")).replace(/\s+/g, " ").trim();
}
function setSourceUpdate(slide, update) {
  const index = slide.updates.findIndex((item) => item.field_id === update.field_id);
  if (index < 0) slide.updates.push(update);
  else slide.updates[index] = update;
}
function writeSourceTitle(slide, variant, value) {
  const id = variant.title_field_id;
  if (variant.title_paragraph_only) {
    const paragraphs = sourceParagraphs(slide, id, variant.fields[id]);
    paragraphs[variant.title_paragraph_index || 0] = value;
    setSourceUpdate(slide, { field_id: id, paragraphs });
  } else setSourceUpdate(slide, { field_id: id, text: value });
}
function sourcePhotoChoice(slide, field) {
  const update = slide.updates.find((item) => item.field_id === field.field_id);
  const id = update ? update.photo_id : field.photo_id;
  return state.images.find((item) => item.id === id);
}
function chooseSourcePhoto(target, image) {
  setSourceUpdate(target.slide, { field_id: target.field.field_id, photo_id: image.id });
  invalidateDownload();
  renderSlides();
  status("Изображение выбрано. При сборке оно заполнит рамку с сохранением пропорций.");
}
async function uploadSourcePhoto(file, target) {
  if (!file || state.busy) return;
  setBusy(true);
  try {
    status(`Загрузка изображения: ${file.name}`);
    const uploaded = await api(`/api/images?filename=${encodeURIComponent(file.name)}`, {
      method: "POST", headers: { ...requestHeaders, "Content-Type": "application/octet-stream" }, body: file
    });
    state.images.push(uploaded);
    chooseSourcePhoto(target, uploaded);
    if ($("image-dialog").open) $("image-dialog").close();
    status("Изображение загружено и добавлено на слайд.");
  } catch (error) {
    status(error.message, true);
    $("image-dialog-status").textContent = error.message;
  } finally { setBusy(false); }
}
function renderImageCatalog() {
  const grid = $("image-catalog");
  grid.replaceChildren();
  const query = $("image-search").value.trim().toLocaleLowerCase("ru");
  const selected = photoTarget && sourcePhotoChoice(photoTarget.slide, photoTarget.field);
  const images = state.images.filter((item) => [item.name, item.description, ...(item.tags || [])].join(" ").toLocaleLowerCase("ru").includes(query));
  $("image-count").textContent = images.length ? `Найдено: ${images.length}` : "Изображения не найдены. Измените запрос или загрузите своё.";
  for (const item of images) {
    const name = item.name || item.description || "Изображение";
    const button = element("button", "catalog-image" + (selected?.id === item.id ? " selected" : ""));
    button.type = "button";
    button.setAttribute("aria-label", `Выбрать фото: ${name}`);
    button.setAttribute("aria-pressed", String(selected?.id === item.id));
    button.disabled = state.busy;
    const image = element("img"); image.src = item.url; image.alt = ""; image.loading = "lazy";
    button.append(image, element("span", "", name));
    button.title = item.description || name;
    button.addEventListener("click", () => {
      if (state.busy || !photoTarget) return;
      const target = photoTarget;
      chooseSourcePhoto(target, item);
      $("image-dialog").close();
    });
    grid.append(button);
  }
}
function openSourcePhotoCatalog(target) {
  if (state.busy) return;
  photoTarget = target;
  $("image-dialog-context").textContent = `Слайд ${state.draft.slides.indexOf(target.slide) + 1} · ${target.label}`;
  $("image-search").value = "";
  $("image-dialog-status").textContent = "";
  renderImageCatalog();
  $("image-dialog").showModal();
  $("image-search").focus();
}
function sourcePhotoPicker(slide, field, index, count) {
  const label = count === 1 ? "Изображение" : `Изображение ${index + 1}`;
  const target = { slide, field, label };
  const choice = sourcePhotoChoice(slide, field);
  const box = element("div", "source-photo-picker");
  const preview = element("div", "source-photo-preview");
  if (choice) { const image = element("img"); image.src = choice.url; image.alt = choice.name || choice.description || label; image.loading = "lazy"; preview.append(image); }
  else preview.append(element("span", "", "Фото из макета"));
  const details = element("div", "source-photo-details");
  details.append(element("div", "image-kind", label));
  if (choice) details.append(element("p", "source-photo-name", choice.name || choice.description));
  const actions = element("div", "source-photo-actions");
  const choose = element("button", "image-choose", "Выбрать из каталога");
  choose.type = "button"; choose.setAttribute("aria-label", `${label}: выбрать из каталога`);
  choose.addEventListener("click", () => openSourcePhotoCatalog(target));
  const upload = element("button", "image-upload", "Загрузить с ПК");
  upload.type = "button"; upload.setAttribute("aria-label", `${label}: загрузить с ПК`);
  const input = element("input"); input.type = "file"; input.accept = "image/png,image/jpeg"; input.hidden = true;
  upload.addEventListener("click", () => input.click());
  input.addEventListener("change", () => { const file = input.files[0]; input.value = ""; uploadSourcePhoto(file, target); });
  actions.append(choose, upload, input); details.append(actions); box.append(preview, details);
  return box;
}
function renderSourceSlides() {
  const list = $("slide-list");
  list.replaceChildren();
  state.draft.slides.forEach((slide, index) => {
    const layout = layoutFor(slide.layout_id);
    const variant = layout?.variants.find((item) => item.id === slide.variant_id);
    const row = element("li", "slide-row");
    const content = element("div");
    content.append(element("div", "slide-kind", layout?.name || slide.layout_id));
    if (variant?.title_field_id) {
      const fieldId = variant.title_field_id, field = variant.fields[fieldId];
      const label = element("label", "slide-field-label", "Заголовок слайда");
      const input = element("input", "slide-title");
      input.type = "text"; input.id = `title-${index}`; label.htmlFor = input.id;
      input.value = sourceTitle(slide, variant);
      input.setAttribute("aria-label", `Заголовок слайда ${index + 1}`);
      let limit = field.max_characters;
      if (variant.title_paragraph_only) limit -= sourceParagraphs(slide, fieldId, field).reduce((sum, text, i) => sum + (i === (variant.title_paragraph_index || 0) ? 0 : text.length), 0);
      if (Number.isFinite(limit)) input.maxLength = Math.max(0, limit);
      input.addEventListener("input", () => { writeSourceTitle(slide, variant, input.value); invalidateDownload(); });
      content.append(label, input);
    }
    const photos = (variant?.photo_fields || []).slice();
    for (const update of slide.updates) {
      const field = variant?.fields?.[update.field_id];
      if (update.photo_id && field?.kind === "picture" && field.scope !== "master" && !photos.some((item) => item.field_id === update.field_id)) {
        photos.push({ field_id: update.field_id, name: field.name });
      }
    }
    photos.forEach((field, i) => content.append(sourcePhotoPicker(slide, field, i, photos.length)));
    row.append(content); list.append(row);
  });
  $("slides-subtitle").textContent = `${slideCountLabel(state.draft.slides.length)} · заголовки и изображения`;
  $("slides-section").hidden = false;
  invalidateDownload();
}
$("image-search").addEventListener("input", renderImageCatalog);
$("image-dialog-close").addEventListener("click", () => $("image-dialog").close());
$("image-dialog").addEventListener("close", () => { photoTarget = null; });
$("image-dialog-upload").addEventListener("click", () => $("image-dialog-file").click());
$("image-dialog-file").addEventListener("change", () => {
  const file = $("image-dialog-file").files[0], target = photoTarget;
  $("image-dialog-file").value = "";
  if (target) uploadSourcePhoto(file, target);
});
function setDraft(draft, sourceJob, warnings) {
  state.draft = draft; state.sourceJob = sourceJob;
  renderSlides(); showWarnings(warnings);
  status("Структура готова. Проверьте слайды и скачайте PPTX.");
  $("slides-section").scrollIntoView({ behavior: "smooth", block: "start" });
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
async function downloadPresentation(url, filename) {
  const response = await fetch(url);
  if (!response.ok) {
    let message = "Не удалось скачать PPTX. Повторите попытку.";
    try { message = (await response.json()).error?.message || message; } catch {}
    throw new Error(message);
  }
  const blob = await response.blob();
  if (!blob.size) throw new Error("Файл презентации пуст. Повторите попытку.");
  const objectUrl = URL.createObjectURL(blob);
  const link = element("a");
  link.href = objectUrl; link.download = filename || "Презентация.pptx"; link.hidden = true;
  document.body.append(link);
  link.click(); link.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 60000);
}
$("download-pptx").addEventListener("click", async () => {
  if (!state.draft || state.busy) return;
  if (state.draft.template_version !== "2.0" && state.draft.slides.some((slide) => !(slide.layout_id === "quote" ? slide.quote : slide.title)?.trim())) {
    downloadStatus("Заполните все заголовки и цитаты.", true); return;
  }
  setBusy(true); invalidateDownload();
  const button = $("download-pptx");
  button.textContent = "Создаём PPTX…";
  button.setAttribute("aria-busy", "true");
  try {
    downloadStatus("Сборка презентации…");
    const started = await api("/api/build", jsonOptions({ draft: state.draft, source_job_id: state.sourceJob }));
    const job = await poll(started.job_id, downloadStatus);
    downloadStatus("Скачивание PPTX…");
    await downloadPresentation(job.downloads.pptx, job.download_filename);
    showWarnings(job.warnings);
    downloadStatus(`Файл «${job.download_filename}» готов. Скачивание началось.`);
  } catch (error) { downloadStatus(error.message, true); }
  finally {
    button.textContent = "Скачать PPTX";
    button.removeAttribute("aria-busy");
    setBusy(false);
  }
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
  catch (error) { status(error.message, true); $("create").disabled = true; }
}
init();
