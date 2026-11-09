"use strict";

const byId = id => document.getElementById(id);
const fileInput = byId("file");
const drop = byId("drop");
const maxFileMb = Number(drop.dataset.maxFileMb);
const convert = byId("convert");
const download = byId("download");
const cancel = byId("cancel");
const review = byId("review");
const expiry = byId("expiry");
const expiryTime = byId("expiry-time");
const pageSelect = byId("page-select");
const headers = {"X-Local-Converter": "1"};
let selected = null;
let busy = false;
let job = null;
let previewSequence = 0;
let expiryTimer = null;

function steps(current) {
  document.querySelectorAll(".steps li").forEach((li, index) => {
    li.classList.toggle("active", current === index);
    li.classList.toggle("done", current > index);
    if (current === index) li.setAttribute("aria-current", "step");
    else li.removeAttribute("aria-current");
  });
}

function message(target, text = "") {
  const node = byId(target);
  node.textContent = text;
  node.hidden = !text;
}

function setBusy(value) {
  busy = value;
  for (const control of [drop, byId("remove"), ...document.querySelectorAll("[name=lang]")]) {
    control.disabled = value;
  }
  convert.disabled = value || !selected;
}

function choose(files) {
  if (busy) return;
  message("error");
  message("notice");
  if (files.length > 1) return message("error", "Mỗi lượt chỉ nhận một tệp PDF.");
  const file = files[0] || null;
  if (file && !/\.pdf$/i.test(file.name)) return message("error", "Hãy chọn tệp có đuôi .pdf.");
  if (file && file.size > maxFileMb * 1024 * 1024) return message("error", `PDF vượt giới hạn ${maxFileMb} MB.`);
  if (file && file.size === 0) return message("error", "Tệp PDF đang trống.");
  selected = file;
  byId("filename").hidden = !file;
  byId("filename").querySelector("span").textContent = file ? file.name : "";
  byId("progress").hidden = true;
  steps(0);
  setBusy(false);
}

drop.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => {
  if (fileInput.files.length) choose(fileInput.files);
  fileInput.value = "";
});
byId("remove").addEventListener("click", () => choose([]));
for (const event of ["dragenter", "dragover", "dragleave", "drop"]) {
  drop.addEventListener(event, e => {
    e.preventDefault();
    drop.classList.toggle("drag", !busy && (event === "dragenter" || event === "dragover"));
    if (event === "drop" && e.dataTransfer.files.length) choose(e.dataTransfer.files);
  });
}
document.addEventListener("dragover", e => {
  if (Array.from(e.dataTransfer.types).includes("Files")) e.preventDefault();
});
document.addEventListener("drop", e => {
  if (e.dataTransfer.files.length) e.preventDefault();
});

async function json(url, options) {
  const response = await fetch(url, options);
  let payload;
  try {
    payload = await response.json();
  } catch {
    throw new Error(response.status === 413 ? `PDF vượt giới hạn ${maxFileMb} MB.` : "Không đọc được phản hồi từ ứng dụng.");
  }
  if (!response.ok) throw new Error(payload.error || "Ứng dụng không thể hoàn tất thao tác.");
  return payload;
}

function progress(info) {
  const bar = byId("progress-bar");
  byId("progress").hidden = false;
  if (info.total) {
    bar.max = info.total;
    bar.value = info.complete;
    byId("progress-value").textContent = `${info.complete}/${info.total}`;
  } else {
    bar.removeAttribute("value");
    byId("progress-value").textContent = "";
  }
  const label = byId("progress-label");
  if (label.textContent !== info.label) label.textContent = info.label;
}

function clearReview() {
  clearInterval(expiryTimer);
  expiryTimer = null;
  expiry.hidden = true;
  previewSequence++;
  review.hidden = true;
  byId("original").replaceChildren();
  byId("rebuild").replaceChildren();
  byId("warnings").replaceChildren();
}

function startExpiryCountdown(remainingMs) {
  if (!Number.isFinite(remainingMs)) return;
  const deadline = Date.now() + remainingMs;
  const update = () => {
    const seconds = Math.max(0, Math.ceil((deadline - Date.now()) / 1000));
    const time = [Math.floor(seconds / 60), seconds % 60].map(value => String(value).padStart(2, "0")).join(":");
    if (expiryTime.textContent !== time) expiryTime.textContent = time;
    expiry.classList.toggle("soon", seconds <= 60);
    if (seconds === 0 && !busy) {
      job = null;
      clearReview();
      byId("progress").hidden = true;
      steps(0);
      message("notice", "Kết quả đã hết hạn. Phiên chuyển đổi đã được giải phóng.");
    }
  };
  expiry.hidden = false;
  expiryTimer = setInterval(update, 250);
  update();
}

convert.addEventListener("click", async () => {
  if (!selected || busy) return;
  clearReview();
  message("error");
  message("notice");
  setBusy(true);
  steps(1);
  progress({label: "Đang đọc tệp PDF…"});
  const language = document.querySelector("[name=lang]:checked").value;
  try {
    const response = await json(`/api/jobs?lang=${encodeURIComponent(language)}`, {
      method: "POST", headers: {...headers, "Content-Type": "application/pdf",
        "X-PDF-Name": encodeURIComponent(selected.name)}, body: selected,
    });
    job = response.id;
    let state;
    do {
      await new Promise(resolve => setTimeout(resolve, 500));
      state = await json(`/api/jobs/${job}`);
      const label = state.state === "queued" ? `Đang chờ xử lý · Vị trí số ${state.queue_position} trong hàng đợi` :
        state.total && state.complete === state.total ? "Đang tạo tệp Word…" :
        state.current ? `Đang xử lý trang ${state.current} / ${state.total}…` : "Đang kiểm tra tệp PDF…";
      progress({...state, label});
    } while (["queued", "working"].includes(state.state));
    if (state.state === "failed") throw new Error(state.error);
    pageSelect.replaceChildren(...state.pages.map((page, index) =>
      new Option(`${index + 1} / ${state.pages.length}${page.warnings ? " · Cần xem" : ""}`, index)));
    steps(2);
    progress({...state, label: "Đã chuyển đổi xong"});
    review.hidden = false;
    download.disabled = false;
    cancel.disabled = false;
    startExpiryCountdown(state.expires_in_ms);
    await showPage();
    byId("review-title").scrollIntoView({block: "nearest"});
  } catch (error) {
    steps(0);
    byId("progress").hidden = true;
    message("error", error.message || "Không kết nối được ứng dụng cục bộ.");
  } finally {
    setBusy(false);
  }
});

function node(tag, className) {
  const child = document.createElement(tag);
  child.className = className;
  return child;
}

function place(child, box) {
  const [left, top, right, bottom] = box;
  Object.assign(child.style, {left: left + "px", top: top + "px",
    width: Math.max(1, right-left) + "px", height: Math.max(1, bottom-top) + "px"});
}

function textLines(paper, lines) {
  for (const line of lines) {
    const lineNode = node("div", "preview-line");
    place(lineNode, line.box);
    if (Math.abs(line.angle) >= 3) lineNode.style.transform = `rotate(${-line.angle}deg)`;
    for (const span of line.spans) {
      const part = node("span", "preview-span");
      part.textContent = span.text;
      Object.assign(part.style, {left: (span.box[0]-line.box[0]) + "px",
        top: (span.baseline-span.size*.92-line.box[1]) + "px", fontSize: span.size + "px",
        fontFamily: `"${span.font.replaceAll('"', '')}", Arial, sans-serif`, color: "#" + span.color,
        fontWeight: span.bold ? "bold" : "normal", fontStyle: span.italic ? "italic" : "normal",
        textDecoration: span.underline ? "underline" : "none",
        transform: `scaleX(${(span.stretch || 100)/100})`});
      lineNode.append(part);
    }
    paper.append(lineNode);
  }
}

function paperFor(id, model) {
  const frame = byId(id);
  const paper = node("div", "paper");
  paper.dataset.width = model.width;
  paper.dataset.height = model.height;
  Object.assign(paper.style, {width: model.width + "px", height: model.height + "px"});
  frame.replaceChildren(paper);
  fit(frame);
  return paper;
}

function fit(frame) {
  const paper = frame.firstElementChild;
  if (!paper || !frame.clientWidth) return;
  const ratio = frame.clientWidth / Number(paper.dataset.width);
  frame.style.height = ratio * Number(paper.dataset.height) + "px";
  paper.style.transform = `scale(${ratio})`;
}

const resizer = new ResizeObserver(entries => entries.forEach(entry => fit(entry.target)));
for (const id of ["original", "rebuild"]) resizer.observe(byId(id));

async function showPage() {
  const sequence = ++previewSequence;
  const url = `/api/jobs/${job}/pages/${pageSelect.value}`;
  try {
    const model = await json(url);
    if (sequence !== previewSequence) return;
    byId("page-mode").textContent = model.mode;
    const original = paperFor("original", model);
    const rebuilt = paperFor("rebuild", model);
    const source = node("img", "original-img");
    source.src = url + "/original";
    source.alt = `PDF gốc, trang ${Number(pageSelect.value)+1}`;
    original.append(source);
    for (const item of model.items.filter(item => item.kind === "image")) {
      const image = node("img", "placed");
      image.src = url + "/assets/" + item.asset;
      image.alt = item.label || "Ảnh từ PDF";
      if (item.preserved_ink) image.style.zIndex = "2";
      place(image, item.box);
      rebuilt.append(image);
    }
    for (const item of model.items) {
      if (item.kind === "table") {
        for (const cell of item.cells) {
          const shape = node("div", "preview-cell");
          place(shape, cell.box);
          shape.style.borderColor = "#" + cell.border;
          shape.style.borderWidth = cell.weight + "px";
          if (cell.fill) shape.style.background = "#" + cell.fill;
          rebuilt.append(shape);
        }
        for (const cell of item.cells) textLines(rebuilt, cell.lines);
      } else if (item.kind === "text") {
        textLines(rebuilt, item.lines);
      }
    }
    const warnings = model.warnings.map(warning => {
      if (warning.box) {
        for (const parent of [original, rebuilt]) {
          const mark = node("div", warning.code === "low" ? "mark low" : "mark");
          place(mark, warning.box);
          mark.setAttribute("aria-hidden", "true");
          parent.append(mark);
        }
      }
      const li = node("li", "");
      li.textContent = warning.text;
      return li;
    });
    byId("warnings").replaceChildren(...warnings);
    byId("warnings").hidden = !warnings.length;
  } catch (error) {
    if (sequence === previewSequence) message("notice", error.message);
  }
}

pageSelect.addEventListener("change", showPage);
byId("show-marks").addEventListener("change", event => {
  review.classList.toggle("marks-off", !event.target.checked);
});

download.addEventListener("click", async () => {
  if (!job || busy) return;
  setBusy(true);
  download.disabled = true;
  cancel.disabled = true;
  message("notice");
  try {
    const response = await fetch(`/api/jobs/${job}/download`, {method: "POST", headers});
    if (!response.ok) {
      const payload = await response.json();
      throw new Error(payload.error);
    }
    const link = document.createElement("a");
    link.href = URL.createObjectURL(await response.blob());
    const fileName = response.headers.get("content-disposition")?.match(/filename\*=UTF-8''([^;]+)/i);
    const fallback = response.headers.get("content-disposition")?.match(/filename="([^"]+)"/i);
    link.download = fileName ? decodeURIComponent(fileName[1]) : fallback?.[1] || "ban-chuyen-doi.docx";
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 4000);
    job = null;
    clearReview();
    message("notice", "Đã gửi tệp Word đến trình duyệt và dọn dữ liệu phiên chuyển đổi.");
  } catch (error) {
    message("notice", error.message || "Không thể tải tệp Word.");
    download.disabled = false;
    cancel.disabled = false;
  } finally {
    setBusy(false);
  }
});

cancel.addEventListener("click", async () => {
  if (!job || busy) return;
  setBusy(true);
  cancel.disabled = true;
  download.disabled = true;
  message("notice");
  try {
    await json(`/api/jobs/${job}/cancel`, {method: "POST", headers});
    job = null;
    clearReview();
    byId("progress").hidden = true;
    steps(0);
    message("notice", "Đã hủy kết quả và giải phóng phiên chuyển đổi.");
  } catch (error) {
    message("notice", error.message || "Không thể hủy phiên chuyển đổi.");
    cancel.disabled = false;
    download.disabled = false;
  } finally {
    setBusy(false);
    if (!job) convert.focus();
  }
});
