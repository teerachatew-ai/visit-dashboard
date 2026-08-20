let DATA = null;
let charts = {};
let currentOwner = null;
const STATUS_LABEL = {
  overdue: "🔴 เกินกำหนดเข้าเยี่ยม",
  done: "✅ ครบตามเป้าหมาย",
  in_progress: "",
};
const STATUS_FILTER_LABEL = {
  overdue: "🔴 เกินกำหนดเข้าเยี่ยม",
  done: "✅ ครบตามเป้าหมาย",
  in_progress: "อยู่ระหว่างดำเนินการ",
};
const STATUS_ORDER = { overdue: 0, in_progress: 1, done: 2 };
const filters = {
  tg_status: "",
  tg_ie: "",
  tg_sl: "",
};
const CHART_COLORS = ["#2563eb", "#16a34a", "#d97706", "#dc2626", "#8b5cf6", "#6b7280"];

function fmt(n) {
  return Number(n).toLocaleString("th-TH");
}

async function loadData(force = false) {
  document.getElementById("loadingIndicator").style.display = "flex";
  if (force) await fetch("/api/refresh", { method: "POST" });
  const res = await fetch("/api/data");
  DATA = await res.json();
  document.getElementById("loadingIndicator").style.display = "none";
  document.getElementById("lastUpdated").textContent = "อัปเดตล่าสุด: " + DATA.generated_at;
  initFilters();
  renderTarget();
  renderSL();
  renderInsights();
}

function uniqueSorted(arr) {
  return Array.from(new Set(arr.filter((v) => v))).sort();
}

function renderTabGroup(containerId, options, allLabel, currentValue, onChange, labelFn, chipClass = "tab-chip") {
  const el = document.getElementById(containerId);
  if (!el) return;
  const values = ["", ...options];
  el.innerHTML = values
    .map((v) => {
      const label = v === "" ? allLabel : (labelFn ? labelFn(v) : v);
      return `<button type="button" class="${chipClass}${v === currentValue ? " active" : ""}" data-value="${v}">${label}</button>`;
    })
    .join("");
  el.querySelectorAll(`.${chipClass}`).forEach((btn) => {
    btn.addEventListener("click", () => {
      el.querySelectorAll(`.${chipClass}`).forEach((b) => b.classList.toggle("active", b === btn));
      onChange(btn.dataset.value);
    });
  });
}

function initFilters() {
  const ies = uniqueSorted(DATA.customers.map((c) => c.ie));
  const sls = uniqueSorted(DATA.customers.map((c) => c.sl_real));

  renderTabGroup("tg-status-tabs", ["overdue", "done", "in_progress"], "ทั้งหมด", filters.tg_status, (v) => { filters.tg_status = v; renderTarget(); }, (v) => STATUS_FILTER_LABEL[v]);
  renderTabGroup("tg-ie-tabs", ies, "ทุกพื้นที่", filters.tg_ie, (v) => { filters.tg_ie = v; renderTarget(); });
  renderTabGroup("tg-sl-tabs", sls, "ทุก Owner", filters.tg_sl, (v) => { filters.tg_sl = v; renderTarget(); }, null, "owner-tab");

  renderOwnerTabs(sls);

  document.getElementById("tg-search").addEventListener("input", renderTarget);
  document.getElementById("search-box").addEventListener("input", renderSearch);
  document.getElementById("tg-export").addEventListener("click", exportTargetCSV);
}

function destroyChart(key) {
  if (charts[key]) charts[key].destroy();
}

function makeChart(key, ctx, config) {
  destroyChart(key);
  charts[key] = new Chart(ctx, config);
}

function kpiCardsHtml(kpis) {
  return kpis
    .map((k) => `<div class="kpi-card"><div class="kpi-value">${k.value}</div><div class="kpi-label">${k.label}</div></div>`)
    .join("");
}

function bindClickableRows(tbodyId) {
  document.querySelectorAll(`#${tbodyId} tr[data-customer-id]`).forEach((tr) => {
    tr.addEventListener("click", () => openCustomerDetail(tr.dataset.customerId));
  });
}

function openCustomerDetail(customerId) {
  switchTab("customers");
  document.querySelectorAll(".sub-tab-btn").forEach((b) => b.classList.toggle("active", b.dataset.subtab === "search"));
  document.querySelectorAll(".sub-tab-panel").forEach((p) => p.classList.remove("active"));
  document.getElementById("subtab-search").classList.add("active");
  document.getElementById("search-box").value = "";
  document.getElementById("search-results").innerHTML = "";
  showCompanyDetail(customerId);
  document.getElementById("company-detail").scrollIntoView({ behavior: "smooth", block: "start" });
}

function progressCellHtml(c) {
  const pct = c.target_per_year ? Math.min(100, Math.round((c.visits_this_year / c.target_per_year) * 100)) : 0;
  const full = c.visits_this_year >= c.target_per_year;
  return `<div class="progress-cell">
    <div class="progress-track"><div class="progress-fill ${full ? "full" : ""}" style="width:${pct}%"></div></div>
    <div class="progress-text">${c.visits_this_year}/${c.target_per_year}</div>
  </div>`;
}

// ---------- TARGET STATUS ----------
function filteredTargetCustomers(includeStatus = true) {
  const search = document.getElementById("tg-search").value.trim().toLowerCase();
  const status = filters.tg_status;
  const ie = filters.tg_ie;
  const sl = filters.tg_sl;
  return DATA.customers.filter((c) => {
    if (search && !c.company_name.toLowerCase().includes(search)) return false;
    if (includeStatus && status && c.status !== status) return false;
    if (ie && c.ie !== ie) return false;
    if (sl && c.sl_real !== sl) return false;
    return true;
  });
}

function renderTarget() {
  const list = filteredTargetCustomers().slice().sort((a, b) => {
    return (STATUS_ORDER[a.status] - STATUS_ORDER[b.status]) || (b.visits_needed - a.visits_needed);
  });

  // Counts reflect the current search/พื้นที่/Owner filter (but not the status
  // tab itself), so switching status tabs doesn't zero out the other counts.
  const scoped = filteredTargetCustomers(false);
  const counts = { overdue: 0, in_progress: 0, done: 0 };
  scoped.forEach((c) => counts[c.status]++);
  document.getElementById("tg-summary").innerHTML = `
    <span class="status-chip">👥 ลูกค้าทั้งหมด: <b>${scoped.length}</b></span>
    <span class="status-chip">🔴 เกินกำหนดเข้าเยี่ยม: <b>${counts.overdue}</b></span>
    <span class="status-chip">✅ ครบตามเป้าหมาย: <b>${counts.done}</b></span>
    <span class="status-chip">อยู่ระหว่างดำเนินการ: <b>${counts.in_progress}</b></span>
  `;

  document.getElementById("tg-tbody").innerHTML = list.length
    ? list.map((c) => `<tr class="clickable-row" data-customer-id="${c.customer_id}">
        <td><b>${c.company_name}</b></td>
        <td>${c.ie}</td>
        <td>${c.sl_real}</td>
        <td>${progressCellHtml(c)}</td>
        <td>${c.visits_needed > 0 ? c.visits_needed + " ครั้ง" : "-"}</td>
        <td>${c.last_visit_date || "ยังไม่เคย"}</td>
        <td>${STATUS_LABEL[c.status] ? `<span class="badge badge-${c.status}">${STATUS_LABEL[c.status]}</span>` : ""}</td>
      </tr>`).join("")
    : `<tr><td colspan="7"><div class="search-empty">ไม่พบลูกค้าตามเงื่อนไขที่เลือก</div></td></tr>`;
  bindClickableRows("tg-tbody");

  document.getElementById("offTargetCount").textContent = DATA.off_target_customers.length;
  document.getElementById("off-tbody").innerHTML = DATA.off_target_customers
    .map((c) => `<tr class="clickable-row" data-customer-id="${c.customer_id}"><td>${c.company_name}</td><td>${c.ie}</td><td>${c.visits_total}</td><td>${c.visits_this_year}</td><td>${c.last_visit_date}</td></tr>`)
    .join("");
  bindClickableRows("off-tbody");
}

function exportTargetCSV() {
  const list = filteredTargetCustomers();
  const header = ["บริษัท", "พื้นที่", "Owner", "Target/ปี", "เยี่ยมแล้ว", "ขาดอีก", "ติดต่อล่าสุด", "สถานะ"];
  const rows = list.map((c) => [c.company_name, c.ie, c.sl_real, c.target_per_year, c.visits_this_year, c.visits_needed, c.last_visit_date || "", STATUS_LABEL[c.status]]);
  const csv = [header, ...rows].map((r) => r.map((x) => `"${String(x).replace(/"/g, '""')}"`).join(",")).join("\n");
  const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "visit_target_status.csv";
  a.click();
}

// ---------- PER SL ----------
function renderOwnerTabs(sls) {
  if (currentOwner === null || !sls.includes(currentOwner)) {
    currentOwner = sls[0] || null;
  }
  const wrap = document.getElementById("sl-owner-tabs");
  wrap.innerHTML = sls
    .map((s) => `<button type="button" class="owner-tab${s === currentOwner ? " active" : ""}" data-owner="${s}">${s}</button>`)
    .join("");
  wrap.querySelectorAll(".owner-tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      currentOwner = btn.dataset.owner;
      wrap.querySelectorAll(".owner-tab").forEach((b) => b.classList.toggle("active", b === btn));
      renderSL();
    });
  });
}

function renderSL() {
  const sl = currentOwner;
  const custOfSL = DATA.customers.filter((c) => c.sl_real === sl);
  const custIdsOfSL = new Set(custOfSL.map((c) => c.customer_id));
  const visitsOfSL = DATA.visits.filter((v) => custIdsOfSL.has(v.customer_id));

  if (!sl) {
    document.getElementById("sl-kpis").innerHTML = `<div class="search-empty" style="grid-column:1/-1">เลือก Owner ด้านบนเพื่อดูสรุปผลงาน</div>`;
    document.getElementById("sl-cust-tbody").innerHTML = "";
    destroyChart("slTrend");
    destroyChart("slTopic");
  } else {
    const year = String(new Date().getFullYear());
    const visitsThisYear = visitsOfSL.filter((v) => v.date.startsWith(year));
    const doneCount = custOfSL.filter((c) => c.status === "done").length;
    const pct = custOfSL.length ? ((doneCount / custOfSL.length) * 100).toFixed(1) : 0;
    const docsOfSL = DATA.documents.filter((d) => custIdsOfSL.has(d.customer_id) && d.date.startsWith(year));

    document.getElementById("sl-kpis").innerHTML = kpiCardsHtml([
      { label: "ลูกค้าในความดูแล", value: fmt(custOfSL.length) },
      { label: "เยี่ยมปีนี้", value: fmt(visitsThisYear.length) },
      { label: "% ครบเป้าหมาย", value: pct + "%" },
      { label: "เอกสารที่รับปีนี้", value: fmt(docsOfSL.length) },
    ]);

    document.getElementById("sl-cust-tbody").innerHTML = custOfSL
      .slice()
      .sort((a, b) => STATUS_ORDER[a.status] - STATUS_ORDER[b.status])
      .map((c) => `<tr class="clickable-row" data-customer-id="${c.customer_id}">
        <td><b>${c.company_name}</b></td>
        <td>${c.ie}</td>
        <td>${progressCellHtml(c)}</td>
        <td>${c.visits_needed > 0 ? c.visits_needed + " ครั้ง" : "-"}</td>
        <td>${c.last_visit_date || "ยังไม่เคย"}</td>
        <td><span class="badge badge-${c.status}">${STATUS_LABEL[c.status]}</span></td>
      </tr>`)
      .join("");
    bindClickableRows("sl-cust-tbody");

    const months = Array.from({ length: 12 }, (_, i) => `${year}-${String(i + 1).padStart(2, "0")}`);
    const slCounts = months.map((m) => visitsThisYear.filter((v) => v.month === m).length);
    const teamAvg = months.map((m) => {
      const total = DATA.visits.filter((v) => v.month === m).length;
      const n = DATA.sl_list.length || 1;
      return total / n;
    });
    makeChart("slTrend", document.getElementById("chartSLTrend"), {
      type: "line",
      data: {
        labels: months,
        datasets: [
          { label: sl, data: slCounts, borderColor: "#2563eb", tension: 0.3 },
          { label: "ค่าเฉลี่ยทีม", data: teamAvg, borderColor: "#9aa1ac", borderDash: [5, 5], tension: 0.3 },
        ],
      },
      options: { responsive: true },
    });

    const topicCounts = {};
    visitsThisYear.forEach((v) => { topicCounts[v.topic_group] = (topicCounts[v.topic_group] || 0) + 1; });
    const topicLabels = Object.keys(topicCounts);
    makeChart("slTopic", document.getElementById("chartSLTopic"), {
      type: "doughnut",
      data: { labels: topicLabels, datasets: [{ data: topicLabels.map((l) => topicCounts[l]), backgroundColor: CHART_COLORS }] },
    });
  }

  const leaderRows = DATA.sl_list.map((s) => {
    const cust = DATA.customers.filter((c) => c.sl_real === s);
    const cids = new Set(cust.map((c) => c.customer_id));
    const year = String(new Date().getFullYear());
    const v = DATA.visits.filter((vv) => cids.has(vv.customer_id) && vv.date.startsWith(year)).length;
    const done = cust.filter((c) => c.status === "done").length;
    const pct = cust.length ? ((done / cust.length) * 100).toFixed(1) : "-";
    return { s, custCount: cust.length, v, done, pct };
  }).sort((a, b) => b.custCount - a.custCount);

  document.getElementById("sl-leader-tbody").innerHTML = leaderRows
    .map((r) => `<tr><td><b>${r.s}</b></td><td>${r.custCount}</td><td>${r.v}</td><td>${r.done}</td><td>${r.pct}%</td></tr>`)
    .join("");

  renderStaffManage();
}

// ---------- STAFF MANAGEMENT ----------
async function setStaffResigned(name, resigned) {
  await fetch("/api/staff/set_status", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, resigned }),
  });
  await loadData(true);
  switchTab("team");
}

function renderStaffManage() {
  const resignedSet = new Set(DATA.resigned_staff || []);
  const active = (DATA.all_staff || []).filter((s) => !resignedSet.has(s));
  const resigned = DATA.resigned_staff || [];

  const activeEl = document.getElementById("staff-active-list");
  const resignedEl = document.getElementById("staff-resigned-list");
  if (!activeEl || !resignedEl) return;

  activeEl.innerHTML = active.length
    ? active.map((s) => `<span class="staff-chip" data-name="${s}">${s}<button title="ทำเครื่องหมายว่าลาออกแล้ว" data-action="resign" data-name="${s}">✕</button></span>`).join("")
    : `<span class="staff-empty">ไม่มีรายชื่อ</span>`;

  resignedEl.innerHTML = resigned.length
    ? resigned.map((s) => `<span class="staff-chip is-resigned" data-name="${s}">${s}<button title="กู้คืนกลับเป็นทีมปัจจุบัน" data-action="restore" data-name="${s}">↺</button></span>`).join("")
    : `<span class="staff-empty">ยังไม่มีใครลาออก</span>`;

  document.querySelectorAll('[data-action="resign"]').forEach((btn) => {
    btn.addEventListener("click", () => {
      if (confirm(`ทำเครื่องหมายว่า "${btn.dataset.name}" ลาออกแล้วใช่ไหม? (ประวัติเก่ายังอยู่ครบ แค่จะไม่ขึ้นในรายการเลือก)`)) {
        setStaffResigned(btn.dataset.name, true);
      }
    });
  });
  document.querySelectorAll('[data-action="restore"]').forEach((btn) => {
    btn.addEventListener("click", () => setStaffResigned(btn.dataset.name, false));
  });
}

async function importOwnerTransferFile() {
  const fileInput = document.getElementById("owner-import-file");
  const resultEl = document.getElementById("owner-import-result");
  const file = fileInput.files[0];
  if (!file) {
    resultEl.innerHTML = `<div class="import-err">กรุณาเลือกไฟล์ Excel ก่อน</div>`;
    return;
  }
  const formData = new FormData();
  formData.append("file", file);
  resultEl.innerHTML = `<div>กำลังนำเข้า...</div>`;
  try {
    const res = await fetch("/api/owners/import", { method: "POST", body: formData });
    const result = await res.json();
    if (!result.ok) {
      resultEl.innerHTML = `<div class="import-err">${result.error || "นำเข้าไม่สำเร็จ"}</div>`;
      return;
    }
    let html = `<div class="import-ok">อัปเดตผู้ดูแลแล้ว ${result.applied_count} ราย${result.skipped_count ? ` (ข้าม ${result.skipped_count} ราย)` : ""}</div>`;
    if (result.applied.length) {
      html += `<table><thead><tr><th>บริษัท</th><th>เดิม</th><th>ใหม่</th></tr></thead><tbody>`;
      html += result.applied.map((r) => `<tr><td>${r.company_name}</td><td>${r.old_owner}</td><td><b>${r.new_owner}</b></td></tr>`).join("");
      html += `</tbody></table>`;
    }
    if (result.skipped.length) {
      html += `<div class="import-err" style="margin-top:8px;">ข้ามรายการที่ไม่พบรหัสลูกค้า: ${result.skipped.map((s) => s.customer_id).join(", ")}</div>`;
    }
    resultEl.innerHTML = html;
    fileInput.value = "";
    await loadData(true);
    switchTab("team");
  } catch (e) {
    resultEl.innerHTML = `<div class="import-err">เกิดข้อผิดพลาด: ${e.message}</div>`;
  }
}

// ---------- SEARCH ----------
function renderSearch() {
  const q = document.getElementById("search-box").value.trim().toLowerCase();
  const resultsEl = document.getElementById("search-results");
  const detailEl = document.getElementById("company-detail");
  if (!q) {
    resultsEl.innerHTML = "";
    detailEl.style.display = "none";
    return;
  }
  const allCompanies = new Map();
  DATA.customers.forEach((c) => allCompanies.set(c.customer_id, { customer_id: c.customer_id, company_name: c.company_name, ie: c.ie }));
  DATA.off_target_customers.forEach((c) => {
    if (!allCompanies.has(c.customer_id)) allCompanies.set(c.customer_id, { customer_id: c.customer_id, company_name: c.company_name, ie: c.ie });
  });

  const matches = Array.from(allCompanies.values()).filter(
    (c) => c.company_name.toLowerCase().includes(q) || c.customer_id.toLowerCase().includes(q)
  ).slice(0, 20);

  resultsEl.innerHTML = matches.length
    ? matches.map((c) => `<div class="search-result-item" data-id="${c.customer_id}">
        <span class="rname">${c.company_name}</span>
        <span class="muted">${c.customer_id} · ${c.ie}</span>
      </div>`).join("")
    : `<div class="search-empty">ไม่พบบริษัทที่ตรงกับคำค้นหา "${q}"</div>`;

  resultsEl.querySelectorAll(".search-result-item").forEach((el) => {
    el.addEventListener("click", () => showCompanyDetail(el.dataset.id));
  });

  if (matches.length === 1) showCompanyDetail(matches[0].customer_id);
  else detailEl.style.display = "none";
}

function showCompanyDetail(customerId) {
  const detailEl = document.getElementById("company-detail");
  const cust = DATA.customers.find((c) => c.customer_id === customerId);
  const offCust = DATA.off_target_customers.find((c) => c.customer_id === customerId);
  const visits = DATA.visits.filter((v) => v.customer_id === customerId);
  const docs = DATA.documents.filter((d) => d.customer_id === customerId);

  const name = cust ? cust.company_name : offCust ? offCust.company_name : customerId;
  const ie = cust ? cust.ie : offCust ? offCust.ie : "";

  let chips = [
    `<div class="meta-chip">รหัสลูกค้า <b>${customerId}</b></div>`,
    `<div class="meta-chip">พื้นที่ <b>${ie}</b></div>`,
    `<div class="meta-chip">เยี่ยมทั้งหมด <b>${visits.length}</b> ครั้ง</div>`,
    `<div class="meta-chip">รับเอกสาร <b>${docs.length}</b> ครั้ง</div>`,
  ];
  if (cust) {
    chips.push(
      `<div class="meta-chip">Owner ผู้ดูแล <b>${cust.sl_real}</b></div>`,
      `<div class="meta-chip">เป้าหมายปีนี้ <b>${cust.visits_this_year}/${cust.target_per_year}</b></div>`,
      `<div class="meta-chip">${STATUS_LABEL[cust.status] ? `<span class="badge badge-${cust.status}">${STATUS_LABEL[cust.status]}</span>` : "อยู่ระหว่างดำเนินการ"}</div>`
    );
    if (cust.csi2025 && cust.csi2025 !== "-") chips.push(`<div class="meta-chip">CSI 2025 <b>${cust.csi2025}</b></div>`);
  } else {
    chips.push(`<div class="meta-chip"><span class="badge badge-never">นอกแผนเป้าหมาย</span></div>`);
  }

  const items = [];
  visits.forEach((v) => items.push({ type: "visit", date: v.date, ...v }));
  docs.forEach((d) => items.push({ type: "doc", date: d.date, ...d }));
  items.sort((a, b) => (a.date < b.date ? 1 : -1));

  const timelineHtml = items
    .map((it) => {
      if (it.type === "visit") {
        return `<div class="timeline-item">
          <div class="tl-date">${it.date} — เยี่ยมโดย ${it.staff}</div>
          <div class="tl-title">${it.topic || "(ไม่ระบุหัวข้อ)"}</div>
          <div class="tl-detail">${it.details || ""}${it.contact ? " · ผู้ติดต่อ: " + it.contact : ""}</div>
        </div>`;
      }
      return `<div class="timeline-item doc">
        <div class="tl-date">${it.date} — รับเอกสารโดย ${it.staff}</div>
        <div class="tl-title">📄 ${it.doc1 || it.doc2 || "เอกสาร"}</div>
        <div class="tl-detail">${it.note || ""}${it.receiver ? " · ผู้รับ: " + it.receiver : ""}</div>
      </div>`;
    })
    .join("");

  detailEl.innerHTML = `<h2>${name}</h2>
    <div class="meta-row">${chips.join("")}</div>
    <div class="timeline">${timelineHtml || "<i>ไม่มีประวัติ</i>"}</div>`;
  detailEl.style.display = "block";
}

// ---------- INSIGHTS ----------
function renderHeatmap(elId) {
  const el = document.getElementById(elId);
  if (!el) return;
  const year = String(new Date().getFullYear());
  const visitsThisYear = DATA.visits.filter((v) => v.date.startsWith(year));
  const months = Array.from({ length: 12 }, (_, i) => `${year}-${String(i + 1).padStart(2, "0")}`);
  const sls = DATA.sl_list;
  const custIdsBySL = {};
  sls.forEach((s) => {
    custIdsBySL[s] = new Set(DATA.customers.filter((c) => c.sl_real === s).map((c) => c.customer_id));
  });
  let heatHtml = '<table class="heat-table"><thead><tr><th>เดือน</th>' + sls.map((s) => `<th>${s}</th>`).join("") + "</tr></thead><tbody>";
  months.forEach((m) => {
    heatHtml += `<tr><td>${m}</td>`;
    sls.forEach((s) => {
      const cnt = visitsThisYear.filter((v) => v.month === m && custIdsBySL[s].has(v.customer_id)).length;
      const alpha = Math.min(cnt / 10, 1);
      heatHtml += `<td style="background:rgba(37,99,235,${alpha.toFixed(2)})">${cnt || ""}</td>`;
    });
    heatHtml += "</tr>";
  });
  heatHtml += "</tbody></table>";
  el.innerHTML = heatHtml;
}

function renderInsights() {
  renderHeatmap("heatmap");

  const now = new Date();
  const dayOfYear = Math.floor((now - new Date(now.getFullYear(), 0, 0)) / 86400000);
  const yearFraction = dayOfYear / 365;
  let projectedShort = 0;
  DATA.customers.forEach((c) => {
    const projectedByYearEnd = yearFraction > 0 ? c.visits_this_year / yearFraction : 0;
    if (projectedByYearEnd < c.target_per_year) projectedShort++;
  });
  document.getElementById("forecastBox").innerHTML = `
    ถ้ารักษาจังหวะการติดต่อลูกค้าแบบปัจจุบันไปตลอดทั้งปี (ตอนนี้ผ่านไปแล้ว ${(yearFraction * 100).toFixed(0)}% ของปี)<br>
    คาดว่าจะมีลูกค้า <b>${fmt(projectedShort)}</b> จากทั้งหมด ${fmt(DATA.customers.length)} ราย ที่จะ<b>ไม่ครบเป้าหมาย</b>ภายในสิ้นปีนี้
  `;
}

function switchTab(tabName) {
  document.querySelectorAll(".nav-btn").forEach((b) => {
    b.classList.toggle("active", b.dataset.tab === tabName);
  });
  document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
  document.getElementById("tab-" + tabName).classList.add("active");
  window.scrollTo({ top: 0, behavior: "instant" });
}

document.querySelectorAll(".nav-btn").forEach((btn) => {
  btn.addEventListener("click", () => switchTab(btn.dataset.tab));
});

document.querySelectorAll(".sub-tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".sub-tab-btn").forEach((b) => b.classList.toggle("active", b === btn));
    document.querySelectorAll(".sub-tab-panel").forEach((p) => p.classList.remove("active"));
    document.getElementById("subtab-" + btn.dataset.subtab).classList.add("active");
  });
});

document.getElementById("refreshBtn").addEventListener("click", () => loadData(true));
document.getElementById("owner-import-btn").addEventListener("click", importOwnerTransferFile);

loadData();
