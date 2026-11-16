const $ = id => document.getElementById(id);
let leads = [], meta = {keyword: "", location: ""};

const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const none = '<span class="none">–</span>';

fetch("/api/status").then(r => r.json()).then(s => {
  if (!s.google) $("foot").textContent = "Currently using OpenStreetMap only, which has limited data. Add a Google Places key in .env for far more complete leads.";
  $("aiNote").textContent = s.ai ? "(OpenRouter connected)" : "(add OpenRouter key in .env)";
  if (!s.ai) $("ai").disabled = true;
}).catch(() => {});

function visible() {
  return $("onlyContact").checked ? leads.filter(l => l.phone || l.email) : leads;
}

function render() {
  const list = visible();
  $("rows").innerHTML = list.map(l => `
    <tr>
      <td class="name">${esc(l.name)}</td>
      <td>${l.phone ? `<a href="tel:${esc(l.phone)}">${esc(l.phone)}</a>` : none}</td>
      <td>${l.email ? `<a href="mailto:${esc(l.email)}">${esc(l.email)}</a>` : none}</td>
      <td>${l.website ? `<a href="${esc(l.website)}" target="_blank" rel="noopener">${esc(l.website.replace(/^https?:\/\/(www\.)?/, "").replace(/\/$/, ""))}</a>` : none}</td>
      <td>${l.address ? esc(l.address) : none}</td>
    </tr>`).join("");
  const has = list.length > 0;
  $("table").hidden = !has;
  $("download").disabled = !has;
  $("empty").hidden = has || leads.length === 0 ? (leads.length === 0 ? false : true) : true;
  if (leads.length) {
    $("count").innerHTML = `<b>${list.length}</b> lead${list.length === 1 ? "" : "s"}`;
    const e = list.filter(l => l.email).length, p = list.filter(l => l.phone).length;
    $("sub").textContent = `${meta.keyword} in ${meta.location} · ${p} with phone · ${e} with email · Source: ${meta.sources || 'OpenStreetMap'}`;
  }
}

function showMsg(t) { $("msg").hidden = !t; $("msg").textContent = t || ""; }

$("form").addEventListener("submit", async e => {
  e.preventDefault();
  const btn = $("go");
  btn.disabled = true;
  btn.innerHTML = '<span class="spin"></span>Searching…';
  showMsg("");
  $("empty").hidden = true;
  $("count").textContent = "Searching…";
  $("sub").textContent = $("emails").checked ? "Collecting contacts and checking websites for emails. This can take a few minutes." : "Collecting contacts.";
  try {
    const r = await fetch("/api/search", {
      method: "POST", headers: {"Content-Type": "application/json"},
      body: JSON.stringify({
        keyword: $("keyword").value, location: $("location").value,
        limit: +$("limit").value, find_emails: $("emails").checked, use_ai: $("ai").checked, gmaps: $("gmaps").checked
      })
    });
    const d = await r.json();
    if (!r.ok) throw new Error(d.error || "Search failed.");
    leads = d.leads;
    meta.sources = (d.sources || []).join(' + ');
    if (d.warning) showMsg(d.warning);
    meta = {keyword: $("keyword").value.trim(), location: $("location").value.trim(), sources: meta.sources};
    if (!leads.length) {
      $("count").textContent = "No leads found";
      $("sub").textContent = "Try a broader keyword or a larger city.";
      $("empty").hidden = true;
      $("table").hidden = true;
      $("download").disabled = true;
    } else render();
  } catch (err) {
    leads = [];
    $("count").textContent = "No leads";
    $("sub").textContent = "";
    showMsg(err.message);
    $("empty").hidden = false;
    $("table").hidden = true;
    $("download").disabled = true;
  } finally {
    btn.disabled = false;
    btn.textContent = "Find leads";
  }
});

$("onlyContact").addEventListener("change", render);

$("download").addEventListener("click", async () => {
  const r = await fetch("/api/export", {
    method: "POST", headers: {"Content-Type": "application/json"},
    body: JSON.stringify({leads: visible(), ...meta})
  });
  const blob = await r.blob();
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `leads_${meta.keyword}_${meta.location}.xlsx`.replace(/\s+/g, "_").toLowerCase();
  a.click();
  URL.revokeObjectURL(a.href);
});
