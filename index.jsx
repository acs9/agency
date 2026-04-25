import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Utilities ────────────────────────────────────────────────────────────────
const today = () => new Date().toISOString().split("T")[0];
const daysBetween = (dateStr) => {
  const diff = new Date(dateStr) - new Date(today());
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
};
const fmt = (dateStr) =>
  dateStr
    ? new Date(dateStr).toLocaleDateString("en-IN", {
        day: "2-digit",
        month: "short",
        year: "numeric",
      })
    : "—";
const currency = (n) =>
  new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 }).format(n);
const uid = () => Math.random().toString(36).slice(2, 9);

const urgencyColor = (days) => {
  if (days < 0) return "#ef4444";
  if (days <= 14) return "#f97316";
  if (days <= 30) return "#eab308";
  return "#22c55e";
};
const urgencyBg = (days) => {
  if (days < 0) return "rgba(239,68,68,0.12)";
  if (days <= 14) return "rgba(249,115,22,0.12)";
  if (days <= 30) return "rgba(234,179,8,0.12)";
  return "rgba(34,197,94,0.08)";
};

// ─── Seed Data ────────────────────────────────────────────────────────────────
const SEED_CLIENTS = [
  {
    id: "c1",
    name: "Zenith Retail Pvt Ltd",
    contact: "Arjun Mehta",
    email: "arjun@zenithretail.in",
    phone: "+91 98765 43210",
    domains: [
      { id: "d1", domain: "zenithretail.in", registrar: "GoDaddy", renewal: "2025-04-10", notes: "Primary domain" },
      { id: "d2", domain: "zenithshop.com", registrar: "Namecheap", renewal: "2025-08-22", notes: "Redirect" },
    ],
    hosting: [
      { id: "h1", provider: "Cloudways", plan: "Pro — 4GB", renewal: "2025-03-05", monthly: 4500, notes: "Main server" },
    ],
    invoices: [],
    tags: ["ecommerce", "active"],
  },
  {
    id: "c2",
    name: "Prakash Legal Associates",
    contact: "Sunita Prakash",
    email: "sunita@prakashlegal.com",
    phone: "+91 91234 56789",
    domains: [
      { id: "d3", domain: "prakashlegal.com", registrar: "BigRock", renewal: "2025-12-15", notes: "" },
    ],
    hosting: [
      { id: "h2", provider: "WP Engine", plan: "Startup", renewal: "2025-07-01", monthly: 2800, notes: "" },
    ],
    invoices: [],
    tags: ["professional", "active"],
  },
  {
    id: "c3",
    name: "BlueSky Travels",
    contact: "Rohan Kapoor",
    email: "rohan@blueskytravels.co",
    phone: "+91 99887 66554",
    domains: [
      { id: "d4", domain: "blueskytravels.co", registrar: "Namecheap", renewal: "2025-02-28", notes: "URGENT" },
    ],
    hosting: [
      { id: "h3", provider: "Hostinger", plan: "Business", renewal: "2025-02-20", monthly: 1200, notes: "Shared hosting" },
    ],
    invoices: [],
    tags: ["travel", "at-risk"],
  },
];

// ─── Invoice Template ─────────────────────────────────────────────────────────
const generateInvoiceHTML = (invoice, client) => `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Invoice #${invoice.number}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:'DM Sans',sans-serif;color:#1a1a2e;background:#fff;padding:48px}
  .header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:48px;padding-bottom:24px;border-bottom:2px solid #1a1a2e}
  .agency-name{font-family:'DM Serif Display',serif;font-size:28px;color:#1a1a2e}
  .agency-sub{font-size:12px;color:#666;margin-top:4px;letter-spacing:0.05em;text-transform:uppercase}
  .inv-badge{background:#1a1a2e;color:#fff;padding:6px 14px;border-radius:4px;font-size:11px;font-weight:600;letter-spacing:0.08em;text-transform:uppercase}
  .inv-number{font-family:'DM Serif Display',serif;font-size:22px;margin-top:8px;text-align:right}
  .meta{display:grid;grid-template-columns:1fr 1fr;gap:32px;margin-bottom:40px}
  .meta-block label{font-size:10px;text-transform:uppercase;letter-spacing:0.1em;color:#888;font-weight:600;display:block;margin-bottom:6px}
  .meta-block .val{font-size:15px;font-weight:500}
  .meta-block .sub{font-size:13px;color:#555;margin-top:2px}
  table{width:100%;border-collapse:collapse;margin-bottom:32px}
  th{background:#1a1a2e;color:#fff;padding:10px 14px;font-size:11px;text-transform:uppercase;letter-spacing:0.07em;font-weight:600;text-align:left}
  td{padding:12px 14px;border-bottom:1px solid #eee;font-size:14px}
  tr:last-child td{border-bottom:none}
  tr:nth-child(even) td{background:#f9f9fc}
  .totals{margin-left:auto;width:280px}
  .totals tr td:first-child{color:#666;font-size:13px}
  .totals tr td:last-child{text-align:right;font-weight:500}
  .totals .grand td{font-size:17px;font-weight:700;border-top:2px solid #1a1a2e;padding-top:14px;color:#1a1a2e}
  .notes-box{background:#f5f5fa;border-left:3px solid #1a1a2e;padding:16px 20px;margin-top:32px;font-size:13px;color:#444;border-radius:0 6px 6px 0}
  .footer{margin-top:48px;padding-top:20px;border-top:1px solid #ddd;font-size:11px;color:#999;text-align:center;letter-spacing:0.03em}
</style>
</head>
<body>
<div class="header">
  <div>
    <div class="agency-name">Your Agency</div>
    <div class="agency-sub">Digital Marketing & Web Services</div>
  </div>
  <div style="text-align:right">
    <div class="inv-badge">Invoice</div>
    <div class="inv-number">#${invoice.number}</div>
  </div>
</div>

<div class="meta">
  <div class="meta-block">
    <label>Bill To</label>
    <div class="val">${client.name}</div>
    <div class="sub">${client.contact}</div>
    <div class="sub">${client.email}</div>
    <div class="sub">${client.phone}</div>
  </div>
  <div class="meta-block" style="text-align:right">
    <label>Issue Date</label>
    <div class="val">${fmt(invoice.issueDate)}</div>
    <label style="margin-top:16px">Due Date</label>
    <div class="val">${fmt(invoice.dueDate)}</div>
    <label style="margin-top:16px">Status</label>
    <div class="val" style="color:${invoice.status==='paid'?'#22c55e':invoice.status==='overdue'?'#ef4444':'#f97316'}">${invoice.status.toUpperCase()}</div>
  </div>
</div>

<table>
  <thead>
    <tr><th>Description</th><th>Qty</th><th>Rate (₹)</th><th>Amount (₹)</th></tr>
  </thead>
  <tbody>
    ${invoice.items.map(i=>`<tr>
      <td>${i.desc}</td>
      <td>${i.qty}</td>
      <td>${Number(i.rate).toLocaleString('en-IN')}</td>
      <td>${(i.qty*i.rate).toLocaleString('en-IN')}</td>
    </tr>`).join('')}
  </tbody>
</table>

<table class="totals">
  <tr><td>Subtotal</td><td>₹${invoice.items.reduce((s,i)=>s+i.qty*i.rate,0).toLocaleString('en-IN')}</td></tr>
  <tr><td>GST (18%)</td><td>₹${Math.round(invoice.items.reduce((s,i)=>s+i.qty*i.rate,0)*0.18).toLocaleString('en-IN')}</td></tr>
  <tr class="grand"><td>Total Due</td><td>₹${Math.round(invoice.items.reduce((s,i)=>s+i.qty*i.rate,0)*1.18).toLocaleString('en-IN')}</td></tr>
</table>

${invoice.notes?`<div class="notes-box"><strong>Notes:</strong> ${invoice.notes}</div>`:''}

<div class="footer">Thank you for your business · Payment due within 15 days · GSTIN: 27XXXXX0000X1ZX</div>
</body>
</html>`;

// ─── XLS Export ───────────────────────────────────────────────────────────────
const exportToXLS = (clients) => {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Clients
  const clientRows = clients.map(c => ({
    "Client ID": c.id,
    "Company Name": c.name,
    "Contact Person": c.contact,
    "Email": c.email,
    "Phone": c.phone,
    "Tags": c.tags.join(", "),
  }));
  const wsClients = XLSX.utils.json_to_sheet(clientRows.length ? clientRows : [{ "Client ID": "", "Company Name": "", "Contact Person": "", "Email": "", "Phone": "", "Tags": "" }]);
  wsClients["!cols"] = [12, 28, 22, 30, 18, 20].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsClients, "Clients");

  // Sheet 2: Domains
  const domainRows = clients.flatMap(c =>
    c.domains.map(d => ({
      "Client ID": c.id,
      "Client Name": c.name,
      "Domain": d.domain,
      "Registrar": d.registrar,
      "Renewal Date": d.renewal,
      "Notes": d.notes,
    }))
  );
  const wsDomains = XLSX.utils.json_to_sheet(domainRows.length ? domainRows : [{ "Client ID": "", "Client Name": "", "Domain": "", "Registrar": "", "Renewal Date": "", "Notes": "" }]);
  wsDomains["!cols"] = [12, 28, 25, 18, 15, 30].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsDomains, "Domains");

  // Sheet 3: Hosting
  const hostingRows = clients.flatMap(c =>
    c.hosting.map(h => ({
      "Client ID": c.id,
      "Client Name": c.name,
      "Provider": h.provider,
      "Plan": h.plan,
      "Monthly Cost (INR)": h.monthly,
      "Renewal Date": h.renewal,
      "Notes": h.notes,
    }))
  );
  const wsHosting = XLSX.utils.json_to_sheet(hostingRows.length ? hostingRows : [{ "Client ID": "", "Client Name": "", "Provider": "", "Plan": "", "Monthly Cost (INR)": "", "Renewal Date": "", "Notes": "" }]);
  wsHosting["!cols"] = [12, 28, 18, 18, 18, 15, 30].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsHosting, "Hosting");

  // Sheet 4: Invoices
  const invoiceRows = clients.flatMap(c =>
    c.invoices.map(inv => {
      const subtotal = inv.items.reduce((s, i) => s + i.qty * i.rate, 0);
      return {
        "Client ID": c.id,
        "Client Name": c.name,
        "Invoice Number": inv.number,
        "Issue Date": inv.issueDate,
        "Due Date": inv.dueDate,
        "Status": inv.status,
        "Subtotal (INR)": subtotal,
        "GST 18% (INR)": Math.round(subtotal * 0.18),
        "Total (INR)": Math.round(subtotal * 1.18),
        "Notes": inv.notes,
        "Line Items (JSON)": JSON.stringify(inv.items),
      };
    })
  );
  const wsInvoices = XLSX.utils.json_to_sheet(invoiceRows.length ? invoiceRows : [{ "Client ID": "", "Client Name": "", "Invoice Number": "", "Issue Date": "", "Due Date": "", "Status": "", "Subtotal (INR)": "", "GST 18% (INR)": "", "Total (INR)": "", "Notes": "", "Line Items (JSON)": "" }]);
  wsInvoices["!cols"] = [12, 28, 16, 14, 14, 10, 16, 14, 14, 25, 40].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsInvoices, "Invoices");

  // Sheet 5: Renewals Summary
  const allRenewals = clients.flatMap(c => [
    ...c.domains.map(d => ({ "Client": c.name, "Type": "Domain", "Name": d.domain, "Renewal Date": d.renewal, "Days Left": daysBetween(d.renewal), "Status": daysBetween(d.renewal) < 0 ? "EXPIRED" : daysBetween(d.renewal) <= 14 ? "URGENT" : daysBetween(d.renewal) <= 30 ? "DUE SOON" : "OK" })),
    ...c.hosting.map(h => ({ "Client": c.name, "Type": "Hosting", "Name": h.provider + " — " + h.plan, "Renewal Date": h.renewal, "Days Left": daysBetween(h.renewal), "Status": daysBetween(h.renewal) < 0 ? "EXPIRED" : daysBetween(h.renewal) <= 14 ? "URGENT" : daysBetween(h.renewal) <= 30 ? "DUE SOON" : "OK" })),
  ]).sort((a, b) => a["Days Left"] - b["Days Left"]);
  const wsRenewals = XLSX.utils.json_to_sheet(allRenewals.length ? allRenewals : [{ "Client": "", "Type": "", "Name": "", "Renewal Date": "", "Days Left": "", "Status": "" }]);
  wsRenewals["!cols"] = [28, 10, 35, 15, 10, 12].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsRenewals, "Renewals Summary");

  const date = new Date().toISOString().split("T")[0];
  XLSX.writeFile(wb, `agency-data-${date}.xlsx`);
};

// ─── XLS Import ───────────────────────────────────────────────────────────────
const importFromXLS = (file, onSuccess, onError) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array" });

      // Parse Clients sheet
      const clientSheet = wb.Sheets["Clients"];
      if (!clientSheet) throw new Error("No 'Clients' sheet found. Please use a file exported from this app.");
      const clientRows = XLSX.utils.sheet_to_json(clientSheet);

      // Parse Domains
      const domainRows = wb.Sheets["Domains"] ? XLSX.utils.sheet_to_json(wb.Sheets["Domains"]) : [];
      // Parse Hosting
      const hostingRows = wb.Sheets["Hosting"] ? XLSX.utils.sheet_to_json(wb.Sheets["Hosting"]) : [];
      // Parse Invoices
      const invoiceRows = wb.Sheets["Invoices"] ? XLSX.utils.sheet_to_json(wb.Sheets["Invoices"]) : [];

      const clients = clientRows
        .filter(r => r["Company Name"])
        .map(r => {
          const cid = r["Client ID"] || uid();
          const domains = domainRows
            .filter(d => d["Client ID"] === r["Client ID"] || d["Client Name"] === r["Company Name"])
            .map(d => ({ id: uid(), domain: String(d["Domain"] || ""), registrar: String(d["Registrar"] || ""), renewal: String(d["Renewal Date"] || ""), notes: String(d["Notes"] || "") }));
          const hosting = hostingRows
            .filter(h => h["Client ID"] === r["Client ID"] || h["Client Name"] === r["Company Name"])
            .map(h => ({ id: uid(), provider: String(h["Provider"] || ""), plan: String(h["Plan"] || ""), monthly: Number(h["Monthly Cost (INR)"] || 0), renewal: String(h["Renewal Date"] || ""), notes: String(h["Notes"] || "") }));
          const invoices = invoiceRows
            .filter(i => i["Client ID"] === r["Client ID"] || i["Client Name"] === r["Company Name"])
            .map(i => {
              let items = [];
              try { items = JSON.parse(i["Line Items (JSON)"] || "[]"); } catch {}
              if (!items.length) items = [{ id: uid(), desc: "Service", qty: 1, rate: Number(i["Subtotal (INR)"] || 0) }];
              return { id: uid(), number: String(i["Invoice Number"] || "INV-000"), issueDate: String(i["Issue Date"] || today()), dueDate: String(i["Due Date"] || today()), status: String(i["Status"] || "pending"), notes: String(i["Notes"] || ""), items };
            });
          return {
            id: cid,
            name: String(r["Company Name"] || ""),
            contact: String(r["Contact Person"] || ""),
            email: String(r["Email"] || ""),
            phone: String(r["Phone"] || ""),
            tags: String(r["Tags"] || "").split(",").map(t => t.trim()).filter(Boolean),
            domains,
            hosting,
            invoices,
          };
        });

      if (!clients.length) throw new Error("No client data found in the file.");
      onSuccess(clients);
    } catch (err) {
      onError(err.message || "Failed to parse file.");
    }
  };
  reader.readAsArrayBuffer(file);
};

// ─── Components ───────────────────────────────────────────────────────────────
const Badge = ({ days }) => {
  const d = typeof days === "number" ? days : daysBetween(days);
  const color = urgencyColor(d);
  const bg = urgencyBg(d);
  const label = d < 0 ? `Expired ${Math.abs(d)}d ago` : d === 0 ? "Expires today" : `${d}d left`;
  return (
    <span style={{ background: bg, color, padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 700, letterSpacing: "0.04em", whiteSpace: "nowrap" }}>
      {label}
    </span>
  );
};

const Tag = ({ label }) => (
  <span style={{ background: "#f0f0f8", color: "#555", padding: "2px 8px", borderRadius: 20, fontSize: 11, fontWeight: 500, marginRight: 4 }}>
    {label}
  </span>
);

const Input = ({ label, ...props }) => (
  <div style={{ marginBottom: 14 }}>
    {label && <label style={{ display: "block", fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.07em", color: "#888", marginBottom: 5 }}>{label}</label>}
    <input
      style={{ width: "100%", padding: "8px 12px", border: "1px solid #e0e0ec", borderRadius: 6, fontSize: 14, fontFamily: "inherit", background: "#fff", outline: "none", color: "#1a1a2e" }}
      {...props}
    />
  </div>
);

const Select = ({ label, children, ...props }) => (
  <div style={{ marginBottom: 14 }}>
    {label && <label style={{ display: "block", fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.07em", color: "#888", marginBottom: 5 }}>{label}</label>}
    <select
      style={{ width: "100%", padding: "8px 12px", border: "1px solid #e0e0ec", borderRadius: 6, fontSize: 14, fontFamily: "inherit", background: "#fff", outline: "none", color: "#1a1a2e" }}
      {...props}
    >
      {children}
    </select>
  </div>
);

const Textarea = ({ label, ...props }) => (
  <div style={{ marginBottom: 14 }}>
    {label && <label style={{ display: "block", fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.07em", color: "#888", marginBottom: 5 }}>{label}</label>}
    <textarea
      style={{ width: "100%", padding: "8px 12px", border: "1px solid #e0e0ec", borderRadius: 6, fontSize: 14, fontFamily: "inherit", background: "#fff", outline: "none", color: "#1a1a2e", resize: "vertical", minHeight: 70 }}
      {...props}
    />
  </div>
);

const Modal = ({ title, onClose, children }) => (
  <div style={{ position: "fixed", inset: 0, background: "rgba(10,10,30,0.55)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 16, backdropFilter: "blur(2px)" }}>
    <div style={{ background: "#fff", borderRadius: 12, width: "100%", maxWidth: 560, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 24px 80px rgba(0,0,0,0.25)" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "20px 24px 16px", borderBottom: "1px solid #eee" }}>
        <h3 style={{ fontFamily: "'DM Serif Display', serif", fontSize: 20, color: "#1a1a2e" }}>{title}</h3>
        <button onClick={onClose} style={{ background: "none", border: "none", fontSize: 20, cursor: "pointer", color: "#999", lineHeight: 1 }}>×</button>
      </div>
      <div style={{ padding: "20px 24px" }}>{children}</div>
    </div>
  </div>
);

const Btn = ({ variant = "primary", size = "md", children, ...props }) => {
  const base = { border: "none", borderRadius: 6, cursor: "pointer", fontFamily: "inherit", fontWeight: 600, letterSpacing: "0.02em", transition: "opacity 0.15s" };
  const variants = {
    primary: { background: "#1a1a2e", color: "#fff" },
    secondary: { background: "#f0f0f8", color: "#1a1a2e" },
    danger: { background: "#fef2f2", color: "#ef4444" },
    success: { background: "#f0fdf4", color: "#16a34a" },
    outline: { background: "transparent", color: "#1a1a2e", border: "1px solid #e0e0ec" },
  };
  const sizes = { sm: { padding: "5px 12px", fontSize: 12 }, md: { padding: "8px 16px", fontSize: 13 }, lg: { padding: "11px 22px", fontSize: 14 } };
  return <button style={{ ...base, ...variants[variant], ...sizes[size] }} {...props}>{children}</button>;
};

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [clients, setClients] = useState([]);
  const [loading, setLoading] = useState(true);
  const [view, setView] = useState("dashboard");
  const [selectedClient, setSelectedClient] = useState(null);
  const [modal, setModal] = useState(null);
  const [toast, setToast] = useState(null);
  const [search, setSearch] = useState("");
  const [importConfirm, setImportConfirm] = useState(null);
  const importRef = useRef();

  // Storage
  const save = useCallback(async (data) => {
    try { await window.storage.set("agency-clients", JSON.stringify(data)); } catch {}
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get("agency-clients");
        if (res?.value) setClients(JSON.parse(res.value));
        else { setClients(SEED_CLIENTS); save(SEED_CLIENTS); }
      } catch { setClients(SEED_CLIENTS); }
      setLoading(false);
    })();
  }, []);

  const update = (newClients) => { setClients(newClients); save(newClients); };

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3000);
  };

  // ── Export / Import ──
  const handleExport = () => {
    exportToXLS(clients);
    showToast("Exported successfully!");
  };

  const handleImportFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    importFromXLS(
      file,
      (parsed) => { setImportConfirm(parsed); },
      (err) => { showToast("Import failed: " + err, "error"); }
    );
    e.target.value = "";
  };

  const confirmImport = (mode) => {
    if (mode === "replace") {
      update(importConfirm);
      showToast(`Imported ${importConfirm.length} clients (replaced all data).`);
    } else {
      // Merge: add clients not already present by name
      const existingNames = new Set(clients.map(c => c.name.toLowerCase()));
      const newOnes = importConfirm.filter(c => !existingNames.has(c.name.toLowerCase()));
      const merged = [...clients, ...newOnes];
      update(merged);
      showToast(`Merged: ${newOnes.length} new client(s) added.`);
    }
    setImportConfirm(null);
  };

  // ── Derived data ──
  const allRenewals = clients.flatMap(c => [
    ...c.domains.map(d => ({ clientId: c.id, clientName: c.name, type: "Domain", label: d.domain, renewal: d.renewal, id: d.id })),
    ...c.hosting.map(h => ({ clientId: c.id, clientName: c.name, type: "Hosting", label: h.provider + " — " + h.plan, renewal: h.renewal, id: h.id })),
  ]).sort((a, b) => new Date(a.renewal) - new Date(b.renewal));

  const urgent = allRenewals.filter(r => daysBetween(r.renewal) <= 30);
  const overdue = allRenewals.filter(r => daysBetween(r.renewal) < 0);

  const allInvoices = clients.flatMap(c => c.invoices.map(inv => ({ ...inv, clientName: c.name, clientId: c.id })))
    .sort((a, b) => new Date(b.issueDate) - new Date(a.issueDate));

  const revenue = allInvoices.filter(i => i.status === "paid").reduce((s, i) => s + Math.round(i.items.reduce((ss, ii) => ss + ii.qty * ii.rate, 0) * 1.18), 0);
  const pending = allInvoices.filter(i => i.status === "pending").reduce((s, i) => s + Math.round(i.items.reduce((ss, ii) => ss + ii.qty * ii.rate, 0) * 1.18), 0);

  // ── Filtered clients ──
  const filtered = clients.filter(c =>
    c.name.toLowerCase().includes(search.toLowerCase()) ||
    c.contact.toLowerCase().includes(search.toLowerCase()) ||
    c.email.toLowerCase().includes(search.toLowerCase())
  );

  // ── Reminder email preview ──
  const reminderText = (r) => {
    const d = daysBetween(r.renewal);
    const sub = d < 0 ? `EXPIRED: ${r.type} for ${r.label}` : `Renewal Reminder: ${r.type} expires in ${d} days`;
    const body = `Hi,\n\nThis is a reminder that your ${r.type.toLowerCase()} for ${r.label} ${d < 0 ? 'expired on ' + fmt(r.renewal) + ' — please renew immediately' : 'is due for renewal on ' + fmt(r.renewal) + ' (' + d + ' days remaining)'}.\n\nPlease take action to avoid any service disruption.\n\nBest regards,\nYour Agency Team`;
    return { sub, body };
  };

  // ─── Views ──────────────────────────────────────────────────────────────────
  const styles = {
    card: { background: "#fff", borderRadius: 10, border: "1px solid #e8e8f0", padding: "20px 22px", boxShadow: "0 2px 8px rgba(0,0,0,0.04)" },
    section: { marginBottom: 28 },
    th: { padding: "9px 14px", fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.07em", color: "#888", textAlign: "left", borderBottom: "2px solid #eee", whiteSpace: "nowrap" },
    td: { padding: "11px 14px", fontSize: 13, borderBottom: "1px solid #f0f0f8", verticalAlign: "middle" },
    h2: { fontFamily: "'DM Serif Display', serif", fontSize: 22, color: "#1a1a2e", marginBottom: 4 },
    label: { fontSize: 11, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.07em", color: "#aaa" },
  };

  const nav = (id, label, icon) => (
    <button key={id} onClick={() => { setView(id); setSelectedClient(null); }}
      style={{ display: "flex", alignItems: "center", gap: 8, padding: "9px 14px", borderRadius: 7, border: "none", cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: view === id ? 700 : 500, background: view === id ? "#1a1a2e" : "transparent", color: view === id ? "#fff" : "#666", width: "100%", transition: "all 0.15s" }}>
      <span style={{ fontSize: 16 }}>{icon}</span>{label}
    </button>
  );

  // ─── Dashboard ───────────────────────────────────────────────────────────────
  const Dashboard = () => (
    <div>
      <div style={{ marginBottom: 24 }}>
        <h1 style={{ fontFamily: "'DM Serif Display', serif", fontSize: 28, color: "#1a1a2e" }}>Dashboard</h1>
        <p style={{ color: "#888", fontSize: 13, marginTop: 2 }}>Overview as of {fmt(today())}</p>
      </div>

      {/* KPI row */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14, marginBottom: 24 }}>
        {[
          { label: "Total Clients", value: clients.length, icon: "👥", color: "#6366f1" },
          { label: "Urgent Renewals", value: urgent.length, icon: "⚡", color: "#f97316" },
          { label: "Overdue", value: overdue.length, icon: "🔴", color: "#ef4444" },
          { label: "Revenue Collected", value: currency(revenue), icon: "✅", color: "#22c55e" },
          { label: "Pending Invoices", value: currency(pending), icon: "⏳", color: "#eab308" },
        ].map((k, i) => (
          <div key={i} style={{ ...styles.card, borderTop: `3px solid ${k.color}` }}>
            <div style={{ fontSize: 22, marginBottom: 6 }}>{k.icon}</div>
            <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 20, color: "#1a1a2e", lineHeight: 1.1 }}>{k.value}</div>
            <div style={{ ...styles.label, marginTop: 4 }}>{k.label}</div>
          </div>
        ))}
      </div>

      {/* Upcoming renewals */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 20 }}>
        <div style={styles.card}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <h3 style={{ fontFamily: "'DM Serif Display', serif", fontSize: 17, color: "#1a1a2e" }}>Upcoming Renewals</h3>
            <span style={{ fontSize: 11, color: "#888" }}>Next 60 days</span>
          </div>
          {allRenewals.filter(r => daysBetween(r.renewal) <= 60).slice(0, 7).length === 0
            ? <p style={{ color: "#aaa", fontSize: 13 }}>No renewals in the next 60 days.</p>
            : allRenewals.filter(r => daysBetween(r.renewal) <= 60).slice(0, 7).map(r => {
              const d = daysBetween(r.renewal);
              return (
                <div key={r.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 0", borderBottom: "1px solid #f5f5fa" }}>
                  <div>
                    <span style={{ fontSize: 10, background: "#f0f0f8", color: "#666", padding: "1px 6px", borderRadius: 3, fontWeight: 700, marginRight: 6 }}>{r.type}</span>
                    <span style={{ fontSize: 13, fontWeight: 500, color: "#1a1a2e" }}>{r.label}</span>
                    <div style={{ fontSize: 11, color: "#aaa", marginTop: 1 }}>{r.clientName}</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <Badge days={d} />
                    <div style={{ fontSize: 11, color: "#aaa", marginTop: 2 }}>{fmt(r.renewal)}</div>
                  </div>
                </div>
              );
            })}
        </div>

        {/* Recent invoices */}
        <div style={styles.card}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <h3 style={{ fontFamily: "'DM Serif Display', serif", fontSize: 17, color: "#1a1a2e" }}>Recent Invoices</h3>
            <Btn size="sm" variant="outline" onClick={() => setView("invoices")}>View all</Btn>
          </div>
          {allInvoices.slice(0, 6).length === 0
            ? <p style={{ color: "#aaa", fontSize: 13 }}>No invoices yet.</p>
            : allInvoices.slice(0, 6).map(inv => {
              const total = Math.round(inv.items.reduce((s, i) => s + i.qty * i.rate, 0) * 1.18);
              const sc = { paid: "#22c55e", pending: "#f97316", overdue: "#ef4444" };
              return (
                <div key={inv.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "7px 0", borderBottom: "1px solid #f5f5fa" }}>
                  <div>
                    <span style={{ fontSize: 13, fontWeight: 600, color: "#1a1a2e" }}>#{inv.number}</span>
                    <span style={{ fontSize: 12, color: "#888", marginLeft: 8 }}>{inv.clientName}</span>
                  </div>
                  <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                    <span style={{ fontSize: 13, fontWeight: 600 }}>{currency(total)}</span>
                    <span style={{ fontSize: 11, fontWeight: 700, color: sc[inv.status] || "#888", background: sc[inv.status] + "20" || "#f5f5fa", padding: "2px 8px", borderRadius: 4 }}>{inv.status.toUpperCase()}</span>
                  </div>
                </div>
              );
            })}
        </div>
      </div>

      {/* Client quick view */}
      <div style={styles.card}>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 14 }}>
          <h3 style={{ fontFamily: "'DM Serif Display', serif", fontSize: 17, color: "#1a1a2e" }}>All Clients</h3>
          <Btn size="sm" variant="outline" onClick={() => setView("clients")}>Manage</Btn>
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr>
            {["Client", "Contact", "Domains", "Hosting", "Next Renewal"].map(h => <th key={h} style={styles.th}>{h}</th>)}
          </tr></thead>
          <tbody>
            {clients.map(c => {
              const nr = [...c.domains, ...c.hosting].map(x => x.renewal).sort()[0];
              return (
                <tr key={c.id} style={{ cursor: "pointer" }} onClick={() => { setSelectedClient(c.id); setView("client-detail"); }}>
                  <td style={styles.td}><span style={{ fontWeight: 600, color: "#1a1a2e" }}>{c.name}</span></td>
                  <td style={styles.td}><span style={{ color: "#666" }}>{c.contact}</span></td>
                  <td style={styles.td}><span style={{ fontWeight: 600 }}>{c.domains.length}</span></td>
                  <td style={styles.td}><span style={{ fontWeight: 600 }}>{c.hosting.length}</span></td>
                  <td style={styles.td}>{nr ? <><Badge days={daysBetween(nr)} /><span style={{ fontSize: 11, color: "#aaa", marginLeft: 6 }}>{fmt(nr)}</span></> : "—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );

  // ─── Clients List ─────────────────────────────────────────────────────────
  const ClientsList = () => {
    const [showAdd, setShowAdd] = useState(false);
    const [form, setForm] = useState({ name: "", contact: "", email: "", phone: "", tags: "" });

    const addClient = () => {
      if (!form.name.trim()) return;
      const c = { id: uid(), name: form.name, contact: form.contact, email: form.email, phone: form.phone, tags: form.tags.split(",").map(t => t.trim()).filter(Boolean), domains: [], hosting: [], invoices: [] };
      const next = [...clients, c];
      update(next);
      setShowAdd(false);
      setForm({ name: "", contact: "", email: "", phone: "", tags: "" });
      showToast("Client added!");
    };

    return (
      <div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <h1 style={styles.h2}>Clients</h1>
          <div style={{ display: "flex", gap: 10 }}>
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search clients…" style={{ padding: "7px 12px", border: "1px solid #e0e0ec", borderRadius: 6, fontSize: 13, fontFamily: "inherit", outline: "none", width: 220 }} />
            <Btn onClick={() => setShowAdd(true)}>+ Add Client</Btn>
          </div>
        </div>
        <div style={{ display: "grid", gap: 12 }}>
          {filtered.map(c => {
            const nr = [...c.domains, ...c.hosting].map(x => x.renewal).sort()[0];
            const d = nr ? daysBetween(nr) : null;
            return (
              <div key={c.id} style={{ ...styles.card, display: "flex", alignItems: "center", justifyContent: "space-between", cursor: "pointer", borderLeft: d !== null ? `3px solid ${urgencyColor(d)}` : "3px solid #e0e0ec" }}
                onClick={() => { setSelectedClient(c.id); setView("client-detail"); }}>
                <div style={{ flex: 1 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
                    <span style={{ fontFamily: "'DM Serif Display', serif", fontSize: 17, color: "#1a1a2e" }}>{c.name}</span>
                    {c.tags.map(t => <Tag key={t} label={t} />)}
                  </div>
                  <div style={{ fontSize: 12, color: "#888" }}>{c.contact} · {c.email} · {c.phone}</div>
                </div>
                <div style={{ display: "flex", gap: 20, alignItems: "center" }}>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ fontWeight: 700, fontSize: 18, color: "#1a1a2e" }}>{c.domains.length}</div>
                    <div style={styles.label}>Domains</div>
                  </div>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ fontWeight: 700, fontSize: 18, color: "#1a1a2e" }}>{c.hosting.length}</div>
                    <div style={styles.label}>Hosting</div>
                  </div>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ fontWeight: 700, fontSize: 18, color: "#1a1a2e" }}>{c.invoices.length}</div>
                    <div style={styles.label}>Invoices</div>
                  </div>
                  {nr && <Badge days={d} />}
                </div>
              </div>
            );
          })}
        </div>

        {showAdd && (
          <Modal title="Add New Client" onClose={() => setShowAdd(false)}>
            <Input label="Company Name *" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} placeholder="Acme Corp" />
            <Input label="Contact Person" value={form.contact} onChange={e => setForm({ ...form, contact: e.target.value })} placeholder="John Doe" />
            <Input label="Email" type="email" value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} placeholder="john@acme.com" />
            <Input label="Phone" value={form.phone} onChange={e => setForm({ ...form, phone: e.target.value })} placeholder="+91 98765 43210" />
            <Input label="Tags (comma separated)" value={form.tags} onChange={e => setForm({ ...form, tags: e.target.value })} placeholder="ecommerce, active" />
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 8 }}>
              <Btn variant="secondary" onClick={() => setShowAdd(false)}>Cancel</Btn>
              <Btn onClick={addClient}>Add Client</Btn>
            </div>
          </Modal>
        )}
      </div>
    );
  };

  // ─── Client Detail ─────────────────────────────────────────────────────────
  const ClientDetail = () => {
    const c = clients.find(x => x.id === selectedClient);
    if (!c) return null;
    const [tab, setTab] = useState("domains");
    const [domainForm, setDomainForm] = useState({ domain: "", registrar: "", renewal: "", notes: "" });
    const [hostingForm, setHostingForm] = useState({ provider: "", plan: "", renewal: "", monthly: "", notes: "" });
    const [showDomain, setShowDomain] = useState(false);
    const [showHosting, setShowHosting] = useState(false);
    const [showReminder, setShowReminder] = useState(null);
    const [showInvoice, setShowInvoice] = useState(false);
    const [editClient, setEditClient] = useState(false);
    const [editForm, setEditForm] = useState({ name: c.name, contact: c.contact, email: c.email, phone: c.phone, tags: c.tags.join(", ") });

    // Invoice form
    const [invForm, setInvForm] = useState({
      number: "INV-" + String(Date.now()).slice(-5),
      issueDate: today(),
      dueDate: new Date(Date.now() + 15 * 86400000).toISOString().split("T")[0],
      status: "pending",
      notes: "",
      items: [{ id: uid(), desc: "", qty: 1, rate: "" }],
    });

    const saveEdit = () => {
      const next = clients.map(x => x.id === c.id ? { ...x, ...editForm, tags: editForm.tags.split(",").map(t => t.trim()).filter(Boolean) } : x);
      update(next);
      setEditClient(false);
      showToast("Client updated!");
    };

    const deleteClient = () => {
      if (!window.confirm(`Delete ${c.name}? This cannot be undone.`)) return;
      update(clients.filter(x => x.id !== c.id));
      setView("clients");
      showToast("Client deleted.", "error");
    };

    const addDomain = () => {
      if (!domainForm.domain) return;
      const next = clients.map(x => x.id === c.id ? { ...x, domains: [...x.domains, { id: uid(), ...domainForm }] } : x);
      update(next);
      setShowDomain(false);
      setDomainForm({ domain: "", registrar: "", renewal: "", notes: "" });
      showToast("Domain added!");
    };

    const deleteDomain = (did) => {
      const next = clients.map(x => x.id === c.id ? { ...x, domains: x.domains.filter(d => d.id !== did) } : x);
      update(next);
    };

    const addHosting = () => {
      if (!hostingForm.provider) return;
      const next = clients.map(x => x.id === c.id ? { ...x, hosting: [...x.hosting, { id: uid(), ...hostingForm, monthly: Number(hostingForm.monthly) }] } : x);
      update(next);
      setShowHosting(false);
      setHostingForm({ provider: "", plan: "", renewal: "", monthly: "", notes: "" });
      showToast("Hosting added!");
    };

    const deleteHosting = (hid) => {
      const next = clients.map(x => x.id === c.id ? { ...x, hosting: x.hosting.filter(h => h.id !== hid) } : x);
      update(next);
    };

    const addInvItem = () => setInvForm(f => ({ ...f, items: [...f.items, { id: uid(), desc: "", qty: 1, rate: "" }] }));
    const updateItem = (id, field, val) => setInvForm(f => ({ ...f, items: f.items.map(i => i.id === id ? { ...i, [field]: field === "qty" || field === "rate" ? Number(val) : val } : i) }));
    const removeItem = (id) => setInvForm(f => ({ ...f, items: f.items.filter(i => i.id !== id) }));

    const saveInvoice = () => {
      const inv = { ...invForm, id: uid(), createdAt: today() };
      const next = clients.map(x => x.id === c.id ? { ...x, invoices: [...x.invoices, inv] } : x);
      update(next);
      setShowInvoice(false);
      showToast("Invoice created!");
    };

    const updateInvStatus = (invId, status) => {
      const next = clients.map(x => x.id === c.id ? { ...x, invoices: x.invoices.map(i => i.id === invId ? { ...i, status } : i) } : x);
      update(next);
      showToast("Invoice updated!");
    };

    const previewInvoice = (inv) => {
      const html = generateInvoiceHTML(inv, c);
      const w = window.open("", "_blank");
      w.document.write(html);
      w.document.close();
    };

    const tabBtn = (id, label) => (
      <button key={id} onClick={() => setTab(id)} style={{ padding: "8px 18px", border: "none", borderBottom: tab === id ? "2px solid #1a1a2e" : "2px solid transparent", background: "transparent", cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: tab === id ? 700 : 500, color: tab === id ? "#1a1a2e" : "#888", transition: "all 0.15s" }}>
        {label}
      </button>
    );

    const subtotal = invForm.items.reduce((s, i) => s + i.qty * (Number(i.rate) || 0), 0);
    const gst = Math.round(subtotal * 0.18);

    return (
      <div>
        {/* Header */}
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 20 }}>
          <div>
            <button onClick={() => setView("clients")} style={{ background: "none", border: "none", cursor: "pointer", color: "#888", fontSize: 12, fontFamily: "inherit", marginBottom: 6, padding: 0 }}>← Back to Clients</button>
            <h1 style={{ fontFamily: "'DM Serif Display', serif", fontSize: 26, color: "#1a1a2e" }}>{c.name}</h1>
            <div style={{ display: "flex", gap: 6, marginTop: 5 }}>{c.tags.map(t => <Tag key={t} label={t} />)}</div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <Btn size="sm" variant="secondary" onClick={() => setEditClient(true)}>Edit</Btn>
            <Btn size="sm" variant="danger" onClick={deleteClient}>Delete</Btn>
          </div>
        </div>

        {/* Contact card */}
        <div style={{ ...styles.card, marginBottom: 20, display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 16 }}>
          {[["Contact", c.contact], ["Email", c.email], ["Phone", c.phone], ["Domains", c.domains.length], ["Hosting", c.hosting.length], ["Invoices", c.invoices.length]].map(([l, v]) => (
            <div key={l}>
              <div style={styles.label}>{l}</div>
              <div style={{ fontSize: 14, fontWeight: 500, color: "#1a1a2e", marginTop: 3 }}>{v}</div>
            </div>
          ))}
        </div>

        {/* Tabs */}
        <div style={{ display: "flex", borderBottom: "1px solid #eee", marginBottom: 20 }}>
          {tabBtn("domains", "Domains")}
          {tabBtn("hosting", "Hosting")}
          {tabBtn("invoices", "Invoices")}
        </div>

        {/* Domains */}
        {tab === "domains" && (
          <div>
            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
              <Btn size="sm" onClick={() => setShowDomain(true)}>+ Add Domain</Btn>
            </div>
            {c.domains.length === 0 ? <p style={{ color: "#aaa", fontSize: 13 }}>No domains added yet.</p> :
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr>{["Domain", "Registrar", "Renewal Date", "Status", "Notes", "Actions"].map(h => <th key={h} style={styles.th}>{h}</th>)}</tr></thead>
                <tbody>
                  {c.domains.map(d => (
                    <tr key={d.id}>
                      <td style={{ ...styles.td, fontWeight: 600, color: "#1a1a2e" }}>{d.domain}</td>
                      <td style={styles.td}>{d.registrar || "—"}</td>
                      <td style={styles.td}>{fmt(d.renewal)}</td>
                      <td style={styles.td}><Badge days={daysBetween(d.renewal)} /></td>
                      <td style={styles.td}><span style={{ color: "#888", fontSize: 12 }}>{d.notes || "—"}</span></td>
                      <td style={styles.td}>
                        <div style={{ display: "flex", gap: 6 }}>
                          <Btn size="sm" variant="secondary" onClick={() => setShowReminder({ type: "Domain", label: d.domain, renewal: d.renewal, id: d.id })}>Reminder</Btn>
                          <Btn size="sm" variant="danger" onClick={() => deleteDomain(d.id)}>×</Btn>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>}
          </div>
        )}

        {/* Hosting */}
        {tab === "hosting" && (
          <div>
            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
              <Btn size="sm" onClick={() => setShowHosting(true)}>+ Add Hosting</Btn>
            </div>
            {c.hosting.length === 0 ? <p style={{ color: "#aaa", fontSize: 13 }}>No hosting records yet.</p> :
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr>{["Provider", "Plan", "Monthly Cost", "Renewal Date", "Status", "Notes", "Actions"].map(h => <th key={h} style={styles.th}>{h}</th>)}</tr></thead>
                <tbody>
                  {c.hosting.map(h => (
                    <tr key={h.id}>
                      <td style={{ ...styles.td, fontWeight: 600 }}>{h.provider}</td>
                      <td style={styles.td}>{h.plan || "—"}</td>
                      <td style={styles.td}>{h.monthly ? currency(h.monthly) : "—"}</td>
                      <td style={styles.td}>{fmt(h.renewal)}</td>
                      <td style={styles.td}><Badge days={daysBetween(h.renewal)} /></td>
                      <td style={styles.td}><span style={{ color: "#888", fontSize: 12 }}>{h.notes || "—"}</span></td>
                      <td style={styles.td}>
                        <div style={{ display: "flex", gap: 6 }}>
                          <Btn size="sm" variant="secondary" onClick={() => setShowReminder({ type: "Hosting", label: h.provider + " — " + h.plan, renewal: h.renewal, id: h.id })}>Reminder</Btn>
                          <Btn size="sm" variant="danger" onClick={() => deleteHosting(h.id)}>×</Btn>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>}
          </div>
        )}

        {/* Invoices */}
        {tab === "invoices" && (
          <div>
            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
              <Btn size="sm" onClick={() => setShowInvoice(true)}>+ Create Invoice</Btn>
            </div>
            {c.invoices.length === 0 ? <p style={{ color: "#aaa", fontSize: 13 }}>No invoices yet.</p> :
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr>{["Invoice #", "Issue Date", "Due Date", "Amount", "Status", "Actions"].map(h => <th key={h} style={styles.th}>{h}</th>)}</tr></thead>
                <tbody>
                  {c.invoices.map(inv => {
                    const total = Math.round(inv.items.reduce((s, i) => s + i.qty * i.rate, 0) * 1.18);
                    const sc = { paid: "#22c55e", pending: "#f97316", overdue: "#ef4444" };
                    return (
                      <tr key={inv.id}>
                        <td style={{ ...styles.td, fontWeight: 700 }}>#{inv.number}</td>
                        <td style={styles.td}>{fmt(inv.issueDate)}</td>
                        <td style={styles.td}>{fmt(inv.dueDate)}</td>
                        <td style={{ ...styles.td, fontWeight: 700 }}>{currency(total)}</td>
                        <td style={styles.td}>
                          <select value={inv.status} onChange={e => updateInvStatus(inv.id, e.target.value)}
                            style={{ padding: "3px 8px", borderRadius: 4, border: "1px solid #e0e0ec", fontSize: 11, fontWeight: 700, color: sc[inv.status], background: sc[inv.status] + "15", fontFamily: "inherit", cursor: "pointer", outline: "none" }}>
                            <option value="pending">PENDING</option>
                            <option value="paid">PAID</option>
                            <option value="overdue">OVERDUE</option>
                          </select>
                        </td>
                        <td style={styles.td}><Btn size="sm" variant="secondary" onClick={() => previewInvoice(inv)}>Preview</Btn></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>}
          </div>
        )}

        {/* Modals */}
        {editClient && (
          <Modal title="Edit Client" onClose={() => setEditClient(false)}>
            <Input label="Company Name" value={editForm.name} onChange={e => setEditForm({ ...editForm, name: e.target.value })} />
            <Input label="Contact Person" value={editForm.contact} onChange={e => setEditForm({ ...editForm, contact: e.target.value })} />
            <Input label="Email" type="email" value={editForm.email} onChange={e => setEditForm({ ...editForm, email: e.target.value })} />
            <Input label="Phone" value={editForm.phone} onChange={e => setEditForm({ ...editForm, phone: e.target.value })} />
            <Input label="Tags" value={editForm.tags} onChange={e => setEditForm({ ...editForm, tags: e.target.value })} />
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <Btn variant="secondary" onClick={() => setEditClient(false)}>Cancel</Btn>
              <Btn onClick={saveEdit}>Save</Btn>
            </div>
          </Modal>
        )}

        {showDomain && (
          <Modal title="Add Domain" onClose={() => setShowDomain(false)}>
            <Input label="Domain Name *" value={domainForm.domain} onChange={e => setDomainForm({ ...domainForm, domain: e.target.value })} placeholder="example.com" />
            <Input label="Registrar" value={domainForm.registrar} onChange={e => setDomainForm({ ...domainForm, registrar: e.target.value })} placeholder="GoDaddy, Namecheap…" />
            <Input label="Renewal Date" type="date" value={domainForm.renewal} onChange={e => setDomainForm({ ...domainForm, renewal: e.target.value })} />
            <Input label="Notes" value={domainForm.notes} onChange={e => setDomainForm({ ...domainForm, notes: e.target.value })} placeholder="Any notes…" />
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <Btn variant="secondary" onClick={() => setShowDomain(false)}>Cancel</Btn>
              <Btn onClick={addDomain}>Add Domain</Btn>
            </div>
          </Modal>
        )}

        {showHosting && (
          <Modal title="Add Hosting" onClose={() => setShowHosting(false)}>
            <Input label="Provider *" value={hostingForm.provider} onChange={e => setHostingForm({ ...hostingForm, provider: e.target.value })} placeholder="Cloudways, WP Engine…" />
            <Input label="Plan" value={hostingForm.plan} onChange={e => setHostingForm({ ...hostingForm, plan: e.target.value })} placeholder="Pro 4GB, Startup…" />
            <Input label="Monthly Cost (₹)" type="number" value={hostingForm.monthly} onChange={e => setHostingForm({ ...hostingForm, monthly: e.target.value })} placeholder="4500" />
            <Input label="Renewal Date" type="date" value={hostingForm.renewal} onChange={e => setHostingForm({ ...hostingForm, renewal: e.target.value })} />
            <Input label="Notes" value={hostingForm.notes} onChange={e => setHostingForm({ ...hostingForm, notes: e.target.value })} />
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
              <Btn variant="secondary" onClick={() => setShowHosting(false)}>Cancel</Btn>
              <Btn onClick={addHosting}>Add Hosting</Btn>
            </div>
          </Modal>
        )}

        {showReminder && (() => {
          const { sub, body } = reminderText(showReminder);
          return (
            <Modal title="Send Reminder" onClose={() => setShowReminder(null)}>
              <div style={{ background: urgencyBg(daysBetween(showReminder.renewal)), border: `1px solid ${urgencyColor(daysBetween(showReminder.renewal))}30`, borderRadius: 8, padding: 14, marginBottom: 16 }}>
                <div style={{ fontSize: 13, fontWeight: 700, color: urgencyColor(daysBetween(showReminder.renewal)) }}>
                  {showReminder.type}: {showReminder.label}
                </div>
                <div style={{ fontSize: 12, color: "#666", marginTop: 4 }}>Renewal: {fmt(showReminder.renewal)} · <Badge days={daysBetween(showReminder.renewal)} /></div>
              </div>
              <div style={{ marginBottom: 12 }}>
                <div style={styles.label}>Email Subject</div>
                <div style={{ background: "#f5f5fa", borderRadius: 6, padding: "10px 14px", fontSize: 13, fontWeight: 600, marginTop: 5, color: "#1a1a2e", userSelect: "all" }}>{sub}</div>
              </div>
              <div style={{ marginBottom: 16 }}>
                <div style={styles.label}>Email Body</div>
                <pre style={{ background: "#f5f5fa", borderRadius: 6, padding: "12px 14px", fontSize: 12, color: "#444", marginTop: 5, whiteSpace: "pre-wrap", fontFamily: "inherit", lineHeight: 1.7, userSelect: "all" }}>{body}</pre>
              </div>
              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                <Btn variant="secondary" onClick={() => setShowReminder(null)}>Close</Btn>
                <Btn onClick={() => { navigator.clipboard?.writeText(`Subject: ${sub}\n\n${body}`); showToast("Copied to clipboard!"); setShowReminder(null); }}>Copy Email</Btn>
                <Btn variant="success" onClick={() => { window.open(`mailto:${c.email}?subject=${encodeURIComponent(sub)}&body=${encodeURIComponent(body)}`); }}>Open in Mail</Btn>
              </div>
            </Modal>
          );
        })()}

        {showInvoice && (
          <Modal title={`Create Invoice for ${c.name}`} onClose={() => setShowInvoice(false)}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
              <Input label="Invoice Number" value={invForm.number} onChange={e => setInvForm({ ...invForm, number: e.target.value })} />
              <Select label="Status" value={invForm.status} onChange={e => setInvForm({ ...invForm, status: e.target.value })}>
                <option value="pending">Pending</option>
                <option value="paid">Paid</option>
                <option value="overdue">Overdue</option>
              </Select>
              <Input label="Issue Date" type="date" value={invForm.issueDate} onChange={e => setInvForm({ ...invForm, issueDate: e.target.value })} />
              <Input label="Due Date" type="date" value={invForm.dueDate} onChange={e => setInvForm({ ...invForm, dueDate: e.target.value })} />
            </div>

            <div style={{ marginBottom: 8 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <div style={styles.label}>Line Items</div>
                <Btn size="sm" variant="outline" onClick={addInvItem}>+ Add Item</Btn>
              </div>
              {invForm.items.map((item, idx) => (
                <div key={item.id} style={{ display: "grid", gridTemplateColumns: "3fr 1fr 2fr auto", gap: 8, marginBottom: 8, alignItems: "flex-end" }}>
                  <Input label={idx === 0 ? "Description" : ""} value={item.desc} onChange={e => updateItem(item.id, "desc", e.target.value)} placeholder="Service description" />
                  <Input label={idx === 0 ? "Qty" : ""} type="number" value={item.qty} onChange={e => updateItem(item.id, "qty", e.target.value)} min={1} />
                  <Input label={idx === 0 ? "Rate (₹)" : ""} type="number" value={item.rate} onChange={e => updateItem(item.id, "rate", e.target.value)} placeholder="5000" />
                  <button onClick={() => removeItem(item.id)} style={{ background: "#fef2f2", color: "#ef4444", border: "none", borderRadius: 6, padding: "8px 10px", cursor: "pointer", fontSize: 14, marginBottom: 14 }}>×</button>
                </div>
              ))}
              <div style={{ background: "#f5f5fa", borderRadius: 6, padding: "10px 14px", marginTop: 8, display: "flex", justifyContent: "flex-end", gap: 24 }}>
                <span style={{ fontSize: 13, color: "#666" }}>Subtotal: <strong style={{ color: "#1a1a2e" }}>{currency(subtotal)}</strong></span>
                <span style={{ fontSize: 13, color: "#666" }}>GST 18%: <strong style={{ color: "#1a1a2e" }}>{currency(gst)}</strong></span>
                <span style={{ fontSize: 14, fontWeight: 700, color: "#1a1a2e" }}>Total: {currency(subtotal + gst)}</span>
              </div>
            </div>

            <Textarea label="Notes" value={invForm.notes} onChange={e => setInvForm({ ...invForm, notes: e.target.value })} placeholder="Payment terms, bank details…" />
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 8 }}>
              <Btn variant="secondary" onClick={() => setShowInvoice(false)}>Cancel</Btn>
              <Btn onClick={saveInvoice}>Save Invoice</Btn>
            </div>
          </Modal>
        )}
      </div>
    );
  };

  // ─── Renewals View ───────────────────────────────────────────────────────────
  const RenewalsView = () => {
    const [filter, setFilter] = useState("all");
    const filtered2 = allRenewals.filter(r => {
      if (filter === "urgent") return daysBetween(r.renewal) <= 14 && daysBetween(r.renewal) >= 0;
      if (filter === "overdue") return daysBetween(r.renewal) < 0;
      if (filter === "month") return daysBetween(r.renewal) >= 0 && daysBetween(r.renewal) <= 30;
      return true;
    });

    return (
      <div>
        <h1 style={{ ...styles.h2, marginBottom: 18 }}>Renewals Calendar</h1>
        <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
          {[["all", "All"], ["overdue", "Overdue"], ["urgent", "≤ 14 Days"], ["month", "This Month"]].map(([id, label]) => (
            <button key={id} onClick={() => setFilter(id)} style={{ padding: "6px 14px", borderRadius: 20, border: filter === id ? "none" : "1px solid #e0e0ec", background: filter === id ? "#1a1a2e" : "#fff", color: filter === id ? "#fff" : "#666", cursor: "pointer", fontSize: 12, fontWeight: 600, fontFamily: "inherit" }}>{label}</button>
          ))}
        </div>
        <div style={styles.card}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>{["Type", "Domain / Service", "Client", "Renewal Date", "Status", "Action"].map(h => <th key={h} style={styles.th}>{h}</th>)}</tr></thead>
            <tbody>
              {filtered2.map(r => {
                const client = clients.find(c => c.id === r.clientId);
                const { sub, body } = reminderText(r);
                return (
                  <tr key={r.id}>
                    <td style={styles.td}><span style={{ fontSize: 11, background: r.type === "Domain" ? "#eff6ff" : "#f0fdf4", color: r.type === "Domain" ? "#3b82f6" : "#16a34a", padding: "2px 8px", borderRadius: 4, fontWeight: 700 }}>{r.type}</span></td>
                    <td style={{ ...styles.td, fontWeight: 600 }}>{r.label}</td>
                    <td style={styles.td}>{r.clientName}</td>
                    <td style={styles.td}>{fmt(r.renewal)}</td>
                    <td style={styles.td}><Badge days={daysBetween(r.renewal)} /></td>
                    <td style={styles.td}>
                      <div style={{ display: "flex", gap: 6 }}>
                        <Btn size="sm" variant="secondary" onClick={() => {
                          navigator.clipboard?.writeText(`Subject: ${sub}\n\n${body}`);
                          showToast("Reminder copied!");
                        }}>Copy Reminder</Btn>
                        {client && <Btn size="sm" variant="outline" onClick={() => window.open(`mailto:${client.email}?subject=${encodeURIComponent(sub)}&body=${encodeURIComponent(body)}`)}>Email</Btn>}
                      </div>
                    </td>
                  </tr>
                );
              })}
              {filtered2.length === 0 && <tr><td colSpan={6} style={{ ...styles.td, textAlign: "center", color: "#aaa", padding: 32 }}>No renewals in this filter.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  // ─── Invoices View ────────────────────────────────────────────────────────────
  const InvoicesView = () => {
    const [filter, setFilter] = useState("all");
    const filtered3 = allInvoices.filter(i => filter === "all" || i.status === filter);

    return (
      <div>
        <h1 style={{ ...styles.h2, marginBottom: 18 }}>All Invoices</h1>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 22 }}>
          {[{ label: "Paid", value: currency(revenue), color: "#22c55e" }, { label: "Pending", value: currency(pending), color: "#f97316" }, { label: "Total Invoices", value: allInvoices.length, color: "#6366f1" }].map(k => (
            <div key={k.label} style={{ ...styles.card, borderTop: `3px solid ${k.color}` }}>
              <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 22, color: "#1a1a2e" }}>{k.value}</div>
              <div style={styles.label}>{k.label}</div>
            </div>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          {["all", "pending", "paid", "overdue"].map(id => (
            <button key={id} onClick={() => setFilter(id)} style={{ padding: "6px 14px", borderRadius: 20, border: filter === id ? "none" : "1px solid #e0e0ec", background: filter === id ? "#1a1a2e" : "#fff", color: filter === id ? "#fff" : "#666", cursor: "pointer", fontSize: 12, fontWeight: 600, fontFamily: "inherit", textTransform: "capitalize" }}>{id}</button>
          ))}
        </div>
        <div style={styles.card}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>{["Invoice #", "Client", "Issue Date", "Due Date", "Amount", "Status", "Actions"].map(h => <th key={h} style={styles.th}>{h}</th>)}</tr></thead>
            <tbody>
              {filtered3.map(inv => {
                const total = Math.round(inv.items.reduce((s, i) => s + i.qty * i.rate, 0) * 1.18);
                const sc = { paid: "#22c55e", pending: "#f97316", overdue: "#ef4444" };
                const client = clients.find(c => c.id === inv.clientId);
                return (
                  <tr key={inv.id}>
                    <td style={{ ...styles.td, fontWeight: 700 }}>#{inv.number}</td>
                    <td style={styles.td}>{inv.clientName}</td>
                    <td style={styles.td}>{fmt(inv.issueDate)}</td>
                    <td style={styles.td}>{fmt(inv.dueDate)}</td>
                    <td style={{ ...styles.td, fontWeight: 700 }}>{currency(total)}</td>
                    <td style={styles.td}><span style={{ fontSize: 11, fontWeight: 700, color: sc[inv.status], background: sc[inv.status] + "20", padding: "3px 8px", borderRadius: 4 }}>{inv.status.toUpperCase()}</span></td>
                    <td style={styles.td}><Btn size="sm" variant="secondary" onClick={() => { if (client) previewInvoice(inv, client); }}>Preview</Btn></td>
                  </tr>
                );
              })}
              {filtered3.length === 0 && <tr><td colSpan={7} style={{ ...styles.td, textAlign: "center", color: "#aaa", padding: 32 }}>No invoices found.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    );

    function previewInvoice(inv, client) {
      const html = generateInvoiceHTML(inv, client);
      const w = window.open("", "_blank");
      w.document.write(html);
      w.document.close();
    }
  };

  // ─── Layout ────────────────────────────────────────────────────────────────
  if (loading) return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh", fontFamily: "'DM Serif Display', serif", fontSize: 22, color: "#888" }}>
      Loading…
    </div>
  );

  return (
    <>
      <link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet" />
      <div style={{ display: "flex", minHeight: "100vh", background: "#f7f7fc", fontFamily: "'DM Sans', sans-serif" }}>
        {/* Sidebar */}
        <div style={{ width: 220, background: "#fff", borderRight: "1px solid #e8e8f0", padding: "24px 12px", display: "flex", flexDirection: "column", flexShrink: 0, position: "sticky", top: 0, height: "100vh" }}>
          <div style={{ padding: "0 6px 24px", borderBottom: "1px solid #eee", marginBottom: 16 }}>
            <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 20, color: "#1a1a2e", lineHeight: 1.1 }}>Agency</div>
            <div style={{ fontSize: 10, letterSpacing: "0.1em", textTransform: "uppercase", color: "#aaa", marginTop: 3 }}>Management Hub</div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 3, flex: 1 }}>
            {nav("dashboard", "Dashboard", "◼")}
            {nav("clients", "Clients", "👥")}
            {nav("renewals", "Renewals", "🔄")}
            {nav("invoices", "Invoices", "🧾")}
          </div>
          {(urgent.length > 0 || overdue.length > 0) && (
            <div style={{ background: "#fff7ed", border: "1px solid #fed7aa", borderRadius: 8, padding: "10px 12px", marginTop: 16 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#f97316", marginBottom: 4 }}>⚡ Attention Needed</div>
              {overdue.length > 0 && <div style={{ fontSize: 11, color: "#ef4444" }}>{overdue.length} overdue renewal{overdue.length > 1 ? "s" : ""}</div>}
              {urgent.length > 0 && <div style={{ fontSize: 11, color: "#f97316" }}>{urgent.length} renewal{urgent.length > 1 ? "s" : ""} within 30 days</div>}
            </div>
          )}
          {/* Export / Import */}
          <div style={{ borderTop: "1px solid #eee", paddingTop: 14, marginTop: 14, display: "flex", flexDirection: "column", gap: 6 }}>
            <div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.1em", color: "#bbb", marginBottom: 2, paddingLeft: 6 }}>Data</div>
            <button onClick={handleExport}
              style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 14px", borderRadius: 7, border: "1px solid #e0e0ec", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 600, background: "#f9f9fc", color: "#1a1a2e", width: "100%", transition: "all 0.15s" }}>
              <span>⬇</span> Export to XLS
            </button>
            <button onClick={() => importRef.current?.click()}
              style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 14px", borderRadius: 7, border: "1px solid #e0e0ec", cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 600, background: "#f9f9fc", color: "#1a1a2e", width: "100%", transition: "all 0.15s" }}>
              <span>⬆</span> Import from XLS
            </button>
            <input ref={importRef} type="file" accept=".xlsx,.xls" onChange={handleImportFile} style={{ display: "none" }} />
          </div>
        </div>

        {/* Main content */}
        <div style={{ flex: 1, padding: "32px 36px", maxWidth: "calc(100vw - 220px)", overflowX: "auto" }}>
          {view === "dashboard" && <Dashboard />}
          {view === "clients" && <ClientsList />}
          {view === "client-detail" && selectedClient && <ClientDetail />}
          {view === "renewals" && <RenewalsView />}
          {view === "invoices" && <InvoicesView />}
        </div>
      </div>

      {/* Import Confirm Modal */}
      {importConfirm && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(10,10,30,0.55)", zIndex: 300, display: "flex", alignItems: "center", justifyContent: "center", padding: 16, backdropFilter: "blur(2px)" }}>
          <div style={{ background: "#fff", borderRadius: 12, width: "100%", maxWidth: 440, boxShadow: "0 24px 80px rgba(0,0,0,0.25)", padding: 28 }}>
            <div style={{ fontFamily: "'DM Serif Display', serif", fontSize: 20, color: "#1a1a2e", marginBottom: 10 }}>Import Data</div>
            <p style={{ fontSize: 14, color: "#555", lineHeight: 1.6, marginBottom: 18 }}>
              Found <strong>{importConfirm.length} client(s)</strong> in the file. How would you like to import?
            </p>
            <div style={{ display: "grid", gap: 10, marginBottom: 20 }}>
              <div style={{ border: "1px solid #e0e0ec", borderRadius: 8, padding: "14px 16px", cursor: "pointer", background: "#f9f9fc" }} onClick={() => confirmImport("merge")}>
                <div style={{ fontWeight: 700, fontSize: 13, color: "#1a1a2e", marginBottom: 3 }}>⊕ Merge with existing data</div>
                <div style={{ fontSize: 12, color: "#888" }}>New clients from the file will be added. Existing clients (matched by name) won't be duplicated.</div>
              </div>
              <div style={{ border: "1px solid #fecaca", borderRadius: 8, padding: "14px 16px", cursor: "pointer", background: "#fef2f2" }} onClick={() => confirmImport("replace")}>
                <div style={{ fontWeight: 700, fontSize: 13, color: "#ef4444", marginBottom: 3 }}>↺ Replace all data</div>
                <div style={{ fontSize: 12, color: "#888" }}>All current data will be deleted and replaced with the imported file. This cannot be undone.</div>
              </div>
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end" }}>
              <button onClick={() => setImportConfirm(null)} style={{ background: "#f0f0f8", color: "#1a1a2e", border: "none", borderRadius: 6, padding: "8px 16px", cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 600 }}>Cancel</button>
            </div>
          </div>
        </div>
      )}

      {/* Toast */}
      {toast && (
        <div style={{ position: "fixed", bottom: 24, right: 24, background: toast.type === "error" ? "#1a1a2e" : "#1a1a2e", color: "#fff", padding: "12px 20px", borderRadius: 8, fontSize: 13, fontWeight: 500, boxShadow: "0 8px 24px rgba(0,0,0,0.2)", zIndex: 999, display: "flex", alignItems: "center", gap: 8, animation: "slideUp 0.2s ease" }}>
          <span>{toast.type === "error" ? "🗑" : "✓"}</span> {toast.msg}
        </div>
      )}
      <style>{`@keyframes slideUp{from{transform:translateY(12px);opacity:0}to{transform:translateY(0);opacity:1}}`}</style>
    </>
  );
}
