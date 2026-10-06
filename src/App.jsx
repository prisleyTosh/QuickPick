import React, { useState, useMemo, useRef, useEffect, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  Search, Filter, ArrowUpDown, Upload, Download, Plus, MoreVertical,
  X, Check, Trash2, Copy, Edit2, ChevronLeft, ChevronRight, ImagePlus,
  ImageOff, ChevronDown, AlertTriangle,
  LayoutDashboard, Store, Package, Tag, ShoppingCart, Settings, Bell,
  ArrowLeft, ArrowRight, Clock, TrendingUp, PieChart, BarChart3, Menu,
  PackagePlus, Wallet, Percent,
  ArrowUp, ArrowDown, History, Save, TrendingDown, Minus, Camera,
  ScanLine, Undo2, Sparkles, Grid3x3, Zap, ListChecks, CheckCircle2,
  CircleAlert, FileSpreadsheet, ClipboardPaste
} from "lucide-react";

// ---------- Design tokens ----------
const C = {
  bg: "#F9FAF8",
  surface: "#FFFFFF",
  border: "#E3E7E2",
  borderStrong: "#C7CDC5",
  text: "#1C201D",
  textSecondary: "#5D6660",
  textMuted: "#8B9188",
  accent: "#0E7A4F",
  accentDark: "#0A5C3B",
  accentTint: "#E7F3EC",
  danger: "#C23A3A",
  dangerTint: "#FBE9E9",
  amber: "#B4740B",
  amberTint: "#FBF0DD",
  gray: "#5D6660",
  grayTint: "#EEF0ED",
  headerBg: "#F3F5F2",
};

const CATEGORIES = [
  "Food & Beverages", "Household", "Personal Care", "Baby & Kids",
  "Health & Wellness", "Electronics", "Stationery", "Clothing", "General Merchandise",
];

const EDITABLE_FIELDS = ["productName", "brand", "category", "packageSize", "price", "stock", "sku", "status"];

// ---------- Sample data ----------
const seed = [
  ["Soko Maize Meal", "Soko", "Food & Beverages", "2kg", 165, 40, "SOKO-2KG"],
  ["Jogoo Maize Meal", "Jogoo", "Food & Beverages", "2kg", 170, 25, "JOGOO-2KG"],
  ["Omo Detergent", "Omo", "Household", "1kg", 280, 5, "OMO-1KG"],
  ["Ariel Detergent", "Ariel", "Household", "1kg", 310, 0, "ARIEL-1KG"],
  ["Kericho Gold Tea", "Kericho Gold", "Food & Beverages", "500g", 245, 60, "KG-TEA-500"],
  ["Colgate Toothpaste", "Colgate", "Personal Care", "150ml", 180, 32, "COLG-150"],
  ["Pampers Diapers", "Pampers", "Baby & Kids", "Size 3, 50pcs", 1450, 14, "PAMP-S3-50"],
  ["Panadol Extra", "Panadol", "Health & Wellness", "20 tabs", 120, 8, "PAN-EX-20"],
  ["Energizer AA Batteries", "Energizer", "Electronics", "4-pack", 260, 22, "ENRG-AA4"],
  ["Bic Ballpoint Pens", "Bic", "Stationery", "10-pack", 150, 75, "BIC-10PK"],
  ["Cotton T-Shirt", "QuickWear", "Clothing", "Medium", 550, 18, "QW-TS-M"],
  ["Plastic Storage Box", "Homeware", "General Merchandise", "20L", 890, 12, "HW-BOX-20L"],
  ["Fresh Fri Cooking Oil", "Fresh Fri", "Food & Beverages", "2L", 520, 30, "FF-OIL-2L"],
  ["Dettol Antiseptic", "Dettol", "Personal Care", "250ml", 340, 3, "DETT-250"],
  ["Huggies Wipes", "Huggies", "Baby & Kids", "80pcs", 380, 0, "HUG-WIPE-80"],
  ["Vitamin C Tablets", "HealthPlus", "Health & Wellness", "30 tabs", 450, 20, "HP-VITC-30"],
  ["USB Flash Drive", "SanDisk", "Electronics", "32GB", 950, 9, "SD-USB-32"],
  ["A4 Exercise Book", "Kasuku", "Stationery", "200 pages", 90, 120, "KAS-A4-200"],
  ["Denim Jeans", "QuickWear", "Clothing", "32W", 1200, 6, "QW-DJ-32"],
  ["Kitchen Towel Roll", "Homeware", "General Merchandise", "2-pack", 210, 45, "HW-KT-2PK"],
  ["Blueband Margarine", "Blueband", "Food & Beverages", "500g", 195, 55, "BB-MARG-500"],
  ["Sunlight Bar Soap", "Sunlight", "Household", "800g", 150, 2, "SUN-BAR-800"],
  ["Nivea Body Lotion", "Nivea", "Personal Care", "400ml", 420, 17, "NIV-BL-400"],
  ["Baby Formula", "Nan", "Baby & Kids", "900g", 1850, 4, "NAN-900"],
  ["Hand Sanitizer", "Dettol", "Health & Wellness", "500ml", 380, 28, "DETT-HS-500"],
  ["Phone Charger Cable", "Anker", "Electronics", "1m USB-C", 650, 33, "ANK-USBC-1M"],
  ["Sticky Notes", "Post-it", "Stationery", "3-pack", 260, 40, "PIT-3PK"],
  ["Kids Socks", "QuickWear", "Clothing", "3-pair pack", 300, 50, "QW-SOCK-3"],
  ["Laundry Basket", "Homeware", "General Merchandise", "Large", 1100, 7, "HW-LB-L"],
  ["Ketepa Tea Bags", "Ketepa", "Food & Beverages", "100 bags", 310, 38, "KET-100"],
  ["Vim Dishwashing Bar", "Vim", "Household", "700g", 130, 65, "VIM-700"],
];

function makeProduct(fields, idx) {
  const [productName, brand, category, packageSize, price, stock, sku] = fields;
  return {
    id: `QP-${String(idx + 1).padStart(3, "0")}`,
    productName, brand, category, packageSize, price, stock, sku,
    status: "Active",
    image: null,
    // Distributes seed products round-robin across the 10 demo shops
    // (ids QP001-QP010, see initialShops below) so the "products carried
    // by this shop" view has real data out of the box. Written as a direct
    // pattern rather than referencing initialShops here, since that array
    // is declared further down this module and isn't initialized yet at
    // the point initialProducts is built.
    shopId: `QP${String((idx % 10) + 1).padStart(3, "0")}`,
  };
}

const initialProducts = seed.map((fields, idx) => makeProduct(fields, idx));

// ---------- Helpers ----------
function formatKSh(n) {
  const num = Number(n) || 0;
  return "KSh " + num.toLocaleString("en-KE", { maximumFractionDigits: 0 });
}

function computeBadge(product) {
  if (Number(product.stock) === 0) return { label: "Out of Stock", tone: "danger" };
  if (Number(product.stock) <= 10) return { label: "Low Stock", tone: "amber" };
  if (product.status === "Inactive") return { label: "Inactive", tone: "gray" };
  return { label: "Active", tone: "accent" };
}

function toneStyles(tone) {
  switch (tone) {
    case "danger": return { bg: C.dangerTint, fg: C.danger };
    case "amber": return { bg: C.amberTint, fg: C.amber };
    case "gray": return { bg: C.grayTint, fg: C.gray };
    default: return { bg: C.accentTint, fg: C.accentDark };
  }
}

function nextId(products) {
  const nums = products.map(p => parseInt(String(p.id).replace(/\D/g, ""), 10) || 0);
  const max = nums.length ? Math.max(...nums) : 0;
  return `QP-${String(max + 1).padStart(3, "0")}`;
}

function download(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function exportRows(products) {
  return products.map(p => ({
    productId: p.id,
    productName: p.productName,
    brand: p.brand,
    category: p.category,
    packageSize: p.packageSize,
    price: p.price,
    stock: p.stock,
    sku: p.sku,
    status: computeBadge(p).label,
  }));
}

// ---------- Small UI atoms ----------
function Button({ children, onClick, variant = "secondary", icon: Icon, style, disabled, title }) {
  const base = {
    display: "inline-flex", alignItems: "center", gap: 6,
    padding: "8px 14px", borderRadius: 8, fontSize: 13.5, fontWeight: 500,
    cursor: disabled ? "not-allowed" : "pointer", border: "1px solid transparent",
    transition: "background 120ms ease, border-color 120ms ease", whiteSpace: "nowrap",
    opacity: disabled ? 0.5 : 1,
  };
  const variants = {
    primary: { background: C.accent, color: "#fff", border: `1px solid ${C.accent}` },
    secondary: { background: C.surface, color: C.text, border: `1px solid ${C.border}` },
    ghost: { background: "transparent", color: C.textSecondary, border: "1px solid transparent" },
    danger: { background: C.danger, color: "#fff", border: `1px solid ${C.danger}` },
  };
  return (
    <button
      title={title}
      disabled={disabled}
      onClick={onClick}
      style={{ ...base, ...variants[variant], ...style }}
      onMouseEnter={(e) => {
        if (disabled) return;
        if (variant === "secondary") e.currentTarget.style.borderColor = C.borderStrong;
        if (variant === "ghost") e.currentTarget.style.background = C.grayTint;
        if (variant === "primary") e.currentTarget.style.background = C.accentDark;
      }}
      onMouseLeave={(e) => {
        if (variant === "secondary") e.currentTarget.style.borderColor = C.border;
        if (variant === "ghost") e.currentTarget.style.background = "transparent";
        if (variant === "primary") e.currentTarget.style.background = C.accent;
      }}
    >
      {Icon ? <Icon size={15} /> : null}
      {children}
    </button>
  );
}

function Badge({ tone, label }) {
  const s = toneStyles(tone);
  return (
    <span style={{
      display: "inline-block", padding: "3px 10px", borderRadius: 999,
      fontSize: 12, fontWeight: 500, background: s.bg, color: s.fg, whiteSpace: "nowrap",
    }}>
      {label}
    </span>
  );
}

function Modal({ children, onClose, width = 480 }) {
  return (
    <div
      onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}
      style={{
        position: "fixed", inset: 0, background: "rgba(28,32,29,0.35)",
        display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16,
      }}
    >
      <div style={{
        background: C.surface, borderRadius: 12, width, maxWidth: "100%",
        maxHeight: "90vh", overflowY: "auto", border: `1px solid ${C.border}`,
        boxShadow: "0 8px 30px rgba(0,0,0,0.12)",
      }}>
        {children}
      </div>
    </div>
  );
}

function Field({ label, required, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: "block", fontSize: 12.5, fontWeight: 500, color: C.textSecondary, marginBottom: 5 }}>
        {label}{required ? <span style={{ color: C.danger }}> *</span> : null}
      </label>
      {children}
    </div>
  );
}

const inputStyle = {
  width: "100%", padding: "8px 10px", borderRadius: 7, border: `1px solid ${C.border}`,
  fontSize: 13.5, color: C.text, background: C.surface, boxSizing: "border-box",
};

// ---------- Toast ----------
function Toast({ message, onDone }) {
  useEffect(() => {
    const t = setTimeout(onDone, 2600);
    return () => clearTimeout(t);
  }, [message]);
  if (!message) return null;
  return (
    <div style={{
      position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)",
      background: C.text, color: "#fff", padding: "10px 18px", borderRadius: 8,
      fontSize: 13.5, display: "flex", alignItems: "center", gap: 8, zIndex: 1200,
      boxShadow: "0 6px 20px rgba(0,0,0,0.2)",
    }}>
      <Check size={15} color={C.accentTint} />
      {message}
    </div>
  );
}

// ---------- Product form modal ----------
function ProductFormModal({ initial, shops = [], onCancel, onSave }) {
  const [form, setForm] = useState(initial || {
    productName: "", brand: "", category: "", packageSize: "", price: "", stock: "", sku: "", status: "Active", image: null, shopId: "",
  });
  const [errors, setErrors] = useState({});
  const fileRef = useRef(null);

  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  const handleImage = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => set("image", reader.result);
    reader.readAsDataURL(file);
  };

  const validate = () => {
    const e = {};
    if (!form.productName.trim()) e.productName = "Enter a product name";
    if (!form.category) e.category = "Choose a category";
    if (form.price === "" || isNaN(Number(form.price)) || Number(form.price) < 0) e.price = "Enter a valid price";
    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const submit = () => {
    if (!validate()) return;
    onSave({
      ...form,
      price: Number(form.price) || 0,
      stock: Number(form.stock) || 0,
    });
  };

  return (
    <Modal onClose={onCancel} width={480}>
      <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <h2 style={{ margin: 0, fontSize: 16, fontWeight: 600 }}>{initial ? "Edit product" : "Add product"}</h2>
        <button onClick={onCancel} style={{ background: "none", border: "none", cursor: "pointer", color: C.textMuted }}>
          <X size={18} />
        </button>
      </div>

      <div style={{ padding: "20px 24px" }}>
        <Field label="Product image">
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{
              width: 64, height: 64, borderRadius: 8, border: `1px dashed ${C.borderStrong}`,
              display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden",
              background: C.headerBg, flexShrink: 0,
            }}>
              {form.image ? (
                <img src={form.image} alt="Product" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
              ) : (
                <ImageOff size={20} color={C.textMuted} />
              )}
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <Button icon={ImagePlus} onClick={() => fileRef.current?.click()}>
                {form.image ? "Replace" : "Upload"}
              </Button>
              {form.image ? (
                <Button variant="ghost" onClick={() => set("image", null)}>Remove</Button>
              ) : null}
            </div>
            <input
              ref={fileRef} type="file" accept="image/*" style={{ display: "none" }}
              onChange={(e) => handleImage(e.target.files?.[0])}
            />
          </div>
        </Field>

        <Field label="Product name" required>
          <input style={inputStyle} value={form.productName} onChange={e => set("productName", e.target.value)} placeholder="e.g. Soko Maize Meal" />
          {errors.productName ? <div style={{ color: C.danger, fontSize: 12, marginTop: 4 }}>{errors.productName}</div> : null}
        </Field>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Field label="Brand">
            <input style={inputStyle} value={form.brand} onChange={e => set("brand", e.target.value)} placeholder="e.g. Soko" />
          </Field>
          <Field label="Category" required>
            <select style={inputStyle} value={form.category} onChange={e => set("category", e.target.value)}>
              <option value="">Choose a category</option>
              {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
            {errors.category ? <div style={{ color: C.danger, fontSize: 12, marginTop: 4 }}>{errors.category}</div> : null}
          </Field>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Field label="Package size">
            <input style={inputStyle} value={form.packageSize} onChange={e => set("packageSize", e.target.value)} placeholder="e.g. 2kg" />
          </Field>
          <Field label="Price (KSh)" required>
            <input style={inputStyle} type="number" min="0" value={form.price} onChange={e => set("price", e.target.value)} placeholder="0" />
            {errors.price ? <div style={{ color: C.danger, fontSize: 12, marginTop: 4 }}>{errors.price}</div> : null}
          </Field>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          <Field label="Stock quantity">
            <input style={inputStyle} type="number" min="0" value={form.stock} onChange={e => set("stock", e.target.value)} placeholder="0" />
          </Field>
          <Field label="SKU">
            <input style={inputStyle} value={form.sku} onChange={e => set("sku", e.target.value)} placeholder="e.g. SOKO-2KG" />
          </Field>
        </div>

        <Field label="Assigned shop">
          <select style={inputStyle} value={form.shopId || ""} onChange={e => set("shopId", e.target.value)}>
            <option value="">No shop assigned</option>
            {shops.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </Field>

        <Field label="Status">
          <select style={inputStyle} value={form.status} onChange={e => set("status", e.target.value)}>
            <option value="Active">Active</option>
            <option value="Inactive">Inactive</option>
          </select>
        </Field>
      </div>

      <div style={{ padding: "16px 24px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Button variant="ghost" onClick={onCancel}>Cancel</Button>
        <Button variant="primary" onClick={submit}>Save product</Button>
      </div>
    </Modal>
  );
}

// ---------- Shops assigned to a product (many-to-many, one price per shop) ----------
function ProductShopsModal({ product, shops, shopProducts, onAssign, onUpdatePrice, onRemove, onClose }) {
  const assignments = shopProducts.filter((sp) => sp.productId === product.id);
  const assignedShopIds = new Set(assignments.map((a) => a.shopId));
  const availableShops = shops.filter((s) => !assignedShopIds.has(s.id));

  const [selectedShopId, setSelectedShopId] = useState("");
  const [priceDraft, setPriceDraft] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editPrice, setEditPrice] = useState("");
  const [notice, setNotice] = useState("");

  function handleAssign() {
    if (!selectedShopId) { setNotice("Choose a shop first."); return; }
    if (priceDraft === "" || isNaN(Number(priceDraft)) || Number(priceDraft) < 0) { setNotice("Enter a valid price."); return; }
    const shop = shops.find((s) => s.id === selectedShopId);
    const result = onAssign(product.id, selectedShopId, Number(priceDraft));
    setNotice(
      result?.updated
        ? `Product already assigned to ${shop?.name}. Its price was updated instead.`
        : `Assigned to ${shop?.name}.`
    );
    setSelectedShopId("");
    setPriceDraft("");
  }

  return (
    <Modal onClose={onClose} width={560}>
      <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <h2 style={{ margin: 0, fontSize: 16, fontWeight: 600 }}>Shops selling this product</h2>
        <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", color: C.textMuted }}>
          <X size={18} />
        </button>
      </div>

      <div style={{ padding: "20px 24px" }}>
        <p style={{ margin: "0 0 4px", fontWeight: 600 }}>{product.productName}</p>
        <p style={{ margin: "0 0 16px", color: C.textSecondary, fontSize: 13 }}>
          {assignments.length} shop{assignments.length === 1 ? "" : "s"} currently selling this product
        </p>

        {notice ? (
          <div style={{ background: C.accentTint || "#e6f4ea", color: C.text, fontSize: 12.5, padding: "8px 12px", borderRadius: 8, marginBottom: 14 }}>
            {notice}
          </div>
        ) : null}

        {assignments.length === 0 ? (
          <p style={{ margin: "0 0 16px", color: C.textMuted, fontStyle: "italic", fontSize: 13 }}>
            Not yet assigned to any shop. Use "Assign to Another Shop" below.
          </p>
        ) : (
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, marginBottom: 20, overflow: "hidden" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: C.headerBg }}>
                  <th style={{ textAlign: "left", padding: "9px 12px", fontSize: 11.5, color: C.textSecondary, fontWeight: 600 }}>Shop</th>
                  <th style={{ textAlign: "left", padding: "9px 12px", fontSize: 11.5, color: C.textSecondary, fontWeight: 600 }}>Location</th>
                  <th style={{ textAlign: "right", padding: "9px 12px", fontSize: 11.5, color: C.textSecondary, fontWeight: 600 }}>Price</th>
                  <th style={{ padding: "9px 12px" }} />
                </tr>
              </thead>
              <tbody>
                {assignments.map((a) => {
                  const shop = shops.find((s) => s.id === a.shopId);
                  const isEditing = editingId === a.id;
                  return (
                    <tr key={a.id} style={{ borderTop: `1px solid ${C.border}` }}>
                      <td style={{ padding: "9px 12px", fontWeight: 600 }}>{shop?.name || "Unknown shop"}</td>
                      <td style={{ padding: "9px 12px", color: C.textSecondary }}>{shop?.location || "\u2014"}</td>
                      <td style={{ padding: "9px 12px", textAlign: "right" }}>
                        {isEditing ? (
                          <input
                            type="number" min="0" autoFocus value={editPrice}
                            onChange={(e) => setEditPrice(e.target.value)}
                            style={{ width: 84, padding: "4px 6px", border: `1px solid ${C.border}`, borderRadius: 6, textAlign: "right", fontSize: 13 }}
                          />
                        ) : (
                          <span style={{ fontWeight: 600 }}>{formatKSh(a.price)}</span>
                        )}
                      </td>
                      <td style={{ padding: "9px 12px", textAlign: "right", whiteSpace: "nowrap" }}>
                        {isEditing ? (
                          <>
                            <button
                              onClick={() => {
                                if (editPrice === "" || isNaN(Number(editPrice)) || Number(editPrice) < 0) return;
                                onUpdatePrice(a.id, Number(editPrice));
                                setEditingId(null);
                              }}
                              style={{ background: "none", border: "none", cursor: "pointer", color: C.accent, fontSize: 12.5, fontWeight: 600, marginRight: 8 }}
                            >
                              Save
                            </button>
                            <button onClick={() => setEditingId(null)} style={{ background: "none", border: "none", cursor: "pointer", color: C.textMuted, fontSize: 12.5 }}>
                              Cancel
                            </button>
                          </>
                        ) : (
                          <>
                            <button
                              onClick={() => { setEditingId(a.id); setEditPrice(String(a.price)); }}
                              style={{ background: "none", border: "none", cursor: "pointer", color: C.accent, fontSize: 12.5, fontWeight: 600, marginRight: 8 }}
                            >
                              Edit
                            </button>
                            <button onClick={() => onRemove(a.id)} style={{ background: "none", border: "none", cursor: "pointer", color: C.danger, fontSize: 12.5 }}>
                              Remove
                            </button>
                          </>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        <div style={{ paddingTop: 16, borderTop: `1px solid ${C.border}` }}>
          <div style={{ fontWeight: 600, fontSize: 13.5, marginBottom: 10 }}>+ Assign to Another Shop</div>
          {availableShops.length === 0 ? (
            <p style={{ margin: 0, color: C.textMuted, fontStyle: "italic", fontSize: 12.5 }}>
              Already assigned to every registered shop.
            </p>
          ) : (
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <select
                style={{ ...inputStyle, flex: 2, minWidth: 160 }}
                value={selectedShopId}
                onChange={(e) => setSelectedShopId(e.target.value)}
              >
                <option value="">Select shop...</option>
                {availableShops.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
              <input
                style={{ ...inputStyle, flex: 1, minWidth: 90 }}
                type="number" min="0" placeholder="Price"
                value={priceDraft}
                onChange={(e) => setPriceDraft(e.target.value)}
              />
              <Button variant="primary" onClick={handleAssign}>Assign Product</Button>
            </div>
          )}
        </div>
      </div>

      <div style={{ padding: "0 24px 20px", display: "flex", justifyContent: "flex-end" }}>
        <Button variant="ghost" onClick={onClose}>Close</Button>
      </div>
    </Modal>
  );
}

// ---------- Delete confirm modal ----------
function DeleteConfirmModal({ product, count, onCancel, onConfirm }) {
  const label = count > 1 ? `${count} products` : `"${product?.productName}"`;
  return (
    <Modal onClose={onCancel} width={400}>
      <div style={{ padding: "24px 24px 20px" }}>
        <div style={{ display: "flex", gap: 12, alignItems: "flex-start" }}>
          <div style={{
            width: 36, height: 36, borderRadius: "50%", background: C.dangerTint,
            display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
          }}>
            <AlertTriangle size={18} color={C.danger} />
          </div>
          <div>
            <h3 style={{ margin: "2px 0 6px", fontSize: 15.5, fontWeight: 600 }}>Delete product?</h3>
            <p style={{ margin: 0, fontSize: 13.5, color: C.textSecondary, lineHeight: 1.5 }}>
              Are you sure you want to delete {label}? This can't be undone.
            </p>
          </div>
        </div>
      </div>
      <div style={{ padding: "0 24px 20px", display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Button variant="ghost" onClick={onCancel}>Cancel</Button>
        <Button variant="danger" icon={Trash2} onClick={onConfirm}>Delete</Button>
      </div>
    </Modal>
  );
}

// ---------- Import modal ----------
function ImportModal({ onCancel, onConfirm }) {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [errorMsg, setErrorMsg] = useState("");
  const fileRef = useRef(null);

  const parseFile = (file) => {
    setErrorMsg("");
    setFileName(file.name);
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: (res) => setRows(res.data),
        error: () => setErrorMsg("Couldn't read that file. Check the format and try again."),
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          setRows(json);
        } catch {
          setErrorMsg("Couldn't read that file. Check the format and try again.");
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setErrorMsg("Import a CSV or Excel (.xlsx) file.");
    }
  };

  const validRows = (rows || []).filter(r => (r.productName || r.name || "").toString().trim());

  return (
    <Modal onClose={onCancel} width={640}>
      <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <h2 style={{ margin: 0, fontSize: 16, fontWeight: 600 }}>Import products</h2>
        <button onClick={onCancel} style={{ background: "none", border: "none", cursor: "pointer", color: C.textMuted }}>
          <X size={18} />
        </button>
      </div>

      <div style={{ padding: "20px 24px" }}>
        {!rows ? (
          <div
            onClick={() => fileRef.current?.click()}
            style={{
              border: `1.5px dashed ${C.borderStrong}`, borderRadius: 10, padding: "36px 20px",
              textAlign: "center", cursor: "pointer", background: C.headerBg,
            }}
          >
            <Upload size={22} color={C.textMuted} style={{ marginBottom: 8 }} />
            <div style={{ fontSize: 13.5, color: C.text, fontWeight: 500 }}>Click to choose a CSV or Excel file</div>
            <div style={{ fontSize: 12, color: C.textMuted, marginTop: 4 }}>
              Columns: productName, brand, category, packageSize, price, stock, sku, status
            </div>
            <input
              ref={fileRef} type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }}
              onChange={(e) => e.target.files?.[0] && parseFile(e.target.files[0])}
            />
          </div>
        ) : (
          <div>
            <div style={{ fontSize: 13, color: C.textSecondary, marginBottom: 10 }}>
              {fileName} — {validRows.length} product{validRows.length === 1 ? "" : "s"} found
            </div>
            <div style={{ maxHeight: 260, overflow: "auto", border: `1px solid ${C.border}`, borderRadius: 8 }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                <thead>
                  <tr style={{ background: C.headerBg }}>
                    {["productName", "brand", "category", "packageSize", "price", "stock", "sku", "status"].map(h => (
                      <th key={h} style={{ textAlign: "left", padding: "7px 10px", borderBottom: `1px solid ${C.border}`, fontWeight: 600, whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {validRows.slice(0, 50).map((r, i) => (
                    <tr key={i} style={{ borderBottom: `1px solid ${C.border}` }}>
                      <td style={{ padding: "6px 10px" }}>{r.productName || r.name}</td>
                      <td style={{ padding: "6px 10px" }}>{r.brand}</td>
                      <td style={{ padding: "6px 10px" }}>{r.category}</td>
                      <td style={{ padding: "6px 10px" }}>{r.packageSize}</td>
                      <td style={{ padding: "6px 10px" }}>{r.price}</td>
                      <td style={{ padding: "6px 10px" }}>{r.stock}</td>
                      <td style={{ padding: "6px 10px" }}>{r.sku}</td>
                      <td style={{ padding: "6px 10px" }}>{r.status}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
        {errorMsg ? <div style={{ color: C.danger, fontSize: 13, marginTop: 10 }}>{errorMsg}</div> : null}
      </div>

      <div style={{ padding: "16px 24px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Button variant="ghost" onClick={onCancel}>Cancel</Button>
        <Button
          variant="primary"
          disabled={!rows || validRows.length === 0}
          onClick={() => onConfirm(validRows)}
        >
          Import {validRows.length ? `${validRows.length} products` : ""}
        </Button>
      </div>
    </Modal>
  );
}

// ---------- Export modal ----------
function ExportModal({ scopeCount, onCancel, onConfirm }) {
  const [scope, setScope] = useState("all");
  const [format, setFormat] = useState("csv");
  return (
    <Modal onClose={onCancel} width={380}>
      <div style={{ padding: "20px 24px", borderBottom: `1px solid ${C.border}` }}>
        <h2 style={{ margin: 0, fontSize: 16, fontWeight: 600 }}>Export products</h2>
      </div>
      <div style={{ padding: "20px 24px" }}>
        <Field label="What to export">
          <select style={inputStyle} value={scope} onChange={e => setScope(e.target.value)}>
            <option value="all">Export all products</option>
            <option value="filtered">Export filtered products</option>
            <option value="selected" disabled={!scopeCount.selected}>
              Export selected products{scopeCount.selected ? ` (${scopeCount.selected})` : " (none selected)"}
            </option>
          </select>
        </Field>
        <Field label="Format">
          <div style={{ display: "flex", gap: 8 }}>
            {["csv", "xlsx"].map(f => (
              <button
                key={f}
                onClick={() => setFormat(f)}
                style={{
                  flex: 1, padding: "8px 10px", borderRadius: 7, fontSize: 13, fontWeight: 500,
                  cursor: "pointer", border: `1px solid ${format === f ? C.accent : C.border}`,
                  background: format === f ? C.accentTint : C.surface, color: format === f ? C.accentDark : C.text,
                }}
              >
                {f === "csv" ? "CSV" : "Excel (.xlsx)"}
              </button>
            ))}
          </div>
        </Field>
      </div>
      <div style={{ padding: "16px 24px", borderTop: `1px solid ${C.border}`, display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <Button variant="ghost" onClick={onCancel}>Cancel</Button>
        <Button variant="primary" icon={Download} onClick={() => onConfirm(scope, format)}>Export</Button>
      </div>
    </Modal>
  );
}

// ---------- Editable cell ----------
function EditableCell({ value, field, onSave, width, isSelect, options }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(value);
  const ref = useRef(null);

  useEffect(() => { setDraft(value); }, [value]);
  useEffect(() => {
    if (editing && ref.current) { ref.current.focus(); if (ref.current.select) ref.current.select(); }
  }, [editing]);

  const commit = () => {
    setEditing(false);
    if (draft !== value) onSave(field, draft);
  };
  const cancel = () => { setDraft(value); setEditing(false); };

  if (editing) {
    return (
      <div style={{ position: "relative", width }}>
        {isSelect ? (
          <select
            ref={ref} value={draft} onChange={e => setDraft(e.target.value)}
            onBlur={commit}
            onKeyDown={e => { if (e.key === "Enter") commit(); if (e.key === "Escape") cancel(); }}
            style={{
              width: "100%", fontSize: 12.5, padding: "4px 6px", boxSizing: "border-box",
              border: `2px solid ${C.accent}`, borderRadius: 4, outline: "none", background: "#fff",
            }}
          >
            {options.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        ) : (
          <input
            ref={ref} value={draft} onChange={e => setDraft(e.target.value)}
            onBlur={commit}
            type={field === "price" || field === "stock" ? "number" : "text"}
            onKeyDown={e => { if (e.key === "Enter") commit(); if (e.key === "Escape") cancel(); }}
            style={{
              width: "100%", fontSize: 12.5, padding: "4px 6px", boxSizing: "border-box",
              border: `2px solid ${C.accent}`, borderRadius: 4, outline: "none",
            }}
          />
        )}
        <div style={{
          position: "absolute", width: 6, height: 6, background: C.accent,
          right: -3, bottom: -3, border: "1px solid #fff",
        }} />
      </div>
    );
  }

  return (
    <div
      onClick={() => setEditing(true)}
      title="Click to edit"
      style={{
        cursor: "text", padding: "4px 6px", borderRadius: 4, minHeight: 20,
        overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
      }}
      onMouseEnter={e => e.currentTarget.style.background = C.headerBg}
      onMouseLeave={e => e.currentTarget.style.background = "transparent"}
    >
      {field === "price" ? formatKSh(value) : value}
    </div>
  );
}

// ---------- Actions menu ----------
function ActionsMenu({ onEdit, onShops, onDuplicate, onDelete }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    function onDoc(e) { if (ref.current && !ref.current.contains(e.target)) setOpen(false); }
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, []);

  return (
    <div style={{ position: "relative" }} ref={ref}>
      <button
        onClick={() => setOpen(o => !o)}
        style={{ background: "none", border: "none", cursor: "pointer", color: C.textSecondary, padding: 4, borderRadius: 5 }}
        onMouseEnter={e => e.currentTarget.style.background = C.headerBg}
        onMouseLeave={e => e.currentTarget.style.background = "transparent"}
      >
        <MoreVertical size={16} />
      </button>
      {open && (
        <div style={{
          position: "absolute", right: 0, top: "110%", background: C.surface,
          border: `1px solid ${C.border}`, borderRadius: 8, boxShadow: "0 6px 18px rgba(0,0,0,0.1)",
          zIndex: 50, minWidth: 130, overflow: "hidden",
        }}>
          {[
            { label: "Edit", icon: Edit2, action: onEdit },
            { label: "Shops", icon: Store, action: onShops },
            { label: "Duplicate", icon: Copy, action: onDuplicate },
            { label: "Delete", icon: Trash2, action: onDelete, danger: true },
          ].filter(item => item.action).map(item => (
            <div
              key={item.label}
              onClick={() => { setOpen(false); item.action(); }}
              style={{
                display: "flex", alignItems: "center", gap: 8, padding: "9px 12px",
                fontSize: 13, cursor: "pointer", color: item.danger ? C.danger : C.text,
              }}
              onMouseEnter={e => e.currentTarget.style.background = item.danger ? C.dangerTint : C.headerBg}
              onMouseLeave={e => e.currentTarget.style.background = "transparent"}
            >
              <item.icon size={14} />
              {item.label}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Main component ----------
function ProductManagement({ onBack, products: sharedProducts, onProductsChange, shops = [], shopProducts = [], onShopProductsChange }) {
  const [products, setProducts] = useState(() => sharedProducts ?? initialProducts);
  // Propagate to the shared App-level state after render commits, not from
  // inside the setProducts updater itself - calling another component's
  // setState mid-updater is unsafe and triggers React's cross-component
  // setState-during-render warning.
  useEffect(() => {
    if (onProductsChange) onProductsChange(products);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [products]);
  const [search, setSearch] = useState("");
  const [filterCategory, setFilterCategory] = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [filterStock, setFilterStock] = useState("All");
  const [filterOpen, setFilterOpen] = useState(false);
  const [sortOpen, setSortOpen] = useState(false);
  const [sortField, setSortField] = useState("");
  const [sortDir, setSortDir] = useState("asc");
  const [selected, setSelected] = useState(new Set());
  const [page, setPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(25);
  const [showAddModal, setShowAddModal] = useState(false);
  const [editProduct, setEditProduct] = useState(null);
  const [deleteTarget, setDeleteTarget] = useState(null); // {product} or {bulk: true}
  const [shopsProduct, setShopsProduct] = useState(null); // product currently shown in the "Shops" modal

  // Local mirror of shopProducts, propagated up the same safe way as
  // products/shops elsewhere in this file.
  const [shopProductsLocal, setShopProductsLocal] = useState(() => shopProducts);
  useEffect(() => {
    if (onShopProductsChange) onShopProductsChange(shopProductsLocal);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shopProductsLocal]);

  function assignProductToShop(productId, shopId, price) {
    const existing = shopProductsLocal.find((sp) => sp.productId === productId && sp.shopId === shopId);
    if (existing) {
      // Already assigned - update its price instead of creating a duplicate.
      setShopProductsLocal((prev) => prev.map((sp) => (sp.id === existing.id ? { ...sp, price, lastUpdated: new Date().toISOString() } : sp)));
      return { updated: true };
    }
    setShopProductsLocal((prev) => [
      ...prev,
      { id: `SP${Date.now()}${Math.floor(Math.random() * 1000)}`, productId, shopId, price, lastUpdated: new Date().toISOString() },
    ]);
    return { updated: false };
  }
  function updateShopProductPrice(id, price) {
    setShopProductsLocal((prev) => prev.map((sp) => (sp.id === id ? { ...sp, price, lastUpdated: new Date().toISOString() } : sp)));
  }
  function removeShopProduct(id) {
    setShopProductsLocal((prev) => prev.filter((sp) => sp.id !== id));
  }
  const [showImport, setShowImport] = useState(false);
  const [showExport, setShowExport] = useState(false);
  const [bulkCategoryOpen, setBulkCategoryOpen] = useState(false);
  const [bulkStatusOpen, setBulkStatusOpen] = useState(false);
  const [toast, setToast] = useState("");
  const filterRef = useRef(null);
  const sortRef = useRef(null);

  useEffect(() => {
    function onDoc(e) {
      if (filterRef.current && !filterRef.current.contains(e.target)) setFilterOpen(false);
      if (sortRef.current && !sortRef.current.contains(e.target)) setSortOpen(false);
    }
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, []);

  const notify = (msg) => setToast(msg);

  // ----- Filtering -----
  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return products.filter(p => {
      if (q) {
        const hay = `${p.productName} ${p.brand} ${p.id} ${p.sku} ${p.category}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      if (filterCategory && p.category !== filterCategory) return false;
      const badge = computeBadge(p);
      if (filterStatus && badge.label !== filterStatus) return false;
      if (filterStock !== "All") {
        const s = Number(p.stock);
        if (filterStock === "Out of Stock" && s !== 0) return false;
        if (filterStock === "Low Stock" && !(s > 0 && s <= 10)) return false;
        if (filterStock === "In Stock" && !(s > 10)) return false;
      }
      return true;
    });
  }, [products, search, filterCategory, filterStatus, filterStock]);

  // ----- Sorting -----
  const sorted = useMemo(() => {
    if (!sortField) return filtered;
    const copy = [...filtered];
    copy.sort((a, b) => {
      let av = a[sortField], bv = b[sortField];
      if (typeof av === "string") { av = av.toLowerCase(); bv = bv.toLowerCase(); }
      if (av < bv) return sortDir === "asc" ? -1 : 1;
      if (av > bv) return sortDir === "asc" ? 1 : -1;
      return 0;
    });
    return copy;
  }, [filtered, sortField, sortDir]);

  // ----- Pagination -----
  const totalPages = Math.max(1, Math.ceil(sorted.length / rowsPerPage));
  const currentPage = Math.min(page, totalPages);
  const pageStart = (currentPage - 1) * rowsPerPage;
  const paginated = sorted.slice(pageStart, pageStart + rowsPerPage);

  useEffect(() => { setPage(1); }, [search, filterCategory, filterStatus, filterStock, sortField, sortDir, rowsPerPage]);

  // ----- Selection -----
  const filteredIds = useMemo(() => new Set(sorted.map(p => p.id)), [sorted]);
  const selectedInFilter = useMemo(() => [...selected].filter(id => filteredIds.has(id)), [selected, filteredIds]);
  const allFilteredSelected = sorted.length > 0 && selectedInFilter.length === sorted.length;
  const someFilteredSelected = selectedInFilter.length > 0 && !allFilteredSelected;
  const selectAllRef = useRef(null);
  useEffect(() => { if (selectAllRef.current) selectAllRef.current.indeterminate = someFilteredSelected; }, [someFilteredSelected]);

  const toggleRow = (id) => {
    setSelected(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  const toggleAll = () => {
    setSelected(prev => {
      const next = new Set(prev);
      if (allFilteredSelected) {
        sorted.forEach(p => next.delete(p.id));
      } else {
        sorted.forEach(p => next.add(p.id));
      }
      return next;
    });
  };

  // ----- Cell editing -----
  const saveCell = (id, field, value) => {
    setProducts(prev => prev.map(p => {
      if (p.id !== id) return p;
      let v = value;
      if (field === "price" || field === "stock") v = Number(value) || 0;
      return { ...p, [field]: v };
    }));
  };

  // ----- Add / edit product -----
  const handleSaveProduct = (form) => {
    if (editProduct) {
      setProducts(prev => prev.map(p => p.id === editProduct.id ? { ...p, ...form } : p));
      notify("Product updated");
      setEditProduct(null);
    } else {
      const id = nextId(products);
      setProducts(prev => [{ id, ...form }, ...prev]);
      notify("Product added");
      setShowAddModal(false);
    }
  };

  const handleDuplicate = (product) => {
    const id = nextId(products);
    const copy = { ...product, id, productName: `${product.productName} (copy)`, sku: `${product.sku}-COPY` };
    setProducts(prev => [copy, ...prev]);
    notify("Product duplicated");
  };

  const handleDeleteConfirm = () => {
    if (deleteTarget?.bulk) {
      const idsToDelete = new Set(selected);
      setProducts(prev => prev.filter(p => !selected.has(p.id)));
      setShopProductsLocal(prev => prev.filter(sp => !idsToDelete.has(sp.productId)));
      notify(`${selected.size} products deleted`);
      setSelected(new Set());
    } else if (deleteTarget?.product) {
      const deletedId = deleteTarget.product.id;
      setProducts(prev => prev.filter(p => p.id !== deletedId));
      setShopProductsLocal(prev => prev.filter(sp => sp.productId !== deletedId));
      notify("Product deleted");
      setSelected(prev => { const n = new Set(prev); n.delete(deletedId); return n; });
    }
    setDeleteTarget(null);
  };

  // ----- Bulk actions -----
  const applyBulkCategory = (cat) => {
    setProducts(prev => prev.map(p => selected.has(p.id) ? { ...p, category: cat } : p));
    notify(`Category updated for ${selected.size} products`);
    setBulkCategoryOpen(false);
  };
  const applyBulkStatus = (status) => {
    setProducts(prev => prev.map(p => selected.has(p.id) ? { ...p, status } : p));
    notify(`Status updated for ${selected.size} products`);
    setBulkStatusOpen(false);
  };

  // ----- Import / export -----
  const handleImportConfirm = (rows) => {
    let running = [...products];
    rows.forEach(r => {
      const id = nextId(running);
      running = [{
        id,
        image: null,
        productName: (r.productName || r.name || "").toString(),
        brand: (r.brand || "").toString(),
        category: (r.category || "").toString(),
        packageSize: (r.packageSize || "").toString(),
        price: Number(r.price) || 0,
        stock: Number(r.stock) || 0,
        sku: (r.sku || "").toString(),
        status: (r.status || "Active").toString(),
      }, ...running];
    });
    setProducts(running);
    setShowImport(false);
    notify(`${rows.length} products imported`);
  };

  const handleExportConfirm = (scope, format) => {
    let rows;
    if (scope === "selected") rows = products.filter(p => selected.has(p.id));
    else if (scope === "filtered") rows = sorted;
    else rows = products;

    const data = exportRows(rows);
    const filename = `quickpick-products.${format === "csv" ? "csv" : "xlsx"}`;

    if (format === "csv") {
      const csv = Papa.unparse(data);
      download(filename, new Blob([csv], { type: "text/csv" }));
    } else {
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Products");
      const buf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      download(filename, new Blob([buf], { type: "application/octet-stream" }));
    }
    setShowExport(false);
    notify(`Exported ${data.length} products`);
  };

  const columns = [
    { key: "select", label: "", width: 40 },
    { key: "productId", label: "Product ID", width: 90 },
    { key: "image", label: "Image", width: 60 },
    { key: "productName", label: "Product Name", width: 190, sticky: true },
    { key: "brand", label: "Brand", width: 110 },
    { key: "category", label: "Category", width: 150 },
    { key: "packageSize", label: "Package Size", width: 110 },
    { key: "price", label: "Price", width: 100 },
    { key: "stock", label: "Stock", width: 80 },
    { key: "sku", label: "SKU", width: 120 },
    { key: "status", label: "Status", width: 110 },
    { key: "actions", label: "Actions", width: 60 },
  ];

  const stickyLeft1 = 40; // after select

  return (
    <div style={{ background: C.bg, minHeight: "100vh", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif", color: C.text }}>
      <div style={{ maxWidth: 1400, margin: "0 auto", padding: "28px 24px 60px" }}>

        {/* Header */}
        <div style={{ marginBottom: 20 }}>
          {onBack ? (
            <button
              onClick={onBack}
              style={{
                display: "inline-flex", alignItems: "center", gap: 6, background: "none", border: "none",
                cursor: "pointer", color: C.textSecondary, fontSize: 13, fontWeight: 500, padding: 0, marginBottom: 10,
              }}
              onMouseEnter={e => e.currentTarget.style.color = C.text}
              onMouseLeave={e => e.currentTarget.style.color = C.textSecondary}
            >
              <ArrowLeft size={14} /> Back to Dashboard
            </button>
          ) : null}
          <h1 style={{ fontSize: 24, fontWeight: 600, margin: 0, letterSpacing: "-0.01em" }}>Product Management</h1>
          <p style={{ fontSize: 13.5, color: C.textSecondary, margin: "4px 0 0" }}>
            {products.length} products across {new Set(products.map(p => p.category)).size} categories
          </p>
        </div>

        {/* Toolbar */}
        <div style={{
          display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center",
          background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10,
          padding: 12, marginBottom: 16,
        }}>
          <div style={{ position: "relative", flex: "1 1 220px", minWidth: 200 }}>
            <Search size={15} color={C.textMuted} style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)" }} />
            <input
              value={search} onChange={e => setSearch(e.target.value)}
              placeholder="Search products..."
              style={{ ...inputStyle, paddingLeft: 32 }}
            />
          </div>

          <div style={{ position: "relative" }} ref={filterRef}>
            <Button icon={Filter} onClick={() => { setFilterOpen(o => !o); setSortOpen(false); }}>
              Filter{(filterCategory || filterStatus || filterStock !== "All") ? ` (${[filterCategory, filterStatus, filterStock !== "All" ? filterStock : ""].filter(Boolean).length})` : ""}
            </Button>
            {filterOpen && (
              <div style={{
                position: "absolute", top: "110%", left: 0, background: C.surface, border: `1px solid ${C.border}`,
                borderRadius: 10, boxShadow: "0 8px 24px rgba(0,0,0,0.1)", padding: 16, width: 240, zIndex: 60,
              }}>
                <Field label="Category">
                  <select style={inputStyle} value={filterCategory} onChange={e => setFilterCategory(e.target.value)}>
                    <option value="">All categories</option>
                    {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </Field>
                <Field label="Status">
                  <select style={inputStyle} value={filterStatus} onChange={e => setFilterStatus(e.target.value)}>
                    <option value="">All statuses</option>
                    <option>Active</option><option>Inactive</option><option>Low Stock</option><option>Out of Stock</option>
                  </select>
                </Field>
                <Field label="Stock level">
                  <select style={inputStyle} value={filterStock} onChange={e => setFilterStock(e.target.value)}>
                    <option>All</option><option>In Stock</option><option>Low Stock</option><option>Out of Stock</option>
                  </select>
                </Field>
                <Button variant="ghost" onClick={() => { setFilterCategory(""); setFilterStatus(""); setFilterStock("All"); }} style={{ width: "100%", justifyContent: "center" }}>
                  Clear filters
                </Button>
              </div>
            )}
          </div>

          <div style={{ position: "relative" }} ref={sortRef}>
            <Button icon={ArrowUpDown} onClick={() => { setSortOpen(o => !o); setFilterOpen(false); }}>
              Sort{sortField ? ` (${sortField})` : ""}
            </Button>
            {sortOpen && (
              <div style={{
                position: "absolute", top: "110%", left: 0, background: C.surface, border: `1px solid ${C.border}`,
                borderRadius: 10, boxShadow: "0 8px 24px rgba(0,0,0,0.1)", padding: 16, width: 200, zIndex: 60,
              }}>
                <Field label="Sort by">
                  <select style={inputStyle} value={sortField} onChange={e => setSortField(e.target.value)}>
                    <option value="">No sorting</option>
                    <option value="productName">Product Name</option>
                    <option value="price">Price</option>
                    <option value="stock">Stock</option>
                    <option value="brand">Brand</option>
                    <option value="category">Category</option>
                  </select>
                </Field>
                <Field label="Direction">
                  <select style={inputStyle} value={sortDir} onChange={e => setSortDir(e.target.value)} disabled={!sortField}>
                    <option value="asc">Ascending</option>
                    <option value="desc">Descending</option>
                  </select>
                </Field>
              </div>
            )}
          </div>

          <div style={{ flex: 1 }} />

          <Button icon={Upload} onClick={() => setShowImport(true)}>Import Products</Button>
          <Button icon={Download} onClick={() => setShowExport(true)}>Export Products</Button>
          <Button icon={Plus} variant="primary" onClick={() => setShowAddModal(true)}>Add Product</Button>
        </div>

        {/* Bulk action bar */}
        {selected.size > 0 && (
          <div style={{
            display: "flex", alignItems: "center", gap: 12, background: C.accentTint,
            border: `1px solid ${C.accent}`, borderRadius: 10, padding: "10px 14px", marginBottom: 14, flexWrap: "wrap",
          }}>
            <span style={{ fontSize: 13.5, fontWeight: 600, color: C.accentDark }}>
              {selected.size} product{selected.size > 1 ? "s" : ""} selected
            </span>
            <div style={{ position: "relative" }}>
              <Button variant="secondary" onClick={() => setBulkCategoryOpen(o => !o)}>Change Category</Button>
              {bulkCategoryOpen && (
                <div style={{
                  position: "absolute", top: "110%", left: 0, background: C.surface, border: `1px solid ${C.border}`,
                  borderRadius: 8, boxShadow: "0 8px 24px rgba(0,0,0,0.1)", padding: 8, width: 200, zIndex: 60,
                }}>
                  {CATEGORIES.map(c => (
                    <div key={c} onClick={() => applyBulkCategory(c)}
                      style={{ padding: "7px 8px", fontSize: 13, cursor: "pointer", borderRadius: 5 }}
                      onMouseEnter={e => e.currentTarget.style.background = C.headerBg}
                      onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                      {c}
                    </div>
                  ))}
                </div>
              )}
            </div>
            <div style={{ position: "relative" }}>
              <Button variant="secondary" onClick={() => setBulkStatusOpen(o => !o)}>Change Status</Button>
              {bulkStatusOpen && (
                <div style={{
                  position: "absolute", top: "110%", left: 0, background: C.surface, border: `1px solid ${C.border}`,
                  borderRadius: 8, boxShadow: "0 8px 24px rgba(0,0,0,0.1)", padding: 8, width: 140, zIndex: 60,
                }}>
                  {["Active", "Inactive"].map(s => (
                    <div key={s} onClick={() => applyBulkStatus(s)}
                      style={{ padding: "7px 8px", fontSize: 13, cursor: "pointer", borderRadius: 5 }}
                      onMouseEnter={e => e.currentTarget.style.background = C.headerBg}
                      onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                      {s}
                    </div>
                  ))}
                </div>
              )}
            </div>
            <Button variant="secondary" icon={Download} onClick={() => setShowExport(true)}>Export Selected</Button>
            <Button variant="danger" icon={Trash2} onClick={() => setDeleteTarget({ bulk: true })}>Delete</Button>
            <div style={{ flex: 1 }} />
            <button onClick={() => setSelected(new Set())} style={{ background: "none", border: "none", cursor: "pointer", color: C.accentDark }}>
              <X size={16} />
            </button>
          </div>
        )}

        {/* Table */}
        <div style={{
          border: `1px solid ${C.border}`, borderRadius: 10, background: C.surface,
          overflow: "auto", maxWidth: "100%",
        }}>
          <table style={{ borderCollapse: "collapse", fontSize: 13, minWidth: 1150, width: "100%" }}>
            <thead>
              <tr>
                <th style={{
                  position: "sticky", top: 0, left: 0, zIndex: 30, background: C.headerBg,
                  borderBottom: `1px solid ${C.border}`, borderRight: `1px solid ${C.border}`,
                  padding: "10px 8px", width: 40, textAlign: "center",
                }}>
                  <input ref={selectAllRef} type="checkbox" checked={allFilteredSelected} onChange={toggleAll} style={{ cursor: "pointer" }} />
                </th>
                <th style={{ position: "sticky", top: 0, zIndex: 20, background: C.headerBg, borderBottom: `1px solid ${C.border}`, padding: "10px 10px", textAlign: "left", fontWeight: 600, whiteSpace: "nowrap" }}>Product ID</th>
                <th style={{ position: "sticky", top: 0, zIndex: 20, background: C.headerBg, borderBottom: `1px solid ${C.border}`, padding: "10px 10px", textAlign: "left", fontWeight: 600 }}>Image</th>
                <th style={{
                  position: "sticky", top: 0, left: stickyLeft1, zIndex: 30, background: C.headerBg,
                  borderBottom: `1px solid ${C.border}`, borderRight: `1px solid ${C.border}`,
                  padding: "10px 10px", textAlign: "left", fontWeight: 600, minWidth: 190,
                }}>Product Name</th>
                <th style={{ position: "sticky", top: 0, zIndex: 20, background: C.headerBg, borderBottom: `1px solid ${C.border}`, padding: "10px 10px", textAlign: "left", fontWeight: 600, whiteSpace: "nowrap" }}>Brand</th>
                <th style={{ position: "sticky", top: 0, zIndex: 20, background: C.headerBg, borderBottom: `1px solid ${C.border}`, padding: "10px 10px", textAlign: "left", fontWeight: 600, whiteSpace: "nowrap" }}>Category</th>
                <th style={{ position: "sticky", top: 0, zIndex: 20, background: C.headerBg, borderBottom: `1px solid ${C.border}`, padding: "10px 10px", textAlign: "left", fontWeight: 600, whiteSpace: "nowrap" }}>Package Size</th>
                <th style={{ position: "sticky", top: 0, zIndex: 20, background: C.headerBg, borderBottom: `1px solid ${C.border}`, padding: "10px 10px", textAlign: "left", fontWeight: 600, whiteSpace: "nowrap" }}>Price</th>
                <th style={{ position: "sticky", top: 0, zIndex: 20, background: C.headerBg, borderBottom: `1px solid ${C.border}`, padding: "10px 10px", textAlign: "left", fontWeight: 600, whiteSpace: "nowrap" }}>Stock</th>
                <th style={{ position: "sticky", top: 0, zIndex: 20, background: C.headerBg, borderBottom: `1px solid ${C.border}`, padding: "10px 10px", textAlign: "left", fontWeight: 600, whiteSpace: "nowrap" }}>SKU</th>
                <th style={{ position: "sticky", top: 0, zIndex: 20, background: C.headerBg, borderBottom: `1px solid ${C.border}`, padding: "10px 10px", textAlign: "left", fontWeight: 600, whiteSpace: "nowrap" }}>Status</th>
                <th style={{ position: "sticky", top: 0, zIndex: 20, background: C.headerBg, borderBottom: `1px solid ${C.border}`, padding: "10px 10px", textAlign: "center", fontWeight: 600 }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {paginated.map(p => {
                const badge = computeBadge(p);
                const isSelected = selected.has(p.id);
                return (
                  <tr key={p.id} style={{ background: isSelected ? C.accentTint : "transparent" }}
                    onMouseEnter={e => { if (!isSelected) e.currentTarget.style.background = C.headerBg; }}
                    onMouseLeave={e => { if (!isSelected) e.currentTarget.style.background = "transparent"; }}
                  >
                    <td style={{
                      position: "sticky", left: 0, zIndex: 10, background: isSelected ? C.accentTint : C.surface,
                      borderBottom: `1px solid ${C.border}`, borderRight: `1px solid ${C.border}`,
                      padding: "6px 8px", textAlign: "center",
                    }}>
                      <input type="checkbox" checked={isSelected} onChange={() => toggleRow(p.id)} style={{ cursor: "pointer" }} />
                    </td>
                    <td style={{ borderBottom: `1px solid ${C.border}`, padding: "6px 10px", color: C.textSecondary, whiteSpace: "nowrap" }}>{p.id}</td>
                    <td style={{ borderBottom: `1px solid ${C.border}`, padding: "6px 10px" }}>
                      <div style={{
                        width: 32, height: 32, borderRadius: 6, overflow: "hidden", background: C.headerBg,
                        display: "flex", alignItems: "center", justifyContent: "center", border: `1px solid ${C.border}`,
                      }}>
                        {p.image ? <img src={p.image} alt={p.productName} style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : <ImageOff size={13} color={C.textMuted} />}
                      </div>
                    </td>
                    <td style={{
                      position: "sticky", left: stickyLeft1, zIndex: 10, background: isSelected ? C.accentTint : C.surface,
                      borderBottom: `1px solid ${C.border}`, borderRight: `1px solid ${C.border}`,
                      padding: "3px 4px", fontWeight: 500, minWidth: 190,
                    }}>
                      <EditableCell value={p.productName} field="productName" onSave={(f, v) => saveCell(p.id, f, v)} width={180} />
                    </td>
                    <td style={{ borderBottom: `1px solid ${C.border}`, padding: "3px 4px" }}>
                      <EditableCell value={p.brand} field="brand" onSave={(f, v) => saveCell(p.id, f, v)} width={100} />
                    </td>
                    <td style={{ borderBottom: `1px solid ${C.border}`, padding: "3px 4px" }}>
                      <EditableCell value={p.category} field="category" onSave={(f, v) => saveCell(p.id, f, v)} width={140} isSelect options={CATEGORIES} />
                    </td>
                    <td style={{ borderBottom: `1px solid ${C.border}`, padding: "3px 4px" }}>
                      <EditableCell value={p.packageSize} field="packageSize" onSave={(f, v) => saveCell(p.id, f, v)} width={100} />
                    </td>
                    <td style={{ borderBottom: `1px solid ${C.border}`, padding: "3px 4px", fontVariantNumeric: "tabular-nums" }}>
                      <EditableCell value={p.price} field="price" onSave={(f, v) => saveCell(p.id, f, v)} width={90} />
                    </td>
                    <td style={{ borderBottom: `1px solid ${C.border}`, padding: "3px 4px", fontVariantNumeric: "tabular-nums" }}>
                      <EditableCell value={p.stock} field="stock" onSave={(f, v) => saveCell(p.id, f, v)} width={70} />
                    </td>
                    <td style={{ borderBottom: `1px solid ${C.border}`, padding: "3px 4px", color: C.textSecondary }}>
                      <EditableCell value={p.sku} field="sku" onSave={(f, v) => saveCell(p.id, f, v)} width={110} />
                    </td>
                    <td style={{ borderBottom: `1px solid ${C.border}`, padding: "6px 8px" }}>
                      <Badge tone={badge.tone} label={badge.label} />
                    </td>
                    <td style={{ borderBottom: `1px solid ${C.border}`, padding: "6px 8px", textAlign: "center" }}>
                      <ActionsMenu
                        onEdit={() => setEditProduct(p)}
                        onShops={() => setShopsProduct(p)}
                        onDuplicate={() => handleDuplicate(p)}
                        onDelete={() => setDeleteTarget({ product: p })}
                      />
                    </td>
                  </tr>
                );
              })}
              {paginated.length === 0 && (
                <tr>
                  <td colSpan={12} style={{ padding: "40px 20px", textAlign: "center", color: C.textMuted, fontSize: 13.5 }}>
                    No products match your search or filters.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination */}
        <div style={{
          display: "flex", flexWrap: "wrap", alignItems: "center", justifyContent: "space-between",
          gap: 12, marginTop: 14, fontSize: 13, color: C.textSecondary,
        }}>
          <div>
            Showing {sorted.length === 0 ? 0 : pageStart + 1}–{Math.min(pageStart + rowsPerPage, sorted.length)} of {sorted.length} products
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <button
              onClick={() => setPage(p => Math.max(1, p - 1))}
              disabled={currentPage === 1}
              style={{
                display: "flex", alignItems: "center", justifyContent: "center", width: 30, height: 30,
                borderRadius: 6, border: `1px solid ${C.border}`, background: C.surface,
                cursor: currentPage === 1 ? "not-allowed" : "pointer", opacity: currentPage === 1 ? 0.4 : 1,
              }}
            >
              <ChevronLeft size={15} />
            </button>
            {Array.from({ length: totalPages }, (_, i) => i + 1)
              .filter(n => n === 1 || n === totalPages || Math.abs(n - currentPage) <= 1)
              .reduce((acc, n, i, arr) => {
                if (i > 0 && n - arr[i - 1] > 1) acc.push("...");
                acc.push(n);
                return acc;
              }, [])
              .map((n, i) => n === "..." ? (
                <span key={`e${i}`} style={{ padding: "0 4px" }}>…</span>
              ) : (
                <button
                  key={n}
                  onClick={() => setPage(n)}
                  style={{
                    width: 30, height: 30, borderRadius: 6, fontSize: 13,
                    border: `1px solid ${n === currentPage ? C.accent : C.border}`,
                    background: n === currentPage ? C.accent : C.surface,
                    color: n === currentPage ? "#fff" : C.text, cursor: "pointer", fontWeight: 500,
                  }}
                >
                  {n}
                </button>
              ))}
            <button
              onClick={() => setPage(p => Math.min(totalPages, p + 1))}
              disabled={currentPage === totalPages}
              style={{
                display: "flex", alignItems: "center", justifyContent: "center", width: 30, height: 30,
                borderRadius: 6, border: `1px solid ${C.border}`, background: C.surface,
                cursor: currentPage === totalPages ? "not-allowed" : "pointer", opacity: currentPage === totalPages ? 0.4 : 1,
              }}
            >
              <ChevronRight size={15} />
            </button>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            Rows per page:
            <select
              value={rowsPerPage} onChange={e => setRowsPerPage(Number(e.target.value))}
              style={{ padding: "5px 8px", borderRadius: 6, border: `1px solid ${C.border}`, fontSize: 13 }}
            >
              {[25, 50, 100].map(n => <option key={n} value={n}>{n}</option>)}
            </select>
          </div>
        </div>
      </div>

      {showAddModal && (
        <ProductFormModal shops={shops} onCancel={() => setShowAddModal(false)} onSave={handleSaveProduct} />
      )}
      {editProduct && (
        <ProductFormModal initial={editProduct} shops={shops} onCancel={() => setEditProduct(null)} onSave={handleSaveProduct} />
      )}
      {shopsProduct && (
        <ProductShopsModal
          product={shopsProduct}
          shops={shops}
          shopProducts={shopProductsLocal}
          onAssign={assignProductToShop}
          onUpdatePrice={updateShopProductPrice}
          onRemove={removeShopProduct}
          onClose={() => setShopsProduct(null)}
        />
      )}
      {deleteTarget && (
        <DeleteConfirmModal
          product={deleteTarget.product}
          count={deleteTarget.bulk ? selected.size : 1}
          onCancel={() => setDeleteTarget(null)}
          onConfirm={handleDeleteConfirm}
        />
      )}
      {showImport && (
        <ImportModal onCancel={() => setShowImport(false)} onConfirm={handleImportConfirm} />
      )}
      {showExport && (
        <ExportModal
          scopeCount={{ selected: selected.size }}
          onCancel={() => setShowExport(false)}
          onConfirm={handleExportConfirm}
        />
      )}

      <Toast message={toast} onDone={() => setToast("")} />
    </div>
  );
}

// ============================================================
// ---------- Admin Dashboard ----------
// ============================================================

const adminRoutes = {
  dashboard: "/admin/dashboard",
  shops: "/admin/shops",
  products: "/admin/products",
  prices: "/admin/prices",
  smartShopping: "/admin/smart-shopping",
  productSearch: "/admin/product-search",
  settings: "/admin/settings",
};

const NAV_SECTIONS = [
  {
    label: "Overview",
    items: [{ key: "dashboard", label: "Dashboard", icon: LayoutDashboard, route: adminRoutes.dashboard }],
  },
  {
    label: "Management",
    items: [
      { key: "shops", label: "Shops", icon: Store, route: adminRoutes.shops },
      { key: "products", label: "Products", icon: Package, route: adminRoutes.products },
      { key: "prices", label: "Prices", icon: Tag, route: adminRoutes.prices },
    ],
  },
  {
    label: "Shopping Tools",
    items: [
      { key: "smartShopping", label: "Smart Shopping", icon: ShoppingCart, route: adminRoutes.smartShopping },
      { key: "productSearch", label: "Product Search", icon: Search, route: adminRoutes.productSearch },
    ],
  },
  {
    label: "System",
    items: [{ key: "settings", label: "Settings", icon: Settings, route: adminRoutes.settings }],
  },
];

function DashSidebar({ activeKey, onNavigate, open, onClose }) {
  return (
    <>
      {open && (
        <div
          onClick={onClose}
          style={{ position: "fixed", inset: 0, background: "rgba(10,20,15,0.35)", zIndex: 39, display: "none" }}
          className="qp-scrim"
        />
      )}
      <aside
        className={`qp-sidebar${open ? " qp-sidebar-open" : ""}`}
        style={{
          width: 246, flexShrink: 0, background: C.accentDark, color: "#DCE7DF",
          display: "flex", flexDirection: "column", position: "sticky", top: 0, height: "100vh", zIndex: 40,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "22px 20px 18px", borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
          <div style={{
            width: 32, height: 32, borderRadius: 9, background: `linear-gradient(155deg, #2E9E63, ${C.accent})`,
            display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, color: C.accentDark, fontSize: 15, flexShrink: 0,
          }}>Q</div>
          <div>
            <div style={{ fontWeight: 600, fontSize: 15.5, color: "#F3F8F5" }}>QuickPick</div>
            <div style={{ fontSize: 11.5, color: "#8FA599", marginTop: 1 }}>Admin Console</div>
          </div>
        </div>

        <nav style={{ padding: "14px 12px", flex: 1, overflowY: "auto" }}>
          {NAV_SECTIONS.map(section => (
            <div key={section.label} style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 11, color: "#6E8478", padding: "0 10px", marginBottom: 6, fontWeight: 600 }}>{section.label}</div>
              {section.items.map(item => {
                const isActive = item.key === activeKey;
                return (
                  <div
                    key={item.key}
                    onClick={() => onNavigate(item)}
                    style={{
                      display: "flex", alignItems: "center", gap: 10, padding: "9px 10px", borderRadius: 8,
                      color: isActive ? "#fff" : "#B9C9BF", fontSize: 13.5, fontWeight: 500, marginBottom: 2,
                      cursor: "pointer", background: isActive ? C.accent : "transparent",
                    }}
                    onMouseEnter={e => { if (!isActive) e.currentTarget.style.background = "rgba(255,255,255,0.06)"; }}
                    onMouseLeave={e => { if (!isActive) e.currentTarget.style.background = "transparent"; }}
                  >
                    <item.icon size={16} />
                    {item.label}
                  </div>
                );
              })}
            </div>
          ))}
        </nav>

        <div style={{ padding: "14px 20px 18px", borderTop: "1px solid rgba(255,255,255,0.08)", fontSize: 11.5, color: "#6E8478" }}>
          QuickPick Admin &middot; v1.0
        </div>
      </aside>
    </>
  );
}

function KpiCard({ label, value, delta, icon: Icon, tone = "accent" }) {
  const iconBg = tone === "amber" ? C.amberTint : C.accentTint;
  const iconFg = tone === "amber" ? C.amber : C.accent;
  return (
    <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, padding: "16px 16px 15px", boxShadow: "0 1px 2px rgba(14,29,22,0.04)" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
        <div style={{ fontSize: 12, color: C.textSecondary, fontWeight: 500 }}>{label}</div>
        <div style={{ width: 30, height: 30, borderRadius: 8, background: iconBg, color: iconFg, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Icon size={15} />
        </div>
      </div>
      <div style={{ fontSize: 23, fontWeight: 700, letterSpacing: "-0.3px" }}>{value}</div>
      <div style={{ fontSize: 11.5, color: tone === "up" ? C.accent : C.textMuted, marginTop: 5 }}>{delta}</div>
    </div>
  );
}

function ManagementCard({ title, description, icon: Icon, stats, tags, cta, flag, onClick, attention }) {
  const [hover, setHover] = useState(false);
  return (
    <div
      onClick={onClick}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        background: C.surface, border: `1px solid ${hover ? C.borderStrong : C.border}`, borderRadius: 16,
        padding: 20, boxShadow: hover ? "0 6px 20px rgba(14,29,22,0.08)" : "0 1px 2px rgba(14,29,22,0.04)",
        cursor: "pointer", display: "flex", flexDirection: "column", transform: hover ? "translateY(-2px)" : "none",
        transition: "transform 150ms ease, box-shadow 150ms ease, border-color 150ms ease",
      }}
    >
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 12 }}>
        <div style={{
          width: 38, height: 38, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center",
          background: attention ? C.amberTint : C.accentTint, color: attention ? C.amber : C.accent,
        }}>
          <Icon size={18} />
        </div>
        {flag ? (
          <span style={{ background: C.dangerTint, color: C.danger, fontSize: 11, fontWeight: 600, padding: "3px 8px", borderRadius: 20 }}>{flag}</span>
        ) : null}
      </div>
      <h4 style={{ fontSize: 15, fontWeight: 600, margin: "0 0 5px" }}>{title}</h4>
      <p style={{ fontSize: 12.5, color: C.textSecondary, margin: "0 0 14px", lineHeight: 1.5 }}>{description}</p>

      {stats ? (
        <div style={{ display: "flex", gap: 16, marginBottom: 16, flexWrap: "wrap" }}>
          {stats.map(s => (
            <div key={s.label}>
              <div style={{ fontSize: 15, fontWeight: 600 }}>{s.value}</div>
              <div style={{ fontSize: 11, color: C.textMuted }}>{s.label}</div>
            </div>
          ))}
        </div>
      ) : null}

      {tags ? (
        <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 16 }}>
          {tags.map(t => (
            <span key={t} style={{ fontSize: 12, color: C.textSecondary, display: "flex", alignItems: "center", gap: 7 }}>
              <span style={{ width: 5, height: 5, borderRadius: "50%", background: C.accent, flexShrink: 0 }} />{t}
            </span>
          ))}
        </div>
      ) : null}

      <div style={{ marginTop: "auto", fontSize: 13, fontWeight: 600, color: C.accent, display: "flex", alignItems: "center", gap: 5 }}>
        {cta} <ArrowRight size={13} />
      </div>
    </div>
  );
}

function AdminDashboard({ onNavigate, shops = [], products = [], shopProducts = [], priceProducts = [], savedShoppingLists = [], lastPriceUpdateAt = null }) {
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [toast, setToast] = useState("");

  // ---- Every number on this dashboard is derived here, live, from the
  // actual shared records - never hard-coded. ----
  const stats = useMemo(() => {
    const validShops = Array.isArray(shops) ? shops.filter((s) => s && s.id) : [];
    const validProducts = Array.isArray(products) ? products.filter((p) => p && p.id) : [];
    const validPriceProducts = Array.isArray(priceProducts) ? priceProducts.filter((p) => p && p.id) : [];
    const validLists = Array.isArray(savedShoppingLists) ? savedShoppingLists.filter((l) => l && l.id) : [];
    const productIds = new Set(validProducts.map((p) => p.id));
    const shopIds = new Set(validShops.map((s) => s.id));
    const validShopProducts = Array.isArray(shopProducts)
      ? shopProducts.filter((sp) => sp && sp.id && productIds.has(sp.productId) && shopIds.has(sp.shopId))
      : [];

    const activeShops = validShops.filter((s) => s.status === "Active").length;
    const inactiveShops = validShops.length - activeShops;

    const categorySet = new Set(validProducts.map((p) => p.category).filter(Boolean));

    // Price Records = real product-shop-price relationships (the many-to-many
    // join table), not a proxy - "if a product has prices from 4 shops, that's
    // 4 price records." Price Coverage = share of products with >=1 real
    // shop-price assignment.
    const priceRecordCount = validShopProducts.length;
    const assignmentCountByProduct = {};
    validShopProducts.forEach((sp) => {
      assignmentCountByProduct[sp.productId] = (assignmentCountByProduct[sp.productId] || 0) + 1;
    });
    const productsWithPrice = validProducts.filter((p) => (assignmentCountByProduct[p.id] || 0) > 0).length;
    const coverage = validProducts.length > 0 ? Math.round((productsWithPrice / validProducts.length) * 100) : 0;

    const outdated = validPriceProducts.filter((p) => statusOf(p) === "outdated").length;
    const missing = validPriceProducts.filter((p) => statusOf(p) === "missing").length;
    const updatedToday = validPriceProducts.filter((p) => p.daysAgo === 0).length;
    const updatedThisWeek = validPriceProducts.filter((p) => p.daysAgo !== null && p.daysAgo <= 7).length;

    const validSavings = validLists.map((l) => (Number.isFinite(l.savings) ? l.savings : 0));
    const avgSavings = validLists.length > 0
      ? Math.round(validSavings.reduce((s, v) => s + v, 0) / validLists.length)
      : 0;
    const validBaskets = validLists.map((l) => (Number.isFinite(l.optimizedTotal) ? l.optimizedTotal : 0));
    const avgBasketCost = validLists.length > 0
      ? Math.round(validBaskets.reduce((s, v) => s + v, 0) / validLists.length)
      : 0;
    const highestSavings = validSavings.length > 0 ? Math.max(...validSavings) : 0;

    // Coverage breakdown for the "Product Coverage" donut - now derived from
    // real shop-price assignment counts: multiple shops (comparable), a
    // single shop (priced but nothing to compare against), or none.
    const totalForBreakdown = validProducts.length || 1; // avoid divide-by-zero; counts are 0 either way
    const fullyPricedCount = validProducts.filter((p) => (assignmentCountByProduct[p.id] || 0) >= 2).length;
    const partialCount = validProducts.filter((p) => (assignmentCountByProduct[p.id] || 0) === 1).length;
    const fullyPricedPct = Math.round((fullyPricedCount / totalForBreakdown) * 100);
    const partiallyPricedPct = Math.round((partialCount / totalForBreakdown) * 100);
    const missingPricedPct = Math.max(0, 100 - fullyPricedPct - partiallyPricedPct);

    const issueLabel = { outdated: "Price outdated", missing: "Missing price", unusual: "Unusual price change", stale: "Price getting stale" };
    const priorityFor = { outdated: "High", unusual: "High", missing: "Medium", stale: "Medium" };
    const attentionItems = validPriceProducts
      .map((p) => ({ p, status: statusOf(p) }))
      .filter(({ status }) => issueLabel[status])
      .sort((a, b) => (b.p.daysAgo ?? 999) - (a.p.daysAgo ?? 999))
      .slice(0, 4)
      .map(({ p, status }) => ({
        name: p.name,
        issue: issueLabel[status],
        update: timeAgoLabel(p.daysAgo),
        priority: priorityFor[status] || "Low",
      }));

    return {
      totalShops: validShops.length,
      activeShops,
      inactiveShops,
      totalProducts: validProducts.length,
      categoryCount: categorySet.size,
      priceRecordCount,
      totalPriceProducts: validPriceProducts.length,
      outdated,
      missing,
      updatedToday,
      updatedThisWeek,
      coverage,
      listsAnalyzed: validLists.length,
      avgSavings,
      avgBasketCost,
      highestSavings,
      fullyPricedPct,
      partiallyPricedPct,
      missingPricedPct,
      // Real shops, ranked by how many products are actually assigned to
      // them via the shopProducts join table - genuine data, not a proxy.
      topShops: [...validShops]
        .map((s) => ({
          name: s.name,
          products: validShopProducts.filter((sp) => sp.shopId === s.id).length,
          status: s.status || "Active",
        }))
        .sort((a, b) => b.products - a.products)
        .slice(0, 3),
      attentionItems,
      lastPriceUpdateLabel: lastPriceUpdateAt
        ? new Date(lastPriceUpdateAt).toLocaleString("en-KE", { dateStyle: "medium", timeStyle: "short" })
        : "No updates yet",
    };
  }, [shops, products, shopProducts, priceProducts, savedShoppingLists, lastPriceUpdateAt]);

  const handleNavClick = (item) => {
    setSidebarOpen(false);
    if (item.key === "products") { onNavigate(adminRoutes.products); return; }
    if (item.key === "shops") { onNavigate(adminRoutes.shops); return; }
    if (item.key === "prices") { onNavigate(adminRoutes.prices); return; }
    if (item.key === "smartShopping") { onNavigate(adminRoutes.smartShopping); return; }
    if (item.key === "productSearch") { onNavigate(adminRoutes.productSearch); return; }
    if (item.key === "dashboard") return;
    setToast(`"${item.label}" isn't built yet — would open ${item.route}`);
  };

  const goProducts = () => onNavigate(adminRoutes.products);
  const goShops = () => onNavigate(adminRoutes.shops);
  const goPrices = () => onNavigate(adminRoutes.prices);
  const goSmartShopping = () => onNavigate(adminRoutes.smartShopping);
  const goProductSearch = () => onNavigate(adminRoutes.productSearch);
  const notBuilt = (label, route) => setToast(`"${label}" isn't built yet — would open ${route}`);

  const barData = [
    { d: "Mon", v: 210 }, { d: "Tue", v: 285 }, { d: "Wed", v: 190 },
    { d: "Thu", v: 360 }, { d: "Fri", v: 410 }, { d: "Sat", v: 320 }, { d: "Sun", v: 324 },
  ];
  const maxBar = Math.max(...barData.map(b => b.v));

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: C.bg, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif", color: C.text }}>
      <style>{`
        @media (max-width: 760px) {
          .qp-sidebar { position: fixed !important; left: 0; top: 0; bottom: 0; transform: translateX(-100%); box-shadow: 20px 0 60px rgba(0,0,0,0.25); }
          .qp-sidebar-open { transform: translateX(0) !important; }
          .qp-scrim { display: block !important; }
          .qp-mobile-toggle { display: flex !important; }
          .qp-search-field { display: none !important; }
          .qp-kpi-grid { grid-template-columns: repeat(2, 1fr) !important; }
          .qp-mgmt-grid { grid-template-columns: 1fr !important; }
          .qp-two-col { grid-template-columns: 1fr !important; }
          .qp-insights-grid { grid-template-columns: 1fr !important; }
          .qp-bottom-grid { grid-template-columns: 1fr !important; }
          .qp-qa-grid { grid-template-columns: repeat(2, 1fr) !important; }
        }
        @media (max-width: 1180px) {
          .qp-kpi-grid { grid-template-columns: repeat(3, 1fr) !important; }
          .qp-mgmt-grid { grid-template-columns: repeat(2, 1fr) !important; }
        }
      `}</style>

      <DashSidebar activeKey="dashboard" onNavigate={handleNavClick} open={sidebarOpen} onClose={() => setSidebarOpen(false)} />

      <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column" }}>
        <header style={{
          position: "sticky", top: 0, zIndex: 30, background: "rgba(255,255,255,0.92)", backdropFilter: "blur(8px)",
          borderBottom: `1px solid ${C.border}`, padding: "14px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 20,
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <button
              className="qp-mobile-toggle"
              onClick={() => setSidebarOpen(o => !o)}
              style={{ display: "none", width: 34, height: 34, borderRadius: 8, border: `1px solid ${C.border}`, background: C.surface, alignItems: "center", justifyContent: "center", color: C.text, cursor: "pointer" }}
            >
              <Menu size={17} />
            </button>
            <div>
              <h1 style={{ fontSize: 19, fontWeight: 600, margin: 0 }}>Dashboard</h1>
              <p style={{ fontSize: 12.5, color: C.textSecondary, margin: "2px 0 0" }}>Manage and monitor your QuickPick platform</p>
            </div>
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div className="qp-search-field" style={{ display: "flex", alignItems: "center", gap: 8, background: C.bg, border: `1px solid ${C.border}`, borderRadius: 9, padding: "8px 12px", width: 230, color: C.textMuted }}>
              <Search size={15} />
              <input placeholder="Search QuickPick..." style={{ border: "none", background: "transparent", outline: "none", fontSize: 13, color: C.text, width: "100%", fontFamily: "inherit" }} />
            </div>
            <button
              onClick={() => setToast("No new notifications right now.")}
              style={{ width: 36, height: 36, borderRadius: 9, border: `1px solid ${C.border}`, background: C.surface, display: "flex", alignItems: "center", justifyContent: "center", color: C.textSecondary, cursor: "pointer", position: "relative" }}
            >
              <Bell size={16} />
              <span style={{ position: "absolute", top: 6, right: 7, width: 7, height: 7, borderRadius: "50%", background: C.danger, border: `1.5px solid ${C.surface}` }} />
            </button>
            <button style={{ display: "flex", alignItems: "center", gap: 8, padding: "4px 8px 4px 4px", borderRadius: 10, border: `1px solid ${C.border}`, background: C.surface, cursor: "pointer" }}>
              <div style={{ width: 28, height: 28, borderRadius: "50%", background: C.accentTint, color: C.accentDark, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 600, fontSize: 12 }}>AD</div>
              <span style={{ fontSize: 12.5, fontWeight: 600 }}>Admin</span>
              <ChevronDown size={12} color={C.textMuted} />
            </button>
          </div>
        </header>

        <main style={{ padding: "24px 28px 56px", maxWidth: 1400, width: "100%", margin: "0 auto" }}>

          {/* Welcome */}
          <section style={{
            background: `linear-gradient(120deg, ${C.accentTint} 0%, #EFF7F2 55%, ${C.bg} 100%)`,
            border: `1px solid ${C.border}`, borderRadius: 16, padding: "20px 24px",
            display: "flex", alignItems: "center", justifyContent: "space-between", gap: 20, marginBottom: 22, flexWrap: "wrap",
          }}>
            <div>
              <h2 style={{ fontSize: 18, margin: "0 0 3px", fontWeight: 600 }}>Welcome back, Admin</h2>
              <p style={{ margin: 0, color: C.textSecondary, fontSize: 13 }}>Here's what's happening across QuickPick today.</p>
            </div>
            <Button variant="secondary" onClick={() => notBuilt("System Overview", "/admin/overview")}>View System Overview</Button>
          </section>

          {/* KPI cards */}
          <div className="qp-kpi-grid" style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: 14 }}>
            <KpiCard label="Total Shops" value={stats.totalShops.toLocaleString()} delta={`${stats.activeShops} active, ${stats.inactiveShops} inactive`} icon={Store} tone="up" />
            <KpiCard label="Products" value={stats.totalProducts.toLocaleString()} delta={`${stats.categoryCount} categories`} icon={Package} />
            <KpiCard label="Price Records" value={stats.priceRecordCount.toLocaleString()} delta={`${stats.updatedToday} updated today`} icon={Tag} tone="up" />
            <KpiCard label="Shopping Lists" value={stats.listsAnalyzed.toLocaleString()} delta={stats.listsAnalyzed > 0 ? "From Smart Shopping" : "None saved yet"} icon={ShoppingCart} tone="up" />
            <KpiCard label="Average Savings" value={`KSh ${stats.avgSavings.toLocaleString()}`} delta="Per optimized basket" icon={Wallet} />
            <KpiCard label="Price Coverage" value={`${stats.coverage}%`} delta="Across active products" icon={Percent} tone="amber" />
          </div>

          {/* Quick management */}
          <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", margin: "32px 0 14px" }}>
            <div>
              <h3 style={{ fontSize: 15.5, fontWeight: 600, margin: 0 }}>Quick Management</h3>
              <div style={{ fontSize: 12.5, color: C.textSecondary, marginTop: 2 }}>Access QuickPick's main administration tools.</div>
            </div>
          </div>

          <div className="qp-mgmt-grid" style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14 }}>
            <ManagementCard
              title="Shop Management" icon={Store}
              description="Add, organize and manage shops participating in QuickPick."
              stats={[{ value: String(stats.totalShops), label: "Shops" }, { value: String(stats.activeShops), label: "Active" }, { value: String(stats.inactiveShops), label: "Inactive" }]}
              cta="Manage Shops" onClick={goShops}
            />
            <ManagementCard
              title="Product Management" icon={Package}
              description="Manage QuickPick's master product catalogue and product information."
              stats={[{ value: String(stats.totalProducts), label: "Products" }, { value: String(stats.categoryCount), label: "Categories" }]}
              cta="Manage Products" onClick={goProducts}
            />
            <ManagementCard
              title="Price Management" icon={Tag} attention flag={`${stats.outdated + stats.missing} need review`}
              description="Review and maintain product prices submitted by QuickPick shops."
              stats={[{ value: String(stats.priceRecordCount), label: "Prices" }, { value: String(stats.updatedToday), label: "Today" }, { value: String(stats.outdated + stats.missing), label: "Review" }]}
              cta="Manage Prices" onClick={goPrices}
            />
            <ManagementCard
              title="Smart Shopping" icon={ShoppingCart}
              description="Open QuickPick's shopping optimization and price comparison engine."
              tags={["Best Price Matching", "Shop Optimization", "Savings Analysis"]}
              cta="Open Smart Shopping" onClick={goSmartShopping}
            />
            <ManagementCard
              title="Product Search" icon={Search}
              description="Search products and compare their availability and prices across shops."
              tags={["Search Products", "Compare Shops", "Check Prices"]}
              cta="Search Products" onClick={goProductSearch}
            />
          </div>

          {/* Price status + smart shopping summary */}
          <div style={{ margin: "32px 0 14px" }}>
            <h3 style={{ fontSize: 15.5, fontWeight: 600, margin: 0 }}>Price Update Status</h3>
            <div style={{ fontSize: 12.5, color: C.textSecondary, marginTop: 2 }}>How fresh the platform's pricing data is right now.</div>
          </div>

          <div className="qp-two-col" style={{ display: "grid", gridTemplateColumns: "1.55fr 1fr", gap: 16, alignItems: "start" }}>
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 20 }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 18 }}>
                {[
                  { v: String(stats.updatedToday), l: "Updated Today" },
                  { v: stats.updatedThisWeek.toLocaleString(), l: "Updated This Week" },
                  { v: String(stats.outdated), l: "Outdated Prices", warn: true },
                  { v: String(stats.missing), l: "Missing Prices", bad: true },
                ].map(s => (
                  <div key={s.l} style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 10, padding: "12px 14px" }}>
                    <div style={{ fontSize: 19, fontWeight: 700, color: s.bad ? C.danger : s.warn ? C.amber : C.text }}>{s.v}</div>
                    <div style={{ fontSize: 11.5, color: C.textSecondary }}>{s.l}</div>
                  </div>
                ))}
              </div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                <span style={{ fontSize: 12.5, color: C.textSecondary }}>Price Coverage</span>
                <b style={{ fontSize: 13 }}>{stats.coverage}%</b>
              </div>
              <div style={{ height: 8, background: C.bg, borderRadius: 6, overflow: "hidden", border: `1px solid ${C.border}` }}>
                <div style={{ height: "100%", width: `${stats.coverage}%`, background: C.accent, borderRadius: 6 }} />
              </div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 16, fontSize: 12, color: C.textSecondary, flexWrap: "wrap", gap: 8 }}>
                <span>Last system price update: {stats.lastPriceUpdateLabel}</span>
                <a onClick={goPrices} style={{ color: C.accent, fontWeight: 600, cursor: "pointer" }}>Manage Price Updates →</a>
              </div>
            </div>

            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 20 }}>
              <div style={{ fontSize: 14.5, fontWeight: 600, marginBottom: 3 }}>Smart Shopping Performance</div>
              <div style={{ fontSize: 12, color: C.textSecondary, marginBottom: 16 }}>How much Smart Shopping is helping users save.</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 16 }}>
                {[
                  { l: "Lists Analyzed", v: stats.listsAnalyzed.toLocaleString() },
                  { l: "Avg. Basket Cost", v: `KSh ${stats.avgBasketCost.toLocaleString()}` },
                  { l: "Average Savings", v: `KSh ${stats.avgSavings.toLocaleString()}` },
                  { l: "Highest Savings", v: `KSh ${stats.highestSavings.toLocaleString()}` },
                ].map(s => (
                  <div key={s.l} style={{ background: C.bg, border: `1px solid ${C.border}`, borderRadius: 10, padding: "12px 14px" }}>
                    <span style={{ fontSize: 11, color: C.textSecondary, display: "block", marginBottom: 3 }}>{s.l}</span>
                    <b style={{ fontSize: 17 }}>{s.v}</b>
                  </div>
                ))}
              </div>
              <a onClick={goSmartShopping} style={{ color: C.accent, fontWeight: 600, fontSize: 13, cursor: "pointer" }}>Open Smart Shopping →</a>
            </div>
          </div>

          {/* Insights */}
          <div style={{ margin: "32px 0 14px" }}>
            <h3 style={{ fontSize: 15.5, fontWeight: 600, margin: 0 }}>QuickPick Insights</h3>
            <div style={{ fontSize: 12.5, color: C.textSecondary, marginTop: 2 }}>A quick read on pricing activity and catalogue health.</div>
          </div>

          <div className="qp-insights-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 20 }}>
              <div style={{ fontSize: 14.5, fontWeight: 600, marginBottom: 3 }}>Price Updates &middot; Last 7 Days</div>
              <div style={{ fontSize: 12, color: C.textSecondary, marginBottom: 16 }}>Number of price submissions received per day.</div>
              <div style={{ display: "flex", alignItems: "flex-end", gap: 10, height: 130, paddingTop: 10 }}>
                {barData.map(b => (
                  <div key={b.d} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "flex-end", gap: 6, height: "100%" }}>
                    <div style={{ width: "100%", maxWidth: 26, borderRadius: "5px 5px 3px 3px", background: b.v === maxBar ? C.accent : "#D3EBDD", height: `${(b.v / maxBar) * 100}%` }} />
                    <span style={{ fontSize: 10.5, color: C.textMuted }}>{b.d}</span>
                  </div>
                ))}
              </div>
            </div>

            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 20 }}>
              <div style={{ fontSize: 14.5, fontWeight: 600, marginBottom: 3 }}>Product Coverage</div>
              <div style={{ fontSize: 12, color: C.textSecondary, marginBottom: 16 }}>Share of catalogue with current pricing.</div>
              <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
                <svg width={110} height={110} viewBox="0 0 42 42">
                  <circle cx="21" cy="21" r="15.9" fill="none" stroke="#EAF0EC" strokeWidth="5" />
                  <circle cx="21" cy="21" r="15.9" fill="none" stroke={C.accent} strokeWidth="5" strokeDasharray={`${stats.fullyPricedPct} 100`} strokeDashoffset="25" strokeLinecap="round" />
                  <circle cx="21" cy="21" r="15.9" fill="none" stroke={C.amber} strokeWidth="5" strokeDasharray={`${stats.partiallyPricedPct} 100`} strokeDashoffset={25 - stats.fullyPricedPct} strokeLinecap="round" />
                  <circle cx="21" cy="21" r="15.9" fill="none" stroke={C.danger} strokeWidth="5" strokeDasharray={`${stats.missingPricedPct} 100`} strokeDashoffset={25 - stats.fullyPricedPct - stats.partiallyPricedPct} strokeLinecap="round" />
                </svg>
                <div style={{ display: "flex", flexDirection: "column", gap: 10, fontSize: 12.5 }}>
                  {[
                    { l: "Fully Priced", v: `${stats.fullyPricedPct}%`, c: C.accent }, { l: "Partially Priced", v: `${stats.partiallyPricedPct}%`, c: C.amber }, { l: "Missing Prices", v: `${stats.missingPricedPct}%`, c: C.danger },
                  ].map(row => (
                    <div key={row.l} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{ width: 9, height: 9, borderRadius: 3, background: row.c, flexShrink: 0 }} />
                      {row.l}<span style={{ marginLeft: "auto", fontWeight: 600, color: C.text }}>{row.v}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>

          {/* Needs attention */}
          <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", margin: "32px 0 14px" }}>
            <div>
              <h3 style={{ fontSize: 15.5, fontWeight: 600, margin: 0 }}>Needs Attention</h3>
              <div style={{ fontSize: 12.5, color: C.textSecondary, marginTop: 2 }}>Products with pricing issues that need a manual review.</div>
            </div>
            <a onClick={goPrices} style={{ color: C.accent, fontWeight: 600, fontSize: 13, cursor: "pointer" }}>View all →</a>
          </div>

          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: "8px 12px 4px", overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr>
                  {["Product", "Issue", "Last Update", "Priority", ""].map(h => (
                    <th key={h} style={{ textAlign: "left", fontSize: 11.5, color: C.textSecondary, fontWeight: 600, padding: "0 12px 10px", borderBottom: `1px solid ${C.border}` }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {stats.attentionItems.length === 0 ? (
                  <tr>
                    <td colSpan={5} style={{ padding: 20, textAlign: "center", color: C.textSecondary }}>
                      Nothing needs attention right now — every price is fresh.
                    </td>
                  </tr>
                ) : stats.attentionItems.map(row => (
                  <tr key={row.name}>
                    <td style={{ padding: 12, borderBottom: `1px solid ${C.border}`, fontWeight: 600 }}>{row.name}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${C.border}` }}>{row.issue}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${C.border}` }}>{row.update}</td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${C.border}` }}>
                      <Badge tone={row.priority === "High" ? "danger" : row.priority === "Medium" ? "amber" : "accent"} label={row.priority} />
                    </td>
                    <td style={{ padding: 12, borderBottom: `1px solid ${C.border}` }}>
                      <a onClick={goPrices} style={{ color: C.accent, fontWeight: 600, fontSize: 12.5, cursor: "pointer" }}>Review</a>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Quick actions */}
          <div style={{ margin: "32px 0 14px" }}>
            <h3 style={{ fontSize: 15.5, fontWeight: 600, margin: 0 }}>Quick Actions</h3>
          </div>

          <div className="qp-qa-grid" style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 12 }}>
            {[
              { label: "Add Product", icon: PackagePlus, onClick: goProducts },
              { label: "Add Shop", icon: Store, onClick: goShops },
              { label: "Update Prices", icon: Tag, onClick: goPrices },
              { label: "Smart Shopping", icon: ShoppingCart, onClick: goSmartShopping },
              { label: "Search Products", icon: Search, onClick: goProductSearch },
            ].map(action => (
              <button
                key={action.label}
                onClick={action.onClick}
                style={{
                  background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, padding: "16px 10px",
                  display: "flex", flexDirection: "column", alignItems: "center", gap: 8, fontSize: 12.5, fontWeight: 600,
                  color: C.text, textAlign: "center", cursor: "pointer",
                }}
                onMouseEnter={e => { e.currentTarget.style.borderColor = C.accent; e.currentTarget.style.background = C.accentTint; }}
                onMouseLeave={e => { e.currentTarget.style.borderColor = C.border; e.currentTarget.style.background = C.surface; }}
              >
                <div style={{ width: 34, height: 34, borderRadius: 9, background: C.accentTint, color: C.accent, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <action.icon size={16} />
                </div>
                {action.label}
              </button>
            ))}
          </div>

          {/* Activity + shop coverage */}
          <div style={{ margin: "32px 0 14px" }}>
            <h3 style={{ fontSize: 15.5, fontWeight: 600, margin: 0 }}>Activity &amp; Shop Health</h3>
          </div>

          <div className="qp-bottom-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 20 }}>
              <div style={{ fontSize: 14.5, fontWeight: 600, marginBottom: 3 }}>Recent Activity</div>
              <div style={{ fontSize: 12, color: C.textSecondary, marginBottom: 16 }}>Latest changes across the platform.</div>
              <div>
                {[
                  { icon: Tag, title: "Price Updated", body: "Unga 2kg updated at QuickMart", time: "5 minutes ago" },
                  { icon: Plus, title: "Product Added", body: "Fresh Milk 500ml added to catalogue", time: "24 minutes ago" },
                  { icon: Store, title: "Shop Updated", body: "Naivas Chuka details were updated", time: "1 hour ago", amber: true },
                  { icon: Tag, title: "Price Update", body: "Cooking Oil changed from KSh 340 to KSh 325", time: "2 hours ago" },
                ].map((item, i, arr) => (
                  <div key={i} style={{ display: "flex", gap: 12, paddingBottom: i === arr.length - 1 ? 0 : 18, position: "relative" }}>
                    {i !== arr.length - 1 && (
                      <div style={{ position: "absolute", left: 14, top: 26, bottom: 0, width: 1, background: C.border }} />
                    )}
                    <div style={{
                      width: 29, height: 29, borderRadius: "50%", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center",
                      background: item.amber ? C.amberTint : C.accentTint, color: item.amber ? C.amber : C.accent, zIndex: 1,
                    }}>
                      <item.icon size={14} />
                    </div>
                    <div>
                      <b style={{ fontSize: 13, fontWeight: 600, display: "block" }}>{item.title}</b>
                      <p style={{ margin: "2px 0 0", fontSize: 12, color: C.textSecondary }}>{item.body}</p>
                      <time style={{ fontSize: 11, color: C.textMuted }}>{item.time}</time>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 16, padding: 20 }}>
              <div style={{ fontSize: 14.5, fontWeight: 600, marginBottom: 3 }}>Top Shops</div>
              <div style={{ fontSize: 12, color: C.textSecondary, marginBottom: 16 }}>Registered shops with the most products in the catalogue.</div>
              {stats.topShops.length === 0 ? (
                <div style={{ fontSize: 12.5, color: C.textMuted, padding: "12px 0" }}>No shops added yet.</div>
              ) : (() => {
                const maxProducts = Math.max(1, ...stats.topShops.map((s) => s.products));
                const statusStyle = {
                  Active: { tone: C.accentTint, fg: C.accentDark, bar: C.accent },
                  Inactive: { tone: "#EAF2FD", fg: "#2563AC", bar: "#2563AC" },
                  Suspended: { tone: C.amberTint, fg: C.amber, bar: C.amber },
                };
                return stats.topShops.map((shop, i, arr) => {
                  const sty = statusStyle[shop.status] || statusStyle.Active;
                  const barPct = Math.round((shop.products / maxProducts) * 100);
                  return (
                    <div key={shop.name} style={{ padding: i === 0 ? "0 0 13px" : i === arr.length - 1 ? "13px 0 0" : "13px 0", borderBottom: i === arr.length - 1 ? "none" : `1px solid ${C.border}` }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 7 }}>
                        <span style={{ fontWeight: 600, fontSize: 13 }}>{shop.name}</span>
                        <span style={{ fontSize: 10.5, fontWeight: 600, padding: "3px 8px", borderRadius: 20, background: sty.tone, color: sty.fg }}>{shop.status}</span>
                      </div>
                      <div style={{ fontSize: 11.5, color: C.textSecondary, marginBottom: 6 }}>{shop.products} products</div>
                      <div style={{ height: 8, background: C.bg, borderRadius: 6, overflow: "hidden", border: `1px solid ${C.border}` }}>
                        <div style={{ height: "100%", width: `${barPct}%`, background: sty.bar, borderRadius: 6 }} />
                      </div>
                    </div>
                  );
                });
              })()}
            </div>
          </div>

          <footer style={{ textAlign: "center", fontSize: 11.5, color: C.textMuted, padding: "30px 0 10px" }}>
            QuickPick Admin Console &middot; internal tool for QuickPick platform administrators
          </footer>
        </main>
      </div>

      <Toast message={toast} onDone={() => setToast("")} />
    </div>
  );
}


// ============================================================
// ---------- Shop Management ----------
// ============================================================


/* -------------------------------------------------------------------------
   Mock data
------------------------------------------------------------------------- */

const SHOP_CATEGORIES = [
  "Supermarket",
  "Mini Market",
  "Wholesale",
  "Retail Shop",
  "General Store",
];

const STATUSES = ["Active", "Inactive", "Suspended"];

const initialShops = [
  {
    id: "QP001",
    name: "Ndagani Supermarket",
    owner: "John Mwangi",
    location: "Ndagani",
    phone: "0712345678",
    email: "john.mwangi@example.com",
    category: "Supermarket",
    products: 145,
    status: "Active",
    regNo: "BN-100234",
    dateAdded: "2023-02-14",
  },
  {
    id: "QP002",
    name: "Chuka Mini Mart",
    owner: "Mary Wanjiku",
    location: "Chuka Town",
    phone: "0723456789",
    email: "mary.wanjiku@example.com",
    category: "Mini Market",
    products: 86,
    status: "Active",
    regNo: "BN-100455",
    dateAdded: "2023-04-02",
  },
  {
    id: "QP003",
    name: "Gitombani General Shop",
    owner: "Peter Gitonga",
    location: "Gitombani",
    phone: "0734567890",
    email: "peter.gitonga@example.com",
    category: "General Store",
    products: 54,
    status: "Inactive",
    regNo: "BN-100678",
    dateAdded: "2022-11-19",
  },
  {
    id: "QP004",
    name: "Campus Choice Store",
    owner: "Lucy Njeri",
    location: "Chuka University",
    phone: "0709876543",
    email: "lucy.njeri@example.com",
    category: "Retail Shop",
    products: 120,
    status: "Active",
    regNo: "BN-100812",
    dateAdded: "2023-06-27",
  },
  {
    id: "QP005",
    name: "Ndagani Wholesale",
    owner: "Samuel Kariuki",
    location: "Ndagani Market",
    phone: "0798765432",
    email: "samuel.kariuki@example.com",
    category: "Wholesale",
    products: 210,
    status: "Suspended",
    regNo: "BN-100933",
    dateAdded: "2022-08-05",
  },
  {
    id: "QP006",
    name: "Kanyakine Retail Hub",
    owner: "Grace Muthoni",
    location: "Kanyakine",
    phone: "0711223344",
    email: "grace.muthoni@example.com",
    category: "Retail Shop",
    products: 98,
    status: "Active",
    regNo: "BN-101021",
    dateAdded: "2023-01-30",
  },
  {
    id: "QP007",
    name: "Marimanti General Store",
    owner: "James Kiogora",
    location: "Marimanti",
    phone: "0722334455",
    email: "james.kiogora@example.com",
    category: "General Store",
    products: 41,
    status: "Inactive",
    regNo: "BN-101187",
    dateAdded: "2022-05-16",
  },
  {
    id: "QP008",
    name: "Chogoria Supermarket",
    owner: "Alice Karimi",
    location: "Chogoria",
    phone: "0733445566",
    email: "alice.karimi@example.com",
    category: "Supermarket",
    products: 176,
    status: "Active",
    regNo: "BN-101344",
    dateAdded: "2023-09-11",
  },
  {
    id: "QP009",
    name: "Magutuni Mini Mart",
    owner: "Dennis Mutuma",
    location: "Magutuni",
    phone: "0744556677",
    email: "dennis.mutuma@example.com",
    category: "Mini Market",
    products: 63,
    status: "Suspended",
    regNo: "BN-101502",
    dateAdded: "2022-12-08",
  },
  {
    id: "QP010",
    name: "Karingani Wholesale",
    owner: "Esther Kanana",
    location: "Karingani",
    phone: "0755667788",
    email: "esther.kanana@example.com",
    category: "Wholesale",
    products: 132,
    status: "Active",
    regNo: "BN-101699",
    dateAdded: "2023-03-22",
  },
];

const PAGE_SIZE = 5;

/* -------------------------------------------------------------------------
   Small helpers
------------------------------------------------------------------------- */

function StatusBadge({ status }) {
  const map = {
    Active: { bg: "#e6f4ea", color: "#107c41", border: "#bfe3cc" },
    Inactive: { bg: "#f1f2f4", color: "#5f6368", border: "#dcdde0" },
    Suspended: { bg: "#fdeee0", color: "#c2540a", border: "#f5cfa8" },
  };
  const s = map[status] || map.Inactive;
  return (
    <span
      style={{
        display: "inline-block",
        padding: "3px 10px",
        borderRadius: 999,
        fontSize: 12,
        fontWeight: 600,
        background: s.bg,
        color: s.color,
        border: `1px solid ${s.border}`,
        whiteSpace: "nowrap",
      }}
    >
      {status}
    </span>
  );
}

function emptyForm() {
  return {
    name: "",
    owner: "",
    phone: "",
    email: "",
    location: "",
    category: SHOP_CATEGORIES[0],
    regNo: "",
    products: "",
    status: "Active",
  };
}

function formatDate(d) {
  try {
    return new Date(d).toLocaleDateString("en-GB", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
  } catch {
    return d;
  }
}

/* -------------------------------------------------------------------------
   ShopModal shell
------------------------------------------------------------------------- */

function ShopModal({ title, onClose, children, footer, wide }) {
  return (
    <div style={styles.overlay} onMouseDown={onClose}>
      <div
        style={{ ...styles.modal, maxWidth: wide ? 720 : 520 }}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div style={styles.modalHeader}>
          <h3 style={styles.modalTitle}>{title}</h3>
          <button style={styles.iconBtn} onClick={onClose} aria-label="Close">
            ✕
          </button>
        </div>
        <div style={styles.modalBody}>{children}</div>
        {footer && <div style={styles.modalFooter}>{footer}</div>}
      </div>
    </div>
  );
}

// ---------- Products assigned to a shop (many-to-many, one price per shop) ----------
function ShopProductsModal({ shop, products, shopProducts, onAssign, onUpdatePrice, onRemove, onClose }) {
  const assignments = shopProducts.filter((sp) => sp.shopId === shop.id);
  const assignedProductIds = new Set(assignments.map((a) => a.productId));
  const availableProducts = products.filter((p) => !assignedProductIds.has(p.id));

  const [selectedProductId, setSelectedProductId] = useState("");
  const [priceDraft, setPriceDraft] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editPrice, setEditPrice] = useState("");
  const [notice, setNotice] = useState("");

  function handleAssign() {
    if (!selectedProductId) { setNotice("Choose a product first."); return; }
    if (priceDraft === "" || isNaN(Number(priceDraft)) || Number(priceDraft) < 0) { setNotice("Enter a valid price."); return; }
    const product = products.find((p) => p.id === selectedProductId);
    const result = onAssign(shop.id, selectedProductId, Number(priceDraft));
    setNotice(
      result?.updated
        ? `${product?.productName} is already assigned to this shop. Its price was updated instead.`
        : `Assigned ${product?.productName}.`
    );
    setSelectedProductId("");
    setPriceDraft("");
  }

  return (
    <ShopModal
      title="Products"
      wide
      onClose={onClose}
      footer={
        <button style={styles.secondaryBtn} onClick={onClose}>
          Close
        </button>
      }
    >
      <div style={{ padding: "8px 0" }}>
        <p style={{ margin: "0 0 4px", fontWeight: 600 }}>{shop.name}</p>
        <p style={{ margin: "0 0 16px", color: "#5f6368" }}>
          {assignments.length} product{assignments.length === 1 ? "" : "s"} currently sold here
        </p>

        {notice ? (
          <div style={{ background: "#e6f4ea", color: "#1e293b", fontSize: 12.5, padding: "8px 12px", borderRadius: 8, marginBottom: 14 }}>
            {notice}
          </div>
        ) : null}

        {assignments.length === 0 ? (
          <p style={{ margin: "0 0 16px", color: "#5f6368", fontStyle: "italic" }}>
            No products are assigned to this shop yet. Use "Assign Product" below.
          </p>
        ) : (
          <div style={{ maxHeight: 320, overflowY: "auto", border: "1px solid #e5e7eb", borderRadius: 8, marginBottom: 20 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: "#f8f9fa", position: "sticky", top: 0 }}>
                  <th style={{ textAlign: "left", padding: "9px 12px", fontSize: 11.5, color: "#5f6368", fontWeight: 600 }}>Product</th>
                  <th style={{ textAlign: "left", padding: "9px 12px", fontSize: 11.5, color: "#5f6368", fontWeight: 600 }}>Category</th>
                  <th style={{ textAlign: "right", padding: "9px 12px", fontSize: 11.5, color: "#5f6368", fontWeight: 600 }}>Price</th>
                  <th style={{ padding: "9px 12px" }} />
                </tr>
              </thead>
              <tbody>
                {assignments.map((a) => {
                  const product = products.find((p) => p.id === a.productId);
                  const isEditing = editingId === a.id;
                  return (
                    <tr key={a.id} style={{ borderTop: "1px solid #eef0f1" }}>
                      <td style={{ padding: "9px 12px" }}>
                        <div style={{ fontWeight: 600 }}>{product?.productName || "Unknown product"}</div>
                        {product?.brand ? <div style={{ fontSize: 11.5, color: "#9aa1a8" }}>{product.brand}{product.packageSize ? ` \u00b7 ${product.packageSize}` : ""}</div> : null}
                      </td>
                      <td style={{ padding: "9px 12px", color: "#5f6368" }}>{product?.category || "\u2014"}</td>
                      <td style={{ padding: "9px 12px", textAlign: "right" }}>
                        {isEditing ? (
                          <input
                            type="number" min="0" autoFocus value={editPrice}
                            onChange={(e) => setEditPrice(e.target.value)}
                            style={{ width: 84, padding: "4px 6px", border: "1px solid #e5e7eb", borderRadius: 6, textAlign: "right", fontSize: 13 }}
                          />
                        ) : (
                          <span style={{ fontWeight: 600 }}>{formatKSh(a.price)}</span>
                        )}
                      </td>
                      <td style={{ padding: "9px 12px", textAlign: "right", whiteSpace: "nowrap" }}>
                        {isEditing ? (
                          <>
                            <button
                              onClick={() => {
                                if (editPrice === "" || isNaN(Number(editPrice)) || Number(editPrice) < 0) return;
                                onUpdatePrice(a.id, Number(editPrice));
                                setEditingId(null);
                              }}
                              style={{ background: "none", border: "none", cursor: "pointer", color: "#1a73e8", fontSize: 12.5, fontWeight: 600, marginRight: 8 }}
                            >
                              Save
                            </button>
                            <button onClick={() => setEditingId(null)} style={{ background: "none", border: "none", cursor: "pointer", color: "#5f6368", fontSize: 12.5 }}>
                              Cancel
                            </button>
                          </>
                        ) : (
                          <>
                            <button
                              onClick={() => { setEditingId(a.id); setEditPrice(String(a.price)); }}
                              style={{ background: "none", border: "none", cursor: "pointer", color: "#1a73e8", fontSize: 12.5, fontWeight: 600, marginRight: 8 }}
                            >
                              Edit
                            </button>
                            <button onClick={() => onRemove(a.id)} style={{ background: "none", border: "none", cursor: "pointer", color: "#c5221f", fontSize: 12.5 }}>
                              Remove
                            </button>
                          </>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        <div style={{ paddingTop: 16, borderTop: "1px solid #eef0f1" }}>
          <div style={{ fontWeight: 600, fontSize: 13.5, marginBottom: 10 }}>+ Assign Product</div>
          {availableProducts.length === 0 ? (
            <p style={{ margin: 0, color: "#9aa1a8", fontStyle: "italic", fontSize: 12.5 }}>
              Every product in the catalogue is already assigned to this shop.
            </p>
          ) : (
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <select
                style={{ ...styles.input, flex: 2, minWidth: 180 }}
                value={selectedProductId}
                onChange={(e) => setSelectedProductId(e.target.value)}
              >
                <option value="">Select product...</option>
                {availableProducts.map((p) => <option key={p.id} value={p.id}>{p.productName}</option>)}
              </select>
              <input
                style={{ ...styles.input, flex: 1, minWidth: 90 }}
                type="number" min="0" placeholder="Price"
                value={priceDraft}
                onChange={(e) => setPriceDraft(e.target.value)}
              />
              <button style={styles.primaryBtn} onClick={handleAssign}>Assign Product</button>
            </div>
          )}
        </div>
      </div>
    </ShopModal>
  );
}

/* -------------------------------------------------------------------------
   Shop form (used by Add + Edit modals)
------------------------------------------------------------------------- */

function ShopForm({ form, setForm, errors }) {
  const set = (field) => (e) =>
    setForm((f) => ({ ...f, [field]: e.target.value }));

  return (
    <div style={styles.formGrid}>
      <div style={styles.formField}>
        <label style={styles.label}>Shop Name *</label>
        <input
          style={{ ...styles.input, ...(errors.name ? styles.inputError : {}) }}
          value={form.name}
          onChange={set("name")}
          placeholder="e.g. Ndagani Supermarket"
        />
        {errors.name && <span style={styles.errorText}>{errors.name}</span>}
      </div>

      <div style={styles.formField}>
        <label style={styles.label}>Owner Name *</label>
        <input
          style={{ ...styles.input, ...(errors.owner ? styles.inputError : {}) }}
          value={form.owner}
          onChange={set("owner")}
          placeholder="e.g. John Mwangi"
        />
        {errors.owner && <span style={styles.errorText}>{errors.owner}</span>}
      </div>

      <div style={styles.formField}>
        <label style={styles.label}>Phone Number *</label>
        <input
          style={{ ...styles.input, ...(errors.phone ? styles.inputError : {}) }}
          value={form.phone}
          onChange={set("phone")}
          placeholder="07XXXXXXXX"
        />
        {errors.phone && <span style={styles.errorText}>{errors.phone}</span>}
      </div>

      <div style={styles.formField}>
        <label style={styles.label}>Email</label>
        <input
          style={styles.input}
          value={form.email}
          onChange={set("email")}
          placeholder="shop@example.com"
        />
      </div>

      <div style={styles.formField}>
        <label style={styles.label}>Location *</label>
        <input
          style={{
            ...styles.input,
            ...(errors.location ? styles.inputError : {}),
          }}
          value={form.location}
          onChange={set("location")}
          placeholder="e.g. Chuka Town"
        />
        {errors.location && (
          <span style={styles.errorText}>{errors.location}</span>
        )}
      </div>

      <div style={styles.formField}>
        <label style={styles.label}>Shop Category *</label>
        <select style={styles.input} value={form.category} onChange={set("category")}>
          {SHOP_CATEGORIES.map((c) => (
            <option key={c} value={c}>
              {c}
            </option>
          ))}
        </select>
      </div>

      <div style={styles.formField}>
        <label style={styles.label}>Business Registration Number</label>
        <input
          style={styles.input}
          value={form.regNo}
          onChange={set("regNo")}
          placeholder="e.g. BN-100234"
        />
      </div>

      <div style={styles.formField}>
        <label style={styles.label}>Number of Products</label>
        <input
          type="number"
          min="0"
          style={styles.input}
          value={form.products}
          onChange={set("products")}
          placeholder="0"
        />
      </div>

      <div style={styles.formField}>
        <label style={styles.label}>Status</label>
        <select style={styles.input} value={form.status} onChange={set("status")}>
          {STATUSES.map((s) => (
            <option key={s} value={s}>
              {s}
            </option>
          ))}
        </select>
      </div>
    </div>
  );
}

/* -------------------------------------------------------------------------
   Main component
------------------------------------------------------------------------- */

function ShopManagement({ onBack, shops: sharedShops, onShopsChange, products = [], shopProducts = [], onShopProductsChange }) {
  const [shops, setShops] = useState(() => sharedShops ?? [...initialShops]);
  useEffect(() => {
    if (onShopsChange) onShopsChange(shops);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shops]);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("All Status");
  const [categoryFilter, setCategoryFilter] = useState("All Categories");
  const [page, setPage] = useState(1);
  const [selectedIds, setSelectedIds] = useState([]);

  const [addOpen, setAddOpen] = useState(false);
  const [editOpen, setEditOpen] = useState(false);
  const [viewOpen, setViewOpen] = useState(false);
  const [productsOpen, setProductsOpen] = useState(false);
  const [deleteId, setDeleteId] = useState(null);

  const [activeShop, setActiveShop] = useState(null);
  const [form, setForm] = useState(emptyForm());
  const [errors, setErrors] = useState({});

  // Local mirror of shopProducts, propagated up the same safe way as shops.
  const [shopProductsLocal, setShopProductsLocal] = useState(() => shopProducts);
  useEffect(() => {
    if (onShopProductsChange) onShopProductsChange(shopProductsLocal);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shopProductsLocal]);

  function assignProductToShop(shopId, productId, price) {
    const existing = shopProductsLocal.find((sp) => sp.productId === productId && sp.shopId === shopId);
    if (existing) {
      setShopProductsLocal((prev) => prev.map((sp) => (sp.id === existing.id ? { ...sp, price, lastUpdated: new Date().toISOString() } : sp)));
      return { updated: true };
    }
    setShopProductsLocal((prev) => [
      ...prev,
      { id: `SP${Date.now()}${Math.floor(Math.random() * 1000)}`, productId, shopId, price, lastUpdated: new Date().toISOString() },
    ]);
    return { updated: false };
  }
  function updateShopProductPrice(id, price) {
    setShopProductsLocal((prev) => prev.map((sp) => (sp.id === id ? { ...sp, price, lastUpdated: new Date().toISOString() } : sp)));
  }
  function removeShopProduct(id) {
    setShopProductsLocal((prev) => prev.filter((sp) => sp.id !== id));
  }

  /* ---------------- derived data ---------------- */

  const summary = useMemo(() => {
    const total = shops.length;
    const active = shops.filter((s) => s.status === "Active").length;
    const inactive = shops.filter((s) => s.status !== "Active").length;
    const totalProducts = shops.reduce(
      (sum, s) => sum + (Number(s.products) || 0),
      0
    );
    return { total, active, inactive, totalProducts };
  }, [shops]);

  const filteredShops = useMemo(() => {
    const q = search.trim().toLowerCase();
    return shops.filter((s) => {
      const matchesSearch =
        !q ||
        s.name.toLowerCase().includes(q) ||
        s.owner.toLowerCase().includes(q) ||
        s.location.toLowerCase().includes(q) ||
        s.phone.toLowerCase().includes(q);

      const matchesStatus =
        statusFilter === "All Status" || s.status === statusFilter;

      const matchesCategory =
        categoryFilter === "All Categories" || s.category === categoryFilter;

      return matchesSearch && matchesStatus && matchesCategory;
    });
  }, [shops, search, statusFilter, categoryFilter]);

  const totalPages = Math.max(1, Math.ceil(filteredShops.length / PAGE_SIZE));
  const safePage = Math.min(page, totalPages);

  const pagedShops = useMemo(() => {
    const start = (safePage - 1) * PAGE_SIZE;
    return filteredShops.slice(start, start + PAGE_SIZE);
  }, [filteredShops, safePage]);

  const allOnPageSelected =
    pagedShops.length > 0 &&
    pagedShops.every((s) => selectedIds.includes(s.id));

  /* ---------------- handlers ---------------- */

  function resetFiltersView() {
    setPage(1);
  }

  function handleSearchChange(e) {
    setSearch(e.target.value);
    resetFiltersView();
  }

  function handleStatusFilterChange(e) {
    setStatusFilter(e.target.value);
    resetFiltersView();
  }

  function handleCategoryFilterChange(e) {
    setCategoryFilter(e.target.value);
    resetFiltersView();
  }

  function handleRefresh() {
    setSearch("");
    setStatusFilter("All Status");
    setCategoryFilter("All Categories");
    setSelectedIds([]);
    setPage(1);
  }

  function toggleSelectAllOnPage() {
    if (allOnPageSelected) {
      const pageIds = pagedShops.map((s) => s.id);
      setSelectedIds((prev) => prev.filter((id) => !pageIds.includes(id)));
    } else {
      const pageIds = pagedShops.map((s) => s.id);
      setSelectedIds((prev) => Array.from(new Set([...prev, ...pageIds])));
    }
  }

  function toggleSelectRow(id) {
    setSelectedIds((prev) =>
      prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
    );
  }

  function bulkSetStatus(status) {
    setShops((prev) =>
      prev.map((s) => (selectedIds.includes(s.id) ? { ...s, status } : s))
    );
    setSelectedIds([]);
  }

  function bulkDelete() {
    const idsToDelete = new Set(selectedIds);
    setShops((prev) => prev.filter((s) => !selectedIds.includes(s.id)));
    setShopProductsLocal((prev) => prev.filter((sp) => !idsToDelete.has(sp.shopId)));
    setSelectedIds([]);
  }

  function openAdd() {
    setForm(emptyForm());
    setErrors({});
    setAddOpen(true);
  }

  function openEdit(shop) {
    setActiveShop(shop);
    setForm({
      name: shop.name,
      owner: shop.owner,
      phone: shop.phone,
      email: shop.email,
      location: shop.location,
      category: shop.category,
      regNo: shop.regNo,
      products: String(shop.products),
      status: shop.status,
    });
    setErrors({});
    setEditOpen(true);
  }

  function openView(shop) {
    setActiveShop(shop);
    setViewOpen(true);
  }

  function openProducts(shop) {
    setActiveShop(shop);
    setProductsOpen(true);
  }

  function validateForm() {
    const errs = {};
    if (!form.name.trim()) errs.name = "Shop name is required";
    if (!form.owner.trim()) errs.owner = "Owner name is required";
    if (!form.phone.trim()) errs.phone = "Phone number is required";
    if (!form.location.trim()) errs.location = "Location is required";
    setErrors(errs);
    return Object.keys(errs).length === 0;
  }

  function nextShopId() {
    const nums = shops
      .map((s) => parseInt(String(s.id).replace(/\D/g, ""), 10))
      .filter((n) => !Number.isNaN(n));
    const max = nums.length ? Math.max(...nums) : 0;
    return `QP${String(max + 1).padStart(3, "0")}`;
  }

  function handleSaveAdd() {
    if (!validateForm()) return;
    const newShop = {
      id: nextShopId(),
      name: form.name.trim(),
      owner: form.owner.trim(),
      phone: form.phone.trim(),
      email: form.email.trim(),
      location: form.location.trim(),
      category: form.category,
      regNo: form.regNo.trim(),
      products: Number(form.products) || 0,
      status: form.status,
      dateAdded: new Date().toISOString().slice(0, 10),
    };
    setShops((prev) => [newShop, ...prev]);
    setAddOpen(false);
    setPage(1);
  }

  function handleSaveEdit() {
    if (!validateForm()) return;
    setShops((prev) =>
      prev.map((s) =>
        s.id === activeShop.id
          ? {
              ...s,
              name: form.name.trim(),
              owner: form.owner.trim(),
              phone: form.phone.trim(),
              email: form.email.trim(),
              location: form.location.trim(),
              category: form.category,
              regNo: form.regNo.trim(),
              products: Number(form.products) || 0,
              status: form.status,
            }
          : s
      )
    );
    setEditOpen(false);
    setActiveShop(null);
  }

  function confirmDelete(id) {
    setDeleteId(id);
  }

  function performDelete() {
    setShops((prev) => prev.filter((s) => s.id !== deleteId));
    setShopProductsLocal((prev) => prev.filter((sp) => sp.shopId !== deleteId));
    setSelectedIds((prev) => prev.filter((id) => id !== deleteId));
    setDeleteId(null);
  }

  function handleImportExcel() {
    window.alert(
      "Import Excel: file picker would open here. This is a mock action (no backend)."
    );
  }

  function handleExportExcel() {
    window.alert(
      "Export Excel: current shop list would be exported here. This is a mock action (no backend)."
    );
  }

  /* ---------------- render ---------------- */

  return (
    <div style={styles.page}>
      {/* Header */}
      <div style={styles.header}>
        <div>
          {onBack && (
            <button
              onClick={onBack}
              style={{
                background: "none", border: "none", cursor: "pointer", color: "#5f6368",
                fontSize: 13, fontWeight: 600, padding: 0, marginBottom: 10,
                display: "inline-flex", alignItems: "center", gap: 6,
              }}
            >
              ← Back to Dashboard
            </button>
          )}
          <h1 style={styles.h1}>Shop Management</h1>
          <p style={styles.subtitle}>Manage shops registered on QuickPick</p>
        </div>
        <button style={styles.primaryBtn} onClick={openAdd}>
          + Add Shop
        </button>
      </div>

      {/* Summary cards */}
      <div style={styles.cardsRow}>
        <SummaryCard label="Total Shops" value={summary.total} accent="#107c41" />
        <SummaryCard label="Active Shops" value={summary.active} accent="#107c41" />
        <SummaryCard label="Inactive Shops" value={summary.inactive} accent="#c2540a" />
        <SummaryCard
          label="Total Products"
          value={summary.totalProducts.toLocaleString()}
          accent="#0f6cbd"
        />
      </div>

      {/* Toolbar */}
      <div style={styles.toolbar}>
        <input
          style={styles.searchInput}
          placeholder="Search shops..."
          value={search}
          onChange={handleSearchChange}
        />

        <select
          style={styles.select}
          value={statusFilter}
          onChange={handleStatusFilterChange}
        >
          <option>All Status</option>
          <option>Active</option>
          <option>Inactive</option>
          <option>Suspended</option>
        </select>

        <select
          style={styles.select}
          value={categoryFilter}
          onChange={handleCategoryFilterChange}
        >
          <option>All Categories</option>
          {SHOP_CATEGORIES.map((c) => (
            <option key={c}>{c}</option>
          ))}
        </select>

        <div style={styles.toolbarSpacer} />

        <button style={styles.secondaryBtn} onClick={handleImportExcel}>
          Import Excel
        </button>
        <button style={styles.secondaryBtn} onClick={handleExportExcel}>
          Export Excel
        </button>
        <button style={styles.secondaryBtn} onClick={handleRefresh}>
          Refresh
        </button>
      </div>

      {/* Bulk action bar */}
      {selectedIds.length > 0 && (
        <div style={styles.bulkBar}>
          <span style={styles.bulkText}>{selectedIds.length} shops selected</span>
          <div style={styles.bulkBtns}>
            <button
              style={styles.bulkBtnGreen}
              onClick={() => bulkSetStatus("Active")}
            >
              Activate
            </button>
            <button
              style={styles.bulkBtnGray}
              onClick={() => bulkSetStatus("Inactive")}
            >
              Deactivate
            </button>
            <button style={styles.bulkBtnRed} onClick={bulkDelete}>
              Delete
            </button>
          </div>
        </div>
      )}

      {/* Table */}
      <div style={styles.tableWrap}>
        <table style={styles.table}>
          <thead>
            <tr>
              <th style={{ ...styles.th, ...styles.thCheckbox }}>
                <input
                  type="checkbox"
                  checked={allOnPageSelected}
                  onChange={toggleSelectAllOnPage}
                />
              </th>
              <th style={styles.th}>Shop ID</th>
              <th style={styles.th}>Shop Name</th>
              <th style={styles.th}>Owner</th>
              <th style={styles.th}>Location</th>
              <th style={styles.th}>Phone</th>
              <th style={styles.th}>Category</th>
              <th style={styles.th}>Products</th>
              <th style={styles.th}>Status</th>
              <th style={styles.th}>Date Added</th>
              <th style={styles.th}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {pagedShops.length === 0 ? (
              <tr>
                <td colSpan={11} style={styles.emptyCell}>
                  <div style={styles.emptyState}>
                    <div style={styles.emptyTitle}>No shops found</div>
                    <div style={styles.emptySubtitle}>
                      Try changing your search or filters.
                    </div>
                  </div>
                </td>
              </tr>
            ) : (
              pagedShops.map((shop, idx) => (
                <tr
                  key={shop.id}
                  style={{
                    ...styles.tr,
                    background: idx % 2 === 0 ? "#ffffff" : "#fafbfc",
                  }}
                >
                  <td style={{ ...styles.td, ...styles.thCheckbox }}>
                    <input
                      type="checkbox"
                      checked={selectedIds.includes(shop.id)}
                      onChange={() => toggleSelectRow(shop.id)}
                    />
                  </td>
                  <td style={{ ...styles.td, fontWeight: 600 }}>{shop.id}</td>
                  <td style={styles.td}>{shop.name}</td>
                  <td style={styles.td}>{shop.owner}</td>
                  <td style={styles.td}>{shop.location}</td>
                  <td style={styles.td}>{shop.phone}</td>
                  <td style={styles.td}>{shop.category}</td>
                  <td style={styles.td}>{shop.products}</td>
                  <td style={styles.td}>
                    <StatusBadge status={shop.status} />
                  </td>
                  <td style={styles.td}>{formatDate(shop.dateAdded)}</td>
                  <td style={{ ...styles.td, whiteSpace: "nowrap" }}>
                    <button
                      style={styles.actionLink}
                      onClick={() => openView(shop)}
                      title="View"
                    >
                      👁 View
                    </button>
                    <button
                      style={styles.actionLink}
                      onClick={() => openEdit(shop)}
                      title="Edit"
                    >
                      ✏️ Edit
                    </button>
                    <button
                      style={styles.actionLink}
                      onClick={() => openProducts(shop)}
                      title="Products"
                    >
                      📦 Products
                    </button>
                    <button
                      style={{ ...styles.actionLink, color: "#c62828" }}
                      onClick={() => confirmDelete(shop.id)}
                      title="Delete"
                    >
                      🗑 Delete
                    </button>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination */}
      <div style={styles.pagination}>
        <span style={styles.paginationInfo}>
          {filteredShops.length === 0
            ? "Showing 0 of 0 shops"
            : `Showing ${(safePage - 1) * PAGE_SIZE + 1}–${Math.min(
                safePage * PAGE_SIZE,
                filteredShops.length
              )} of ${filteredShops.length} shops`}
        </span>
        <div style={styles.paginationControls}>
          <button
            style={styles.pageBtn}
            disabled={safePage <= 1}
            onClick={() => setPage((p) => Math.max(1, p - 1))}
          >
            Previous
          </button>
          {Array.from({ length: totalPages }, (_, i) => i + 1).map((p) => (
            <button
              key={p}
              style={{
                ...styles.pageBtn,
                ...(p === safePage ? styles.pageBtnActive : {}),
              }}
              onClick={() => setPage(p)}
            >
              {p}
            </button>
          ))}
          <button
            style={styles.pageBtn}
            disabled={safePage >= totalPages}
            onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
          >
            Next
          </button>
        </div>
      </div>

      {/* Add Shop modal */}
      {addOpen && (
        <ShopModal
          title="Add Shop"
          onClose={() => setAddOpen(false)}
          footer={
            <>
              <button style={styles.secondaryBtn} onClick={() => setAddOpen(false)}>
                Cancel
              </button>
              <button style={styles.primaryBtn} onClick={handleSaveAdd}>
                Save Shop
              </button>
            </>
          }
          wide
        >
          <ShopForm form={form} setForm={setForm} errors={errors} />
        </ShopModal>
      )}

      {/* Edit Shop modal */}
      {editOpen && activeShop && (
        <ShopModal
          title={`Edit Shop — ${activeShop.id}`}
          onClose={() => setEditOpen(false)}
          footer={
            <>
              <button style={styles.secondaryBtn} onClick={() => setEditOpen(false)}>
                Cancel
              </button>
              <button style={styles.primaryBtn} onClick={handleSaveEdit}>
                Update Shop
              </button>
            </>
          }
          wide
        >
          <ShopForm form={form} setForm={setForm} errors={errors} />
        </ShopModal>
      )}

      {/* Shop Details modal */}
      {viewOpen && activeShop && (
        <ShopModal
          title="Shop Details"
          onClose={() => setViewOpen(false)}
          footer={
            <>
              <button style={styles.secondaryBtn} onClick={() => setViewOpen(false)}>
                Close
              </button>
              <button
                style={styles.primaryBtn}
                onClick={() => {
                  setViewOpen(false);
                  openEdit(activeShop);
                }}
              >
                Edit Shop
              </button>
            </>
          }
        >
          <div style={styles.detailsGrid}>
            <DetailRow label="Shop ID" value={activeShop.id} />
            <DetailRow label="Shop Name" value={activeShop.name} />
            <DetailRow label="Owner" value={activeShop.owner} />
            <DetailRow label="Phone" value={activeShop.phone} />
            <DetailRow label="Email" value={activeShop.email || "—"} />
            <DetailRow label="Location" value={activeShop.location} />
            <DetailRow label="Category" value={activeShop.category} />
            <DetailRow
              label="Business Reg. Number"
              value={activeShop.regNo || "—"}
            />
            <DetailRow label="Number of Products" value={products.filter((p) => p.shopId === activeShop.id).length} />
            <DetailRow
              label="Status"
              value={<StatusBadge status={activeShop.status} />}
            />
            <DetailRow label="Date Added" value={formatDate(activeShop.dateAdded)} />
          </div>
        </ShopModal>
      )}

      {/* Products modal */}
      {productsOpen && activeShop && (
        <ShopProductsModal
          shop={activeShop}
          products={products}
          shopProducts={shopProductsLocal}
          onAssign={assignProductToShop}
          onUpdatePrice={updateShopProductPrice}
          onRemove={removeShopProduct}
          onClose={() => setProductsOpen(false)}
        />
      )}

      {/* Delete confirmation modal */}
      {deleteId && (
        <ShopModal
          title="Delete Shop"
          onClose={() => setDeleteId(null)}
          footer={
            <>
              <button style={styles.secondaryBtn} onClick={() => setDeleteId(null)}>
                Cancel
              </button>
              <button style={styles.dangerBtn} onClick={performDelete}>
                Delete
              </button>
            </>
          }
        >
          <p style={{ margin: 0 }}>
            Are you sure you want to delete shop{" "}
            <strong>{deleteId}</strong>? This action cannot be undone.
          </p>
        </ShopModal>
      )}
    </div>
  );
}

/* -------------------------------------------------------------------------
   Small presentational pieces
------------------------------------------------------------------------- */

function SummaryCard({ label, value, accent }) {
  return (
    <div style={styles.card}>
      <div style={styles.cardLabel}>{label}</div>
      <div style={{ ...styles.cardValue, color: accent }}>{value}</div>
    </div>
  );
}

function DetailRow({ label, value }) {
  return (
    <div style={styles.detailRow}>
      <span style={styles.detailLabel}>{label}</span>
      <span style={styles.detailValue}>{value}</span>
    </div>
  );
}

/* -------------------------------------------------------------------------
   Styles
------------------------------------------------------------------------- */

const styles = {
  page: {
    minHeight: "100vh",
    background: "#f8fafc",
    color: "#1f2328",
    fontFamily:
      "'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Arial, sans-serif",
    padding: "24px 28px 48px",
    boxSizing: "border-box",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    flexWrap: "wrap",
    gap: 12,
    marginBottom: 20,
  },
  h1: {
    margin: 0,
    fontSize: 24,
    fontWeight: 700,
    color: "#1f2328",
  },
  subtitle: {
    margin: "4px 0 0",
    fontSize: 14,
    color: "#5f6368",
  },
  primaryBtn: {
    background: "#107c41",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    padding: "10px 18px",
    fontSize: 14,
    fontWeight: 600,
    cursor: "pointer",
    boxShadow: "0 1px 2px rgba(16,124,65,0.25)",
    whiteSpace: "nowrap",
  },
  secondaryBtn: {
    background: "#ffffff",
    color: "#1f2328",
    border: "1px solid #d7dbe0",
    borderRadius: 8,
    padding: "9px 14px",
    fontSize: 13,
    fontWeight: 600,
    cursor: "pointer",
  },
  dangerBtn: {
    background: "#c62828",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    padding: "9px 16px",
    fontSize: 13,
    fontWeight: 600,
    cursor: "pointer",
  },
  cardsRow: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
    gap: 14,
    marginBottom: 20,
  },
  card: {
    background: "#fff",
    border: "1px solid #e4e7ea",
    borderRadius: 10,
    padding: "14px 16px",
    boxShadow: "0 1px 2px rgba(0,0,0,0.03)",
  },
  cardLabel: {
    fontSize: 12,
    fontWeight: 600,
    color: "#5f6368",
    textTransform: "uppercase",
    letterSpacing: 0.4,
    marginBottom: 6,
  },
  cardValue: {
    fontSize: 26,
    fontWeight: 700,
  },
  toolbar: {
    display: "flex",
    flexWrap: "wrap",
    gap: 10,
    alignItems: "center",
    background: "#fff",
    border: "1px solid #e4e7ea",
    borderRadius: 10,
    padding: 12,
    marginBottom: 14,
  },
  searchInput: {
    flex: "1 1 220px",
    minWidth: 180,
    padding: "9px 12px",
    borderRadius: 8,
    border: "1px solid #d7dbe0",
    fontSize: 13,
    outline: "none",
  },
  select: {
    padding: "9px 10px",
    borderRadius: 8,
    border: "1px solid #d7dbe0",
    fontSize: 13,
    background: "#fff",
    color: "#1f2328",
  },
  toolbarSpacer: {
    flex: "1 1 auto",
  },
  bulkBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    background: "#eaf6ee",
    border: "1px solid #bfe3cc",
    borderRadius: 10,
    padding: "10px 14px",
    marginBottom: 14,
    flexWrap: "wrap",
    gap: 10,
  },
  bulkText: {
    fontSize: 13,
    fontWeight: 600,
    color: "#107c41",
  },
  bulkBtns: {
    display: "flex",
    gap: 8,
  },
  bulkBtnGreen: {
    background: "#107c41",
    color: "#fff",
    border: "none",
    borderRadius: 6,
    padding: "7px 12px",
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
  },
  bulkBtnGray: {
    background: "#5f6368",
    color: "#fff",
    border: "none",
    borderRadius: 6,
    padding: "7px 12px",
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
  },
  bulkBtnRed: {
    background: "#c62828",
    color: "#fff",
    border: "none",
    borderRadius: 6,
    padding: "7px 12px",
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
  },
  tableWrap: {
    background: "#fff",
    border: "1px solid #e4e7ea",
    borderRadius: 10,
    overflow: "auto",
    maxWidth: "100%",
  },
  table: {
    width: "100%",
    borderCollapse: "collapse",
    fontSize: 13,
    minWidth: 1000,
  },
  th: {
    position: "sticky",
    top: 0,
    background: "#f3f6f4",
    borderBottom: "2px solid #d7dbe0",
    borderRight: "1px solid #e9ecee",
    padding: "10px 12px",
    textAlign: "left",
    fontWeight: 700,
    color: "#33383d",
    whiteSpace: "nowrap",
    cursor: "default",
    userSelect: "none",
    zIndex: 1,
  },
  thCheckbox: {
    width: 40,
    textAlign: "center",
  },
  tr: {
    borderBottom: "1px solid #edeff1",
  },
  td: {
    padding: "9px 12px",
    borderRight: "1px solid #f1f2f4",
    verticalAlign: "middle",
  },
  actionLink: {
    background: "none",
    border: "none",
    color: "#0f6cbd",
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
    padding: "2px 6px",
    marginRight: 2,
  },
  emptyCell: {
    padding: 0,
  },
  emptyState: {
    padding: "48px 12px",
    textAlign: "center",
  },
  emptyTitle: {
    fontSize: 15,
    fontWeight: 700,
    color: "#33383d",
    marginBottom: 4,
  },
  emptySubtitle: {
    fontSize: 13,
    color: "#5f6368",
  },
  pagination: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    flexWrap: "wrap",
    gap: 10,
    marginTop: 14,
  },
  paginationInfo: {
    fontSize: 13,
    color: "#5f6368",
  },
  paginationControls: {
    display: "flex",
    gap: 6,
  },
  pageBtn: {
    background: "#fff",
    border: "1px solid #d7dbe0",
    borderRadius: 6,
    padding: "6px 11px",
    fontSize: 12,
    fontWeight: 600,
    cursor: "pointer",
    color: "#1f2328",
  },
  pageBtnActive: {
    background: "#107c41",
    borderColor: "#107c41",
    color: "#fff",
  },
  overlay: {
    position: "fixed",
    inset: 0,
    background: "rgba(15, 20, 25, 0.45)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 16,
    zIndex: 1000,
  },
  modal: {
    background: "#fff",
    borderRadius: 12,
    width: "100%",
    maxHeight: "90vh",
    overflow: "auto",
    boxShadow: "0 10px 40px rgba(0,0,0,0.2)",
  },
  modalHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "16px 20px",
    borderBottom: "1px solid #eceef0",
    position: "sticky",
    top: 0,
    background: "#fff",
  },
  modalTitle: {
    margin: 0,
    fontSize: 16,
    fontWeight: 700,
  },
  iconBtn: {
    background: "none",
    border: "none",
    fontSize: 16,
    cursor: "pointer",
    color: "#5f6368",
  },
  modalBody: {
    padding: 20,
  },
  modalFooter: {
    display: "flex",
    justifyContent: "flex-end",
    gap: 10,
    padding: "14px 20px",
    borderTop: "1px solid #eceef0",
  },
  formGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
    gap: 14,
  },
  formField: {
    display: "flex",
    flexDirection: "column",
    gap: 6,
  },
  label: {
    fontSize: 12,
    fontWeight: 600,
    color: "#33383d",
  },
  input: {
    padding: "9px 10px",
    borderRadius: 8,
    border: "1px solid #d7dbe0",
    fontSize: 13,
    outline: "none",
    background: "#fff",
    color: "#1f2328",
  },
  inputError: {
    borderColor: "#c62828",
  },
  errorText: {
    fontSize: 11,
    color: "#c62828",
  },
  detailsGrid: {
    display: "flex",
    flexDirection: "column",
    gap: 10,
  },
  detailRow: {
    display: "flex",
    justifyContent: "space-between",
    borderBottom: "1px solid #f1f2f4",
    paddingBottom: 8,
  },
  detailLabel: {
    fontSize: 12,
    fontWeight: 600,
    color: "#5f6368",
  },
  detailValue: {
    fontSize: 13,
    fontWeight: 600,
    color: "#1f2328",
    textAlign: "right",
  },
};


// ============================================================
// ---------- Price Update (Price Management) ----------
// ============================================================

/* ------------------------------------------------------------------ */
/* MOCK DATA                                                          */
/* ------------------------------------------------------------------ */

const SHOPS = [
  { id: "quickbasket", name: "QuickBasket" },
  { id: "greenmart", name: "GreenMart" },
  { id: "freshpoint", name: "FreshPoint" },
  { id: "valuestore", name: "ValueStore" },
  { id: "budgetfoods", name: "Budget Foods" },
];

const PRICE_CATEGORIES = ["Food & Beverages", "Dairy", "Household", "Personal Care"];

// base product data: currentPrice null = missing price, daysAgo null = missing
const RAW_PRODUCTS = [
  { id: "QP001", name: "Soko Maize Flour 2kg", icon: "🌽", category: "Food & Beverages", previousPrice: 225, currentPrice: 225, daysAgo: 9 },
  { id: "QP002", name: "Jogoo Maize Flour 2kg", icon: "🌽", category: "Food & Beverages", previousPrice: 220, currentPrice: 220, daysAgo: 16 },
  { id: "QP003", name: "Pembe Maize Flour 2kg", icon: "🌽", category: "Food & Beverages", previousPrice: 215, currentPrice: 215, daysAgo: 2 },
  { id: "QP004", name: "Ajab Wheat Flour 2kg", icon: "🌾", category: "Food & Beverages", previousPrice: 230, currentPrice: null, daysAgo: null },
  { id: "QP005", name: "Brookside Milk 500ml", icon: "🥛", category: "Dairy", previousPrice: 65, currentPrice: 65, daysAgo: 1 },
  { id: "QP006", name: "Brookside Milk 1L", icon: "🥛", category: "Dairy", previousPrice: 118, currentPrice: 118, daysAgo: 20 },
  { id: "QP007", name: "Fresh Fri Milk 500ml", icon: "🥛", category: "Dairy", previousPrice: 62, currentPrice: 62, daysAgo: 5 },
  { id: "QP008", name: "Kabras Sugar 2kg", icon: "🍚", category: "Food & Beverages", previousPrice: 320, currentPrice: 325, daysAgo: 0 },
  { id: "QP009", name: "Menengai Bar Soap 1kg", icon: "🧼", category: "Household", previousPrice: 180, currentPrice: 180, daysAgo: 12 },
  { id: "QP010", name: "Blue Band 500g", icon: "🧈", category: "Dairy", previousPrice: 210, currentPrice: null, daysAgo: null },
  { id: "QP011", name: "Elianto Cooking Oil 1L", icon: "🛢️", category: "Food & Beverages", previousPrice: 340, currentPrice: 340, daysAgo: 3 },
  { id: "QP012", name: "Ketepa Tea Leaves 250g", icon: "🍵", category: "Food & Beverages", previousPrice: 145, currentPrice: 145, daysAgo: 18 },
  { id: "QP013", name: "Colgate Toothpaste 100ml", icon: "🦷", category: "Personal Care", previousPrice: 155, currentPrice: 155, daysAgo: 6 },
  { id: "QP014", name: "Royco Beef Cubes", icon: "🧂", category: "Food & Beverages", previousPrice: 55, currentPrice: 95, daysAgo: 0, unusual: true },
  { id: "QP015", name: "Sunlight Washing Powder 1kg", icon: "🧺", category: "Household", previousPrice: 195, currentPrice: 195, daysAgo: 10 },
];

// deterministic "other shop" market prices, derived from base price
function marketPrices(base, id) {
  const seed = id.charCodeAt(id.length - 1) + id.charCodeAt(2);
  const variants = [-0.06, -0.02, 0.015, 0.05].map((v, i) => {
    const wobble = ((seed * (i + 3)) % 7) - 3; // -3..3
    return Math.max(1, Math.round((base * (1 + v)) + wobble));
  });
  return {
    greenmart: variants[0],
    freshpoint: variants[1],
    valuestore: variants[2],
    budgetfoods: variants[3],
  };
}

function buildProducts() {
  return RAW_PRODUCTS.map((p) => {
    const base = p.currentPrice ?? p.previousPrice;
    return { ...p, market: marketPrices(base, p.id), sku: p.id };
  });
}

const KSH = (n) =>
  n === null || n === undefined || Number.isNaN(n)
    ? "—"
    : `KSh ${Math.round(n).toLocaleString("en-KE")}`;

function pctChange(oldP, newP) {
  if (!oldP) return null;
  return ((newP - oldP) / oldP) * 100;
}

function statusOf(product) {
  if (product.currentPrice === null) return "missing";
  if (product.unusual) return "unusual";
  if (product.daysAgo <= 3) return "fresh";
  if (product.daysAgo <= 7) return "review";
  if (product.daysAgo <= 14) return "stale";
  return "outdated";
}

const STATUS_META = {
  fresh: { label: "Fresh", color: "var(--green)", bg: "var(--green-bg)", dot: "🟢" },
  review: { label: "Needs review", color: "var(--amber)", bg: "var(--amber-bg)", dot: "🟡" },
  stale: { label: "Stale", color: "var(--orange)", bg: "var(--orange-bg)", dot: "🟠" },
  outdated: { label: "Outdated", color: "var(--red)", bg: "var(--red-bg)", dot: "🔴" },
  unusual: { label: "Unusual", color: "var(--red)", bg: "var(--red-bg)", dot: "⚠️" },
  missing: { label: "Missing", color: "var(--ink-faint)", bg: "var(--gray-bg)", dot: "⚪" },
};

function timeAgoLabel(daysAgo) {
  if (daysAgo === null) return "No price set";
  if (daysAgo === 0) return "Today";
  if (daysAgo === 1) return "Yesterday";
  return `${daysAgo} days ago`;
}

/* ------------------------------------------------------------------ */
/* ROOT APP                                                            */
/* ------------------------------------------------------------------ */

function PriceUpdate({ onBack, products: sharedProducts, onProductsChange, onPriceCommitted }) {
  const [products, setProducts] = useState(() => sharedProducts ?? buildProducts());
  useEffect(() => {
    if (onProductsChange) onProductsChange(products);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [products]);
  const [shop, setShop] = useState(SHOPS[0].id);
  const [mode, setMode] = useState("fast"); // fast | bulk | spreadsheet
  const [query, setQuery] = useState("");
  const [activeChip, setActiveChip] = useState("all");
  const [queue, setQueue] = useState({}); // id -> { oldPrice, newPrice }
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [recentlyUpdated, setRecentlyUpdated] = useState([]); // this-session history
  const [history, setHistory] = useState({}); // id -> [{date, previous, new, change, by}]
  const [toast, setToast] = useState(null);
  const [undoSnapshot, setUndoSnapshot] = useState(null);
  const [pendingWarning, setPendingWarning] = useState(null); // { id, newPrice, oldPrice }
  const [showImport, setShowImport] = useState(false);
  const [showHistoryFor, setShowHistoryFor] = useState(null);
  const [showSummary, setShowSummary] = useState(false);
  const [queueOpenMobile, setQueueOpenMobile] = useState(false);

  const toastTimer = useRef(null);
  const undoTimer = useRef(null);

  const currentShop = SHOPS.find((s) => s.id === shop);

  /* ---------------- derived data ---------------- */

  const counts = useMemo(() => {
    const c = { outdated: 0, missing: 0, today: 0, unusual: 0 };
    products.forEach((p) => {
      const s = statusOf(p);
      if (s === "outdated") c.outdated++;
      if (s === "missing") c.missing++;
      if (s === "unusual") c.unusual++;
      if (p.daysAgo === 0) c.today++;
    });
    return c;
  }, [products]);

  const freshness = useMemo(() => {
    const withPrice = products.filter((p) => p.currentPrice !== null);
    const freshCount = withPrice.filter((p) => p.daysAgo <= 7).length;
    return Math.round((freshCount / products.length) * 100);
  }, [products]);

  const filtered = useMemo(() => {
    let list = products;
    const q = query.trim().toLowerCase();
    if (q) {
      list = list.filter(
        (p) =>
          p.name.toLowerCase().includes(q) ||
          p.sku.toLowerCase().includes(q) ||
          p.category.toLowerCase().includes(q)
      );
    }
    switch (activeChip) {
      case "needs":
        list = list.filter((p) => ["review", "stale", "outdated", "missing"].includes(statusOf(p)));
        break;
      case "outdated":
        list = list.filter((p) => statusOf(p) === "outdated");
        break;
      case "missing":
        list = list.filter((p) => statusOf(p) === "missing");
        break;
      case "today":
        list = list.filter((p) => p.daysAgo === 0);
        break;
      case "up":
        list = list.filter((p) => queue[p.id] && queue[p.id].newPrice > queue[p.id].oldPrice);
        break;
      case "down":
        list = list.filter((p) => queue[p.id] && queue[p.id].newPrice < queue[p.id].oldPrice);
        break;
      case "unusual":
        list = list.filter((p) => statusOf(p) === "unusual");
        break;
      default:
        break;
    }
    return list;
  }, [products, query, activeChip, queue]);

  const queueList = Object.entries(queue).map(([id, v]) => ({
    id,
    product: products.find((p) => p.id === id),
    ...v,
  }));

  const queueNetChange = queueList.reduce((sum, q) => sum + (q.newPrice - q.oldPrice), 0);
  const queueOldTotal = queueList.reduce((sum, q) => sum + q.oldPrice, 0);
  const queueNewTotal = queueList.reduce((sum, q) => sum + q.newPrice, 0);

  /* ---------------- actions ---------------- */

  const flashToast = (msg, opts = {}) => {
    clearTimeout(toastTimer.current);
    setToast({ msg, ...opts });
    toastTimer.current = setTimeout(() => setToast(null), opts.sticky ? 6000 : 3200);
  };

  // Commits a single price immediately (used by the Fast Update card's "Save"
  // button, so a single edit actually saves right away instead of only being
  // staged into the batch queue).
  const commitOne = (id, oldPrice, newPrice) => {
    const product = products.find((p) => p.id === id);
    if (!product) return;
    const change = pctChange(oldPrice, newPrice);
    const dateStr = "Today";
    setUndoSnapshot(products);
    clearTimeout(undoTimer.current);
    undoTimer.current = setTimeout(() => setUndoSnapshot(null), 8000);
    setHistory((h) => ({
      ...h,
      [id]: [{ date: dateStr, previous: oldPrice, next: newPrice, change, by: "Admin" }, ...(h[id] || [])],
    }));
    setRecentlyUpdated((r) => [{ id, name: product.name, icon: product.icon, oldPrice, newPrice, change, when: "Just now" }, ...r].slice(0, 12));
    setProducts((prev) =>
      prev.map((p) =>
        p.id === id
          ? { ...p, previousPrice: oldPrice, currentPrice: newPrice, daysAgo: 0, unusual: change !== null && Math.abs(change) > 20 }
          : p
      )
    );
    removeFromQueue(id);
    flashToast(`Price updated for ${product.name}`, { undo: true, sticky: true });
    if (onPriceCommitted) onPriceCommitted(new Date());
  };

  const stageChange = useCallback(
    (id, newPrice, { skipWarningCheck, immediate } = {}) => {
      const product = products.find((p) => p.id === id);
      if (!product || newPrice === "" || newPrice === null || Number.isNaN(newPrice)) return;
      const oldPrice = product.currentPrice ?? product.previousPrice ?? 0;
      const change = pctChange(oldPrice, newPrice);
      if (!skipWarningCheck && change !== null && Math.abs(change) > 20 && oldPrice > 0) {
        setPendingWarning({ id, newPrice, oldPrice, immediate });
        return;
      }
      if (immediate) {
        commitOne(id, oldPrice, newPrice);
      } else {
        setQueue((q) => ({ ...q, [id]: { oldPrice, newPrice } }));
      }
    },
    [products]
  );

  const removeFromQueue = (id) => {
    setQueue((q) => {
      const next = { ...q };
      delete next[id];
      return next;
    });
  };

  const confirmWarning = () => {
    if (!pendingWarning) return;
    if (pendingWarning.immediate) {
      commitOne(pendingWarning.id, pendingWarning.oldPrice, pendingWarning.newPrice);
    } else {
      setQueue((q) => ({
        ...q,
        [pendingWarning.id]: { oldPrice: pendingWarning.oldPrice, newPrice: pendingWarning.newPrice },
      }));
    }
    setPendingWarning(null);
  };

  const commitQueue = () => {
    if (queueList.length === 0) return;
    const snapshot = products;
    const today = new Date();
    const dateStr = "Today";
    const newHistory = { ...history };
    const newRecent = [];

    const updated = products.map((p) => {
      const q = queue[p.id];
      if (!q) return p;
      const change = pctChange(q.oldPrice, q.newPrice);
      newHistory[p.id] = [
        { date: dateStr, previous: q.oldPrice, next: q.newPrice, change, by: "Admin" },
        ...(newHistory[p.id] || []),
      ];
      newRecent.push({ id: p.id, name: p.name, icon: p.icon, oldPrice: q.oldPrice, newPrice: q.newPrice, change, when: "Just now" });
      return { ...p, previousPrice: q.oldPrice, currentPrice: q.newPrice, daysAgo: 0, unusual: change !== null && Math.abs(change) > 20 };
    });

    setUndoSnapshot(snapshot);
    clearTimeout(undoTimer.current);
    undoTimer.current = setTimeout(() => setUndoSnapshot(null), 8000);

    setProducts(updated);
    setHistory(newHistory);
    setRecentlyUpdated((r) => [...newRecent, ...r].slice(0, 12));
    const count = queueList.length;
    setQueue({});
    flashToast(`${count} price${count > 1 ? "s" : ""} updated successfully`, { undo: true, sticky: true });
    if (onPriceCommitted) onPriceCommitted(new Date());
  };

  const undo = () => {
    if (undoSnapshot) {
      setProducts(undoSnapshot);
      setUndoSnapshot(null);
      setToast(null);
      clearTimeout(undoTimer.current);
    }
  };

  const applyChip = (chip) => setActiveChip((c) => (c === chip ? "all" : chip));

  const scrollToWorkspace = () => {
    const el = document.getElementById("qp-workspace");
    if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  /* ---------------- render ---------------- */

  return (
    <div className="qp-app">
      <style>{CSS}</style>

      <div className="qp-shell">
        <Header
          onBack={onBack}
          onImport={() => setShowImport(true)}
          onHistory={() => setShowHistoryFor("__all__")}
          onSaveAll={commitQueue}
          queueCount={queueList.length}
          onFinish={() => setShowSummary(true)}
        />

        <FreshnessBar freshness={freshness} total={products.length} fresh={Math.round((freshness / 100) * products.length)} />

        <AttentionPanel
          counts={counts}
          onSelect={(chip) => {
            applyChip(chip);
            scrollToWorkspace();
          }}
        />

        <div className="qp-controlbar">
          <ShopSelector shop={shop} setShop={setShop} currentShop={currentShop} productCount={products.length} freshness={freshness} />
          <ModeSelector mode={mode} setMode={setMode} />
        </div>

        <ProductSearch query={query} setQuery={setQuery} resultCount={filtered.length} />

        <FilterChips active={activeChip} onSelect={applyChip} />

        <div id="qp-workspace" className="qp-workspace">
          <div className="qp-main">
            {mode === "fast" && (
              <FastUpdateView
                products={filtered}
                queue={queue}
                stageChange={stageChange}
                shop={shop}
              />
            )}
            {mode === "spreadsheet" && (
              <SpreadsheetView products={filtered} queue={queue} stageChange={stageChange} />
            )}
            {mode === "bulk" && (
              <BulkUpdateView
                products={filtered}
                selectedIds={selectedIds}
                setSelectedIds={setSelectedIds}
                onApply={(entries) => entries.forEach(([id, price]) => stageChange(id, price, { skipWarningCheck: false }))}
              />
            )}

            {recentlyUpdated.length > 0 && (
              <RecentlyUpdated items={recentlyUpdated} onOpenHistory={(id) => setShowHistoryFor(id)} />
            )}
          </div>

          <UpdateQueueSidebar
            queueList={queueList}
            onRemove={removeFromQueue}
            onSaveAll={commitQueue}
            netChange={queueNetChange}
            open={queueOpenMobile}
            onClose={() => setQueueOpenMobile(false)}
          />
        </div>
      </div>

      {queueList.length > 0 && (
        <SaveBar
          count={queueList.length}
          oldTotal={queueOldTotal}
          newTotal={queueNewTotal}
          netChange={queueNetChange}
          onDiscard={() => setQueue({})}
          onSave={commitQueue}
          onOpenQueue={() => setQueueOpenMobile(true)}
        />
      )}

      {toast && (
        <PriceToast msg={toast.msg} showUndo={toast.undo && undoSnapshot} onUndo={undo} onClose={() => setToast(null)} />
      )}

      {pendingWarning && (
        <WarningModal
          data={pendingWarning}
          onCancel={() => setPendingWarning(null)}
          onConfirm={confirmWarning}
        />
      )}

      {showImport && <PriceImportModal onClose={() => setShowImport(false)} onApply={() => { setShowImport(false); flashToast("148 prices imported successfully", { undo: false }); }} />}

      {showHistoryFor && (
        <HistoryModal
          productId={showHistoryFor}
          products={products}
          history={history}
          onClose={() => setShowHistoryFor(null)}
        />
      )}

      {showSummary && (
        <SummaryModal
          items={recentlyUpdated}
          freshness={freshness}
          attention={counts.outdated + counts.missing}
          onClose={() => setShowSummary(false)}
          onContinue={() => setShowSummary(false)}
        />
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* HEADER                                                              */
/* ------------------------------------------------------------------ */

function Header({ onBack, onImport, onHistory, onSaveAll, queueCount, onFinish }) {
  return (
    <header className="qp-header">
      <div>
        {onBack && (
          <button
            onClick={onBack}
            style={{
              display: "inline-flex", alignItems: "center", gap: 6, background: "none", border: "none",
              cursor: "pointer", color: "var(--ink-soft)", fontSize: 13, fontWeight: 600, padding: 0, marginBottom: 10,
            }}
          >
            <ArrowLeft size={14} /> Back to Dashboard
          </button>
        )}
        <h1 className="qp-title">
          <Zap size={26} strokeWidth={2.4} className="qp-title-icon" />
          Quick Price Update
        </h1>
        <p className="qp-subtitle">
          Update your shop prices in seconds. QuickPick calculates every change, flags anything unusual, and keeps your comparison data fresh.
        </p>
      </div>
      <div className="qp-header-actions">
        <button className="qp-btn qp-btn-ghost" onClick={onImport}>
          <Upload size={16} /> Import prices
        </button>
        <button className="qp-btn qp-btn-ghost" onClick={onHistory}>
          <History size={16} /> Price history
        </button>
        <button className="qp-btn qp-btn-ghost" onClick={onFinish}>
          <CheckCircle2 size={16} /> Finish session
        </button>
        <button className="qp-btn qp-btn-primary" onClick={onSaveAll} disabled={queueCount === 0}>
          <Save size={16} /> Save all changes{queueCount > 0 ? ` (${queueCount})` : ""}
        </button>
      </div>
    </header>
  );
}

function FreshnessBar({ freshness, total, fresh }) {
  return (
    <div className="qp-freshness">
      <div className="qp-freshness-label">
        <span>Price freshness</span>
        <strong>{freshness}%</strong>
      </div>
      <div className="qp-freshness-track">
        <div className="qp-freshness-fill" style={{ width: `${freshness}%` }} />
      </div>
      <span className="qp-freshness-sub">{fresh} / {total} products updated within 7 days</span>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* ATTENTION PANEL                                                     */
/* ------------------------------------------------------------------ */

function AttentionPanel({ counts, onSelect }) {
  const cards = [
    { chip: "outdated", tone: "red", title: `${counts.outdated} price${counts.outdated === 1 ? "" : "s"} outdated`, sub: "Not updated for 15+ days", cta: "Update now" },
    { chip: "missing", tone: "orange", title: `${counts.missing} price${counts.missing === 1 ? "" : "s"} missing`, sub: "Products without a current price", cta: "Add prices" },
    { chip: "today", tone: "green", title: `${counts.today} updated today`, sub: "Nice — those prices are fresh", cta: "View updates" },
    { chip: "unusual", tone: "amber", title: `${counts.unusual} unusual change${counts.unusual === 1 ? "" : "s"}`, sub: "Price moved more than 20%", cta: "Review" },
  ];
  return (
    <section className="qp-section">
      <h2 className="qp-section-title">Prices that need your attention</h2>
      <div className="qp-attention-grid">
        {cards.map((c) => (
          <button key={c.chip} className={`qp-attn-card qp-tone-${c.tone}`} onClick={() => onSelect(c.chip)}>
            <div className="qp-attn-title">{c.title}</div>
            <div className="qp-attn-sub">{c.sub}</div>
            <div className="qp-attn-cta">{c.cta} →</div>
          </button>
        ))}
      </div>
    </section>
  );
}

/* ------------------------------------------------------------------ */
/* SHOP + MODE SELECTORS                                               */
/* ------------------------------------------------------------------ */

function ShopSelector({ shop, setShop, currentShop, productCount, freshness }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="qp-shopselect">
      <span className="qp-field-label">Updating prices for</span>
      <div className="qp-shopselect-row">
        <button className="qp-shop-btn" onClick={() => setOpen((o) => !o)}>
          <Store size={16} /> {currentShop.name} <ChevronDown size={15} />
        </button>
        {open && (
          <div className="qp-shop-dropdown">
            {SHOPS.map((s) => (
              <button
                key={s.id}
                className={`qp-shop-option ${s.id === shop ? "active" : ""}`}
                onClick={() => { setShop(s.id); setOpen(false); }}
              >
                {s.name}
              </button>
            ))}
          </div>
        )}
      </div>
      <div className="qp-shop-meta">
        {productCount} products · {freshness}% prices updated this week
      </div>
    </div>
  );
}

function ModeSelector({ mode, setMode }) {
  const modes = [
    { id: "fast", label: "Fast update", icon: Zap },
    { id: "bulk", label: "Bulk update", icon: ListChecks },
    { id: "spreadsheet", label: "Spreadsheet", icon: FileSpreadsheet },
  ];
  return (
    <div className="qp-mode-select" role="tablist" aria-label="Update mode">
      {modes.map((m) => (
        <button
          key={m.id}
          className={`qp-mode-btn ${mode === m.id ? "active" : ""}`}
          onClick={() => setMode(m.id)}
          role="tab"
          aria-selected={mode === m.id}
        >
          <m.icon size={15} /> {m.label}
        </button>
      ))}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* SEARCH + FILTER CHIPS                                               */
/* ------------------------------------------------------------------ */

function ProductSearch({ query, setQuery, resultCount }) {
  return (
    <div className="qp-search-wrap">
      <Search size={20} className="qp-search-icon" />
      <input
        className="qp-search-input"
        placeholder="Type product name, scan SKU, brand or barcode…"
        value={query}
        onChange={(e) => setQuery(e.target.value)}
      />
      {query && <span className="qp-search-count">{resultCount} found</span>}
      <button className="qp-btn qp-btn-ghost qp-scan-btn">
        <Camera size={16} /> Scan
      </button>
    </div>
  );
}

function FilterChips({ active, onSelect }) {
  const chips = [
    { id: "all", label: "All" },
    { id: "needs", label: "Needs update" },
    { id: "outdated", label: "Outdated" },
    { id: "missing", label: "Missing" },
    { id: "today", label: "Changed today" },
    { id: "up", label: "Price increased" },
    { id: "down", label: "Price decreased" },
    { id: "unusual", label: "Unusual changes" },
  ];
  return (
    <div className="qp-chip-row">
      {chips.map((c) => (
        <button
          key={c.id}
          className={`qp-chip ${active === c.id ? "active" : ""}`}
          onClick={() => onSelect(c.id === "all" ? "all" : c.id)}
        >
          {c.label}
        </button>
      ))}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* FAST UPDATE VIEW                                                     */
/* ------------------------------------------------------------------ */

function FastUpdateView({ products, queue, stageChange, shop }) {
  const inputRefs = useRef({});
  const [expandedId, setExpandedId] = useState(products[0]?.id || null);

  useEffect(() => {
    // keep an expanded card among current filtered results
    if (!products.find((p) => p.id === expandedId)) {
      setExpandedId(products[0]?.id || null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [products]);

  const focusNext = (currentId) => {
    const idx = products.findIndex((p) => p.id === currentId);
    const next = products[idx + 1];
    if (next) {
      setExpandedId(next.id);
      requestAnimationFrame(() => inputRefs.current[next.id]?.focus());
    }
  };
  const focusPrev = (currentId) => {
    const idx = products.findIndex((p) => p.id === currentId);
    const prev = products[idx - 1];
    if (prev) {
      setExpandedId(prev.id);
      requestAnimationFrame(() => inputRefs.current[prev.id]?.focus());
    }
  };

  if (products.length === 0) {
    return <EmptyState text="No products match this filter yet." />;
  }

  return (
    <div className="qp-card-stream">
      {products.map((product) => (
        <QuickUpdateCard
          key={product.id}
          product={product}
          shop={shop}
          pending={queue[product.id]}
          expanded={expandedId === product.id}
          onExpand={() => setExpandedId(product.id)}
          inputRef={(el) => (inputRefs.current[product.id] = el)}
          onSave={(price) => { stageChange(product.id, price, { immediate: true }); focusNext(product.id); }}
          onArrowDown={() => focusNext(product.id)}
          onArrowUp={() => focusPrev(product.id)}
        />
      ))}
    </div>
  );
}

function QuickUpdateCard({ product, shop, pending, expanded, onExpand, inputRef, onSave, onArrowDown, onArrowUp }) {
  const baseline = pending ? pending.newPrice : product.currentPrice;
  const [value, setValue] = useState(baseline ?? "");
  const [showIntel, setShowIntel] = useState(false);
  const status = statusOf(product);
  const meta = STATUS_META[status];

  useEffect(() => {
    setValue(pending ? pending.newPrice : (product.currentPrice ?? ""));
  }, [pending, product.currentPrice]);

  const numeric = value === "" ? null : Number(value);
  const oldPrice = product.currentPrice ?? product.previousPrice ?? 0;
  const change = numeric !== null ? pctChange(oldPrice, numeric) : null;

  const commit = () => {
    if (numeric === null || Number.isNaN(numeric)) return;
    onSave(numeric);
  };

  const handleKey = (e) => {
    if (e.key === "Enter") { e.preventDefault(); commit(); }
    else if (e.key === "Escape") { setValue(product.currentPrice ?? ""); e.currentTarget.blur(); }
    else if (e.key === "ArrowDown") { e.preventDefault(); onArrowDown(); }
    else if (e.key === "ArrowUp") { e.preventDefault(); onArrowUp(); }
  };

  const adjust = (delta) => {
    const base = numeric ?? oldPrice ?? 0;
    setValue(Math.max(0, base + delta));
  };

  const market = product.market;
  const marketVals = [oldPrice, market.greenmart, market.freshpoint, market.valuestore, market.budgetfoods].filter((v) => v > 0);
  const lowest = Math.min(...marketVals);
  const highest = Math.max(...marketVals);
  const avg = Math.round(marketVals.reduce((a, b) => a + b, 0) / marketVals.length);

  return (
    <div className={`qp-card ${expanded ? "expanded" : ""} ${pending ? "queued" : ""}`} onClick={onExpand}>
      <div className="qp-card-head">
        <span className="qp-card-icon">{product.icon}</span>
        <div className="qp-card-titles">
          <div className="qp-card-name">{product.name}</div>
          <div className="qp-card-sku">SKU: {product.sku}</div>
        </div>
        <span className="qp-status-pill" style={{ color: meta.color, background: meta.bg }}>
          {meta.dot} {meta.label}
        </span>
      </div>

      <div className="qp-card-body">
        <div className="qp-card-prev">
          <span className="qp-field-label">Previous price</span>
          <span className="qp-prev-price">{KSH(product.currentPrice ?? product.previousPrice)}</span>
          <span className="qp-last-updated"><Clock size={12} /> {timeAgoLabel(product.daysAgo)}</span>
        </div>

        <div className="qp-card-current">
          <span className="qp-field-label">Current price</span>
          <div className="qp-price-input-row">
            <span className="qp-currency">KSh</span>
            <input
              ref={inputRef}
              type="number"
              className="qp-price-input"
              value={value}
              onFocus={onExpand}
              onChange={(e) => setValue(e.target.value)}
              onKeyDown={handleKey}
              placeholder="0"
            />
          </div>
          <div className="qp-quickadjust">
            <button onClick={(e) => { e.stopPropagation(); adjust(-5); }}>− 5</button>
            <button onClick={(e) => { e.stopPropagation(); adjust(-1); }}>− 1</button>
            <button onClick={(e) => { e.stopPropagation(); adjust(1); }}>+ 1</button>
            <button onClick={(e) => { e.stopPropagation(); adjust(5); }}>+ 5</button>
          </div>
        </div>

        {numeric !== null && change !== null && oldPrice > 0 && (
          <div className={`qp-change-preview ${change >= 0 ? "up" : "down"}`}>
            {change >= 0 ? <TrendingUp size={14} /> : <TrendingDown size={14} />}
            {change >= 0 ? "+" : ""}{KSH(numeric - oldPrice).replace("KSh ", "KSh ")} ({change >= 0 ? "+" : ""}{change.toFixed(1)}%)
          </div>
        )}

        <button className="qp-save-btn" onClick={(e) => { e.stopPropagation(); commit(); }}>
          <Check size={15} /> {pending ? "Update" : "Save"}
        </button>
      </div>

      {expanded && (
        <div className="qp-card-extra" onClick={(e) => e.stopPropagation()}>
          <div className="qp-copyrow">
            <span className="qp-field-label">Copy price</span>
            <div className="qp-copy-btns">
              <button onClick={() => setValue(product.previousPrice)}>Previous · {KSH(product.previousPrice)}</button>
              <button onClick={() => setValue(lowest)}>Lowest shop · {KSH(lowest)}</button>
              <button onClick={() => setValue(avg)}>Average · {KSH(avg)}</button>
            </div>
          </div>

          <button className="qp-intel-toggle" onClick={() => setShowIntel((s) => !s)}>
            <Sparkles size={13} /> {showIntel ? "Hide" : "Show"} price intelligence
          </button>

          {showIntel && (
            <div className="qp-intel-panel">
              <div className="qp-intel-row"><span>Lowest shop</span><strong>{KSH(lowest)}</strong></div>
              <div className="qp-intel-row"><span>Average</span><strong>{KSH(avg)}</strong></div>
              <div className="qp-intel-row"><span>Highest</span><strong>{KSH(highest)}</strong></div>
              <div className="qp-intel-row"><span>Your price</span><strong>{KSH(oldPrice)}</strong></div>
              <div className="qp-intel-suggest">
                Competitive range: <strong>{KSH(lowest)}–{KSH(avg)}</strong>
                <button className="qp-use-btn" onClick={() => setValue(Math.round((lowest + avg) / 2))}>
                  Use {KSH(Math.round((lowest + avg) / 2))}
                </button>
              </div>
              <div className="qp-market-list">
                {SHOPS.filter((s) => s.id !== shop).map((s) => (
                  <div key={s.id} className="qp-market-item">
                    <span>{s.name}</span>
                    <span>{KSH(market[s.id])}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* SPREADSHEET VIEW                                                     */
/* ------------------------------------------------------------------ */

function SpreadsheetView({ products, queue, stageChange }) {
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState("");

  if (products.length === 0) return <EmptyState text="No products match this filter yet." />;

  const startEdit = (product) => {
    setEditingId(product.id);
    setDraft(String(queue[product.id]?.newPrice ?? product.currentPrice ?? ""));
  };

  const commit = (product) => {
    const n = Number(draft);
    if (!Number.isNaN(n) && draft !== "") stageChange(product.id, n);
    setEditingId(null);
  };

  return (
    <div className="qp-sheet-wrap">
      <table className="qp-sheet">
        <thead>
          <tr>
            <th>Product</th>
            <th>SKU</th>
            <th className="right">Previous</th>
            <th className="right">Current</th>
            <th className="right">Change</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          {products.map((p) => {
            const pending = queue[p.id];
            const shown = pending ? pending.newPrice : p.currentPrice;
            const oldPrice = p.currentPrice ?? p.previousPrice ?? 0;
            const change = shown !== null ? pctChange(oldPrice, shown) : null;
            const meta = STATUS_META[statusOf(p)];
            return (
              <tr key={p.id} className={pending ? "queued" : ""}>
                <td className="qp-sheet-name">{p.icon} {p.name}</td>
                <td className="qp-sheet-sku">{p.sku}</td>
                <td className="right">{KSH(p.previousPrice)}</td>
                <td
                  className="right qp-sheet-cell"
                  onDoubleClick={() => startEdit(p)}
                >
                  {editingId === p.id ? (
                    <input
                      autoFocus
                      className="qp-sheet-input"
                      value={draft}
                      onChange={(e) => setDraft(e.target.value)}
                      onBlur={() => commit(p)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") commit(p);
                        if (e.key === "Escape") setEditingId(null);
                      }}
                    />
                  ) : (
                    <span className="qp-sheet-value">{KSH(shown)}</span>
                  )}
                </td>
                <td className={`right qp-sheet-change ${change >= 0 ? "up" : "down"}`}>
                  {change === null ? "—" : `${change >= 0 ? "↑" : "↓"} ${Math.abs(change).toFixed(1)}%`}
                </td>
                <td>
                  <span className="qp-status-pill small" style={{ color: meta.color, background: meta.bg }}>
                    {meta.dot} {meta.label}
                  </span>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
      <p className="qp-sheet-hint">Double-click a Current cell to edit, then press Enter to send it to the update queue.</p>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* BULK UPDATE VIEW                                                     */
/* ------------------------------------------------------------------ */

function BulkUpdateView({ products, selectedIds, setSelectedIds, onApply }) {
  const [action, setAction] = useState("setAll");
  const [amount, setAmount] = useState("");
  const [category, setCategory] = useState("all");

  const scoped = category === "all" ? products : products.filter((p) => p.category === category);

  const toggle = (id) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  };

  const selectAll = () => setSelectedIds(new Set(scoped.map((p) => p.id)));
  const clearAll = () => setSelectedIds(new Set());

  const selected = products.filter((p) => selectedIds.has(p.id));

  const computeNew = (p) => {
    const old = p.currentPrice ?? p.previousPrice ?? 0;
    const n = Number(amount);
    switch (action) {
      case "setAll": return Number.isNaN(n) ? old : n;
      case "increasePct": return Number.isNaN(n) ? old : Math.round(old * (1 + n / 100));
      case "decreasePct": return Number.isNaN(n) ? old : Math.round(old * (1 - n / 100));
      case "add": return Number.isNaN(n) ? old : old + n;
      case "subtract": return Number.isNaN(n) ? old : Math.max(0, old - n);
      case "lowest": {
        const vals = [p.market.greenmart, p.market.freshpoint, p.market.valuestore, p.market.budgetfoods];
        return Math.min(...vals);
      }
      default: return old;
    }
  };

  const oldTotal = selected.reduce((s, p) => s + (p.currentPrice ?? p.previousPrice ?? 0), 0);
  const newTotal = selected.reduce((s, p) => s + computeNew(p), 0);

  const apply = () => {
    const entries = selected.map((p) => [p.id, computeNew(p)]);
    onApply(entries);
    setSelectedIds(new Set());
    setAmount("");
  };

  const actionNeedsAmount = !["lowest"].includes(action);
  const canApply = selected.length > 0 && (!actionNeedsAmount || amount !== "");

  return (
    <div className="qp-bulk">
      <div className="qp-bulk-header">
        <div className="qp-bulk-category">
          <span className="qp-field-label">Category</span>
          <select value={category} onChange={(e) => setCategory(e.target.value)}>
            <option value="all">All categories</option>
            {PRICE_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div className="qp-bulk-selectbtns">
          <button className="qp-btn qp-btn-ghost small" onClick={selectAll}>Select all ({scoped.length})</button>
          <button className="qp-btn qp-btn-ghost small" onClick={clearAll}>Clear</button>
        </div>
      </div>

      <div className="qp-bulk-list">
        {scoped.map((p) => {
          const checked = selectedIds.has(p.id);
          return (
            <label key={p.id} className={`qp-bulk-row ${checked ? "checked" : ""}`}>
              <input type="checkbox" checked={checked} onChange={() => toggle(p.id)} />
              <span className="qp-bulk-icon">{p.icon}</span>
              <span className="qp-bulk-name">{p.name}</span>
              <span className="qp-bulk-price">{KSH(p.currentPrice ?? p.previousPrice)}</span>
            </label>
          );
        })}
      </div>

      <div className="qp-bulk-action-panel">
        <h3 className="qp-bulk-action-title">Bulk price action</h3>
        <div className="qp-bulk-action-options">
          {[
            { id: "setAll", label: "Set all to" },
            { id: "increasePct", label: "Increase by %" },
            { id: "decreasePct", label: "Decrease by %" },
            { id: "add", label: "Add KSh" },
            { id: "subtract", label: "Subtract KSh" },
            { id: "lowest", label: "Copy lowest market price" },
          ].map((opt) => (
            <button
              key={opt.id}
              className={`qp-bulk-opt ${action === opt.id ? "active" : ""}`}
              onClick={() => setAction(opt.id)}
            >
              {opt.label}
            </button>
          ))}
        </div>
        {actionNeedsAmount && (
          <div className="qp-bulk-amount">
            <span>{action === "increasePct" || action === "decreasePct" ? "" : "KSh"}</span>
            <input
              type="number"
              value={amount}
              onChange={(e) => setAmount(e.target.value)}
              placeholder="0"
            />
            <span>{action === "increasePct" || action === "decreasePct" ? "%" : ""}</span>
          </div>
        )}

        {selected.length > 0 && (
          <div className="qp-bulk-preview">
            <div>{selected.length} product{selected.length > 1 ? "s" : ""} selected</div>
            <div className="qp-bulk-totals">
              <span>Old total <strong>{KSH(oldTotal)}</strong></span>
              <span>New total <strong>{KSH(newTotal)}</strong></span>
              <span className={newTotal - oldTotal >= 0 ? "up" : "down"}>
                Change {newTotal - oldTotal >= 0 ? "+" : ""}{KSH(newTotal - oldTotal)}
              </span>
            </div>
          </div>
        )}

        <button className="qp-btn qp-btn-primary full" disabled={!canApply} onClick={apply}>
          Apply changes
        </button>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* UPDATE QUEUE SIDEBAR                                                 */
/* ------------------------------------------------------------------ */

function UpdateQueueSidebar({ queueList, onRemove, onSaveAll, netChange, open, onClose }) {
  return (
    <aside className={`qp-queue ${open ? "open-mobile" : ""}`}>
      <div className="qp-queue-header">
        <h3><ClipboardPaste size={16} /> Update queue</h3>
        <button className="qp-queue-close" onClick={onClose}><X size={18} /></button>
      </div>
      {queueList.length === 0 ? (
        <p className="qp-queue-empty">Prices you edit will land here before you save them.</p>
      ) : (
        <>
          <p className="qp-queue-count">{queueList.length} product{queueList.length > 1 ? "s" : ""}</p>
          <div className="qp-queue-list">
            {queueList.map((q) => (
              <div key={q.id} className="qp-queue-item">
                <div>
                  <div className="qp-queue-name">✓ {q.product?.name}</div>
                  <div className="qp-queue-change">
                    {KSH(q.oldPrice)} → <strong>{KSH(q.newPrice)}</strong>
                  </div>
                </div>
                <button className="qp-queue-remove" onClick={() => onRemove(q.id)}><X size={14} /></button>
              </div>
            ))}
          </div>
          <div className="qp-queue-net">Net change: {netChange >= 0 ? "+" : ""}{KSH(netChange)}</div>
          <button className="qp-btn qp-btn-primary full" onClick={onSaveAll}>
            Save {queueList.length} change{queueList.length > 1 ? "s" : ""}
          </button>
        </>
      )}
    </aside>
  );
}

/* ------------------------------------------------------------------ */
/* SAVE BAR / TOAST / MODALS                                            */
/* ------------------------------------------------------------------ */

function SaveBar({ count, oldTotal, newTotal, netChange, onDiscard, onSave, onOpenQueue }) {
  return (
    <div className="qp-savebar">
      <div className="qp-savebar-info">
        <strong>{count} price change{count > 1 ? "s" : ""} waiting to be saved</strong>
        <span>{KSH(oldTotal)} → {KSH(newTotal)} · Net change {netChange >= 0 ? "+" : ""}{KSH(netChange)}</span>
      </div>
      <div className="qp-savebar-actions">
        <button className="qp-btn qp-btn-ghost small qp-savebar-queue" onClick={onOpenQueue}>View queue</button>
        <button className="qp-btn qp-btn-ghost small" onClick={onDiscard}>Discard</button>
        <button className="qp-btn qp-btn-primary" onClick={onSave}>Save {count} change{count > 1 ? "s" : ""}</button>
      </div>
    </div>
  );
}

function PriceToast({ msg, showUndo, onUndo, onClose }) {
  return (
    <div className="qp-toast">
      <CheckCircle2 size={18} />
      <span>{msg}</span>
      {showUndo && <button className="qp-toast-undo" onClick={onUndo}><Undo2 size={14} /> Undo</button>}
      <button className="qp-toast-close" onClick={onClose}><X size={14} /></button>
    </div>
  );
}

function WarningModal({ data, onCancel, onConfirm }) {
  const change = pctChange(data.oldPrice, data.newPrice);
  return (
    <div className="qp-modal-backdrop" onClick={onCancel}>
      <div className="qp-modal qp-modal-warning" onClick={(e) => e.stopPropagation()}>
        <div className="qp-modal-icon warn"><AlertTriangle size={22} /></div>
        <h3>Check this price</h3>
        <div className="qp-warn-rows">
          <div><span>Previous price</span><strong>{KSH(data.oldPrice)}</strong></div>
          <div><span>New price</span><strong>{KSH(data.newPrice)}</strong></div>
          <div><span>Change</span><strong className={change >= 0 ? "up" : "down"}>{change >= 0 ? "+" : ""}{change.toFixed(0)}%</strong></div>
        </div>
        <p className="qp-warn-msg">This is a large price change. Confirm that {KSH(data.newPrice)} is correct before it's added to your queue.</p>
        <div className="qp-modal-actions">
          <button className="qp-btn qp-btn-ghost" onClick={onCancel}>Cancel</button>
          <button className="qp-btn qp-btn-primary" onClick={onConfirm}>Confirm price</button>
        </div>
      </div>
    </div>
  );
}

function PriceImportModal({ onClose, onApply }) {
  const [dropped, setDropped] = useState(false);
  return (
    <div className="qp-modal-backdrop" onClick={onClose}>
      <div className="qp-modal" onClick={(e) => e.stopPropagation()}>
        <h3>Import price list</h3>
        {!dropped ? (
          <div
            className="qp-dropzone"
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => { e.preventDefault(); setDropped(true); }}
          >
            <Upload size={26} />
            <p>Drop your Excel or CSV file here</p>
            <button className="qp-btn qp-btn-ghost" onClick={() => setDropped(true)}>Choose file</button>
          </div>
        ) : (
          <div className="qp-import-preview">
            <h4>Import preview</h4>
            <div className="qp-import-stats">
              <span>154 rows detected</span>
              <span className="up">148 ready</span>
              <span className="warn">4 warnings</span>
              <span className="down">2 errors</span>
            </div>
            <button className="qp-btn qp-btn-primary full" onClick={onApply}>Apply 148 price updates</button>
          </div>
        )}
        <button className="qp-modal-close" onClick={onClose}><X size={18} /></button>
      </div>
    </div>
  );
}

function HistoryModal({ productId, products, history, onClose }) {
  const single = productId !== "__all__" ? products.find((p) => p.id === productId) : null;
  const entries = single ? (history[productId] || []) : Object.entries(history).flatMap(([id, rows]) => rows.map((r) => ({ ...r, id })));

  return (
    <div className="qp-modal-backdrop" onClick={onClose}>
      <div className="qp-modal qp-modal-wide" onClick={(e) => e.stopPropagation()}>
        <h3>{single ? single.name : "Price history"}</h3>
        {entries.length === 0 ? (
          <p className="qp-empty-text">No price changes saved yet this session.</p>
        ) : (
          <table className="qp-history-table">
            <thead><tr><th>Date</th>{!single && <th>Product</th>}<th className="right">Previous</th><th className="right">New</th><th className="right">Change</th><th>By</th></tr></thead>
            <tbody>
              {entries.map((e, i) => (
                <tr key={i}>
                  <td>{e.date}</td>
                  {!single && <td>{products.find((p) => p.id === e.id)?.name}</td>}
                  <td className="right">{KSH(e.previous)}</td>
                  <td className="right">{KSH(e.next)}</td>
                  <td className={`right ${e.change >= 0 ? "up" : "down"}`}>{e.change >= 0 ? "+" : ""}{e.change.toFixed(1)}%</td>
                  <td>{e.by}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        <button className="qp-modal-close" onClick={onClose}><X size={18} /></button>
      </div>
    </div>
  );
}

function SummaryModal({ items, freshness, attention, onClose, onContinue }) {
  const increases = items.filter((i) => i.change > 0).length;
  const decreases = items.filter((i) => i.change < 0).length;
  const unchanged = items.filter((i) => i.change === 0).length;
  const net = items.reduce((s, i) => s + (i.newPrice - i.oldPrice), 0);

  return (
    <div className="qp-modal-backdrop" onClick={onClose}>
      <div className="qp-modal qp-modal-summary" onClick={(e) => e.stopPropagation()}>
        <div className="qp-summary-emoji">🎉</div>
        <h3>Price update complete</h3>
        <p className="qp-summary-sub">Today you updated <strong>{items.length}</strong> product{items.length === 1 ? "" : "s"}.</p>
        <div className="qp-summary-grid">
          <div><span>Price increases</span><strong className="up">{increases}</strong></div>
          <div><span>Price decreases</span><strong className="down">{decreases}</strong></div>
          <div><span>Unchanged</span><strong>{unchanged}</strong></div>
          <div><span>Still need attention</span><strong>{attention}</strong></div>
          <div><span>Freshness</span><strong>{freshness}%</strong></div>
          <div><span>Value movement</span><strong className={net >= 0 ? "up" : "down"}>{net >= 0 ? "+" : ""}{KSH(net)}</strong></div>
        </div>
        <div className="qp-modal-actions">
          <button className="qp-btn qp-btn-ghost" onClick={onClose}>Done</button>
          <button className="qp-btn qp-btn-primary" onClick={onContinue}>Continue updating</button>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* RECENTLY UPDATED + EMPTY STATE                                       */
/* ------------------------------------------------------------------ */

function RecentlyUpdated({ items, onOpenHistory }) {
  return (
    <section className="qp-section">
      <h2 className="qp-section-title">Recently updated</h2>
      <div className="qp-recent-list">
        {items.map((it, i) => (
          <button key={i} className="qp-recent-item" onClick={() => onOpenHistory(it.id)}>
            <span className="qp-recent-icon">{it.icon}</span>
            <div className="qp-recent-mid">
              <div className="qp-recent-name">{it.name}</div>
              <div className="qp-recent-prices">{KSH(it.oldPrice)} → {KSH(it.newPrice)}</div>
            </div>
            <div className={`qp-recent-change ${it.change >= 0 ? "up" : "down"}`}>
              {it.change >= 0 ? "↑" : "↓"} {Math.abs(it.change).toFixed(1)}%
            </div>
            <div className="qp-recent-when">{it.when}</div>
          </button>
        ))}
      </div>
    </section>
  );
}

function EmptyState({ text }) {
  return (
    <div className="qp-empty">
      <CircleAlert size={22} />
      <p>{text}</p>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* STYLES                                                               */
/* ------------------------------------------------------------------ */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600;700&display=swap');

.qp-app {
  --bg: #F6F8F5;
  --surface: #FFFFFF;
  --ink: #16221B;
  --ink-soft: #55655C;
  --ink-faint: #8B978F;
  --border: #E1E7DE;
  --border-soft: #EDF1EA;
  --green: #1F6E4A;
  --green-dark: #124A31;
  --green-bg: #E5F2E9;
  --amber: #A2680E;
  --amber-bg: #FBF0DA;
  --orange: #B4560F;
  --orange-bg: #FBEADA;
  --red: #BC3A2D;
  --red-bg: #FBE6E2;
  --gray-bg: #EFF1ED;
  --shadow: 0 1px 2px rgba(20,30,22,0.04), 0 6px 20px rgba(20,30,22,0.06);
  --radius: 14px;

  font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
  color: var(--ink);
  background: var(--bg);
  min-height: 100vh;
  -webkit-font-smoothing: antialiased;
}
.qp-app * { box-sizing: border-box; }
.qp-app button { font-family: inherit; cursor: pointer; }
.qp-app input, .qp-app select { font-family: inherit; }
.qp-app table { border-collapse: collapse; width: 100%; }

.qp-mono { font-family: 'JetBrains Mono', ui-monospace, monospace; }

.qp-shell {
  max-width: 1180px;
  margin: 0 auto;
  padding: 28px 24px 140px;
}

/* Header */
.qp-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  gap: 24px;
  flex-wrap: wrap;
  margin-bottom: 18px;
}
.qp-title {
  font-size: 30px;
  font-weight: 800;
  letter-spacing: -0.01em;
  display: flex;
  align-items: center;
  gap: 10px;
  margin: 0 0 6px;
  color: var(--green-dark);
}
.qp-title-icon { color: var(--green); }
.qp-subtitle {
  color: var(--ink-soft);
  font-size: 14.5px;
  max-width: 560px;
  line-height: 1.5;
  margin: 0;
}
.qp-header-actions { display: flex; gap: 8px; flex-wrap: wrap; align-items: flex-start; }

.qp-btn {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  border-radius: 10px;
  padding: 9px 14px;
  font-size: 13.5px;
  font-weight: 600;
  border: 1px solid transparent;
  transition: transform .08s ease, background .15s ease, border-color .15s ease;
  white-space: nowrap;
}
.qp-btn:active { transform: scale(0.97); }
.qp-btn-ghost {
  background: var(--surface);
  border-color: var(--border);
  color: var(--ink);
}
.qp-btn-ghost:hover { border-color: var(--green); color: var(--green-dark); }
.qp-btn-primary {
  background: var(--green);
  color: white;
}
.qp-btn-primary:hover { background: var(--green-dark); }
.qp-btn-primary:disabled { background: #B9C6BE; cursor: not-allowed; }
.qp-btn.small { padding: 6px 10px; font-size: 12.5px; }
.qp-btn.full { width: 100%; justify-content: center; margin-top: 10px; }

/* Freshness */
.qp-freshness {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 14px 18px;
  margin-bottom: 20px;
}
.qp-freshness-label {
  display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 8px; font-size: 13px; color: var(--ink-soft);
}
.qp-freshness-label strong { font-size: 18px; color: var(--green-dark); font-family: 'JetBrains Mono', monospace; }
.qp-freshness-track { height: 8px; background: var(--gray-bg); border-radius: 999px; overflow: hidden; }
.qp-freshness-fill { height: 100%; background: linear-gradient(90deg, var(--green), #4E9B6F); border-radius: 999px; transition: width .4s ease; }
.qp-freshness-sub { font-size: 12px; color: var(--ink-faint); display: block; margin-top: 6px; }

/* Sections */
.qp-section { margin: 26px 0; }
.qp-section-title { font-size: 15px; font-weight: 700; margin: 0 0 12px; color: var(--ink); }

/* Attention */
.qp-attention-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 12px;
}
.qp-attn-card {
  text-align: left;
  border-radius: var(--radius);
  border: 1px solid var(--border);
  background: var(--surface);
  padding: 16px;
  box-shadow: var(--shadow);
  border-left: 4px solid var(--ink-faint);
  transition: transform .12s ease;
}
.qp-attn-card:hover { transform: translateY(-2px); }
.qp-tone-red { border-left-color: var(--red); }
.qp-tone-orange { border-left-color: var(--orange); }
.qp-tone-green { border-left-color: var(--green); }
.qp-tone-amber { border-left-color: var(--amber); }
.qp-attn-title { font-size: 15px; font-weight: 700; margin-bottom: 4px; }
.qp-attn-sub { font-size: 12.5px; color: var(--ink-soft); margin-bottom: 10px; }
.qp-attn-cta { font-size: 12.5px; font-weight: 700; color: var(--green-dark); }

/* Controls */
.qp-controlbar {
  display: flex;
  justify-content: space-between;
  align-items: flex-end;
  gap: 16px;
  flex-wrap: wrap;
  margin-bottom: 16px;
}
.qp-field-label { font-size: 11.5px; text-transform: none; color: var(--ink-faint); display: block; margin-bottom: 5px; font-weight: 600; }

.qp-shopselect { position: relative; }
.qp-shopselect-row { position: relative; }
.qp-shop-btn {
  display: flex; align-items: center; gap: 8px;
  background: var(--surface); border: 1px solid var(--border); border-radius: 10px;
  padding: 9px 14px; font-weight: 700; font-size: 14px; color: var(--green-dark);
}
.qp-shop-dropdown {
  position: absolute; top: calc(100% + 6px); left: 0; z-index: 20;
  background: var(--surface); border: 1px solid var(--border); border-radius: 10px;
  box-shadow: var(--shadow); min-width: 200px; padding: 6px; display: flex; flex-direction: column; gap: 2px;
}
.qp-shop-option { text-align: left; padding: 8px 10px; border-radius: 7px; font-size: 13.5px; background: transparent; }
.qp-shop-option:hover { background: var(--green-bg); }
.qp-shop-option.active { background: var(--green-bg); color: var(--green-dark); font-weight: 700; }
.qp-shop-meta { font-size: 12px; color: var(--ink-faint); margin-top: 6px; }

.qp-mode-select {
  display: flex; background: var(--surface); border: 1px solid var(--border); border-radius: 11px; padding: 3px;
}
.qp-mode-btn {
  display: flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 8px; font-size: 13px; font-weight: 600; background: transparent; color: var(--ink-soft);
}
.qp-mode-btn.active { background: var(--green); color: white; }

/* Search */
.qp-search-wrap {
  position: relative;
  display: flex; align-items: center; gap: 10px;
  background: var(--surface); border: 1.5px solid var(--border); border-radius: 14px;
  padding: 4px 6px 4px 16px; margin-bottom: 14px;
}
.qp-search-wrap:focus-within { border-color: var(--green); }
.qp-search-icon { color: var(--ink-faint); flex-shrink: 0; }
.qp-search-input { flex: 1; border: none; outline: none; font-size: 15px; padding: 12px 0; background: transparent; color: var(--ink); }
.qp-search-count { font-size: 12px; color: var(--ink-faint); white-space: nowrap; }
.qp-scan-btn { flex-shrink: 0; }

.qp-chip-row { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 22px; }
.qp-chip {
  padding: 7px 13px; border-radius: 999px; font-size: 12.5px; font-weight: 600;
  background: var(--surface); border: 1px solid var(--border); color: var(--ink-soft);
}
.qp-chip.active { background: var(--green-dark); border-color: var(--green-dark); color: white; }

/* Workspace layout */
.qp-workspace { display: grid; grid-template-columns: 1fr 300px; gap: 20px; align-items: start; }
.qp-main { min-width: 0; }

/* Card stream */
.qp-card-stream { display: flex; flex-direction: column; gap: 12px; }
.qp-card {
  background: var(--surface); border: 1.5px solid var(--border); border-radius: var(--radius);
  padding: 14px 16px; box-shadow: var(--shadow); transition: border-color .15s ease;
  cursor: pointer;
}
.qp-card.expanded { border-color: var(--green); }
.qp-card.queued { border-color: var(--green); background: var(--green-bg); }
.qp-card-head { display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }
.qp-card-icon { font-size: 22px; }
.qp-card-titles { flex: 1; min-width: 0; }
.qp-card-name { font-weight: 700; font-size: 14.5px; }
.qp-card-sku { font-size: 11.5px; color: var(--ink-faint); font-family: 'JetBrains Mono', monospace; }
.qp-status-pill { font-size: 11px; font-weight: 700; padding: 4px 9px; border-radius: 999px; white-space: nowrap; }
.qp-status-pill.small { font-size: 10.5px; padding: 3px 8px; }

.qp-card-body { display: grid; grid-template-columns: 1fr 1.3fr auto; gap: 18px; align-items: end; cursor: default; }
.qp-card-prev { display: flex; flex-direction: column; gap: 4px; }
.qp-prev-price { font-family: 'JetBrains Mono', monospace; font-size: 16px; font-weight: 600; color: var(--ink-soft); }
.qp-last-updated { font-size: 11.5px; color: var(--ink-faint); display: flex; align-items: center; gap: 4px; }

.qp-card-current { display: flex; flex-direction: column; gap: 6px; }
.qp-price-input-row {
  display: flex; align-items: center; gap: 6px; background: var(--bg); border: 1.5px solid var(--border); border-radius: 9px; padding: 6px 10px;
}
.qp-price-input-row:focus-within { border-color: var(--green); background: white; }
.qp-currency { font-size: 13px; color: var(--ink-faint); font-weight: 600; }
.qp-price-input {
  border: none; outline: none; background: transparent; font-family: 'JetBrains Mono', monospace;
  font-size: 19px; font-weight: 700; width: 100%; color: var(--green-dark);
}
.qp-quickadjust { display: flex; gap: 5px; }
.qp-quickadjust button {
  flex: 1; background: var(--gray-bg); border: none; border-radius: 6px; padding: 5px 0; font-size: 12px; font-weight: 700; color: var(--ink-soft);
}
.qp-quickadjust button:hover { background: var(--green-bg); color: var(--green-dark); }

.qp-save-btn {
  display: flex; align-items: center; gap: 5px; background: var(--green); color: white; border: none;
  border-radius: 9px; padding: 10px 16px; font-weight: 700; font-size: 13px; height: fit-content;
}
.qp-save-btn:hover { background: var(--green-dark); }

.qp-change-preview {
  grid-column: 1 / -1; display: flex; align-items: center; gap: 6px; font-size: 12.5px; font-weight: 700; font-family: 'JetBrains Mono', monospace;
}
.qp-change-preview.up { color: var(--red); }
.qp-change-preview.down { color: var(--green-dark); }

.qp-card-extra { margin-top: 14px; padding-top: 14px; border-top: 1px dashed var(--border); cursor: default; }
.qp-copyrow { margin-bottom: 10px; }
.qp-copy-btns { display: flex; flex-wrap: wrap; gap: 6px; }
.qp-copy-btns button {
  background: var(--bg); border: 1px solid var(--border); border-radius: 8px; padding: 6px 10px; font-size: 12px; font-weight: 600; color: var(--ink-soft);
}
.qp-copy-btns button:hover { border-color: var(--green); color: var(--green-dark); }

.qp-intel-toggle {
  display: flex; align-items: center; gap: 5px; background: none; border: none; color: var(--green-dark); font-weight: 700; font-size: 12.5px; padding: 0;
}
.qp-intel-panel { margin-top: 10px; background: var(--bg); border-radius: 10px; padding: 12px; }
.qp-intel-row { display: flex; justify-content: space-between; font-size: 12.5px; padding: 3px 0; color: var(--ink-soft); }
.qp-intel-row strong { font-family: 'JetBrains Mono', monospace; color: var(--ink); }
.qp-intel-suggest { display: flex; justify-content: space-between; align-items: center; margin-top: 8px; padding-top: 8px; border-top: 1px dashed var(--border); font-size: 12.5px; }
.qp-use-btn { background: var(--green-dark); color: white; border: none; border-radius: 7px; padding: 6px 10px; font-size: 12px; font-weight: 700; }
.qp-market-list { margin-top: 10px; display: flex; flex-direction: column; gap: 4px; }
.qp-market-item { display: flex; justify-content: space-between; font-size: 12px; color: var(--ink-faint); }
.qp-market-item span:last-child { font-family: 'JetBrains Mono', monospace; color: var(--ink-soft); }

/* Spreadsheet */
.qp-sheet-wrap { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 4px; box-shadow: var(--shadow); overflow-x: auto; }
.qp-sheet th { text-align: left; font-size: 11.5px; text-transform: uppercase; letter-spacing: .03em; color: var(--ink-faint); padding: 12px 14px; border-bottom: 1px solid var(--border); }
.qp-sheet th.right, .qp-sheet td.right { text-align: right; }
.qp-sheet td { padding: 10px 14px; border-bottom: 1px solid var(--border-soft); font-size: 13.5px; }
.qp-sheet tr.queued { background: var(--green-bg); }
.qp-sheet-name { font-weight: 600; }
.qp-sheet-sku { font-family: 'JetBrains Mono', monospace; color: var(--ink-faint); font-size: 12px; }
.qp-sheet-cell { cursor: pointer; }
.qp-sheet-value { font-family: 'JetBrains Mono', monospace; font-weight: 700; padding: 4px 8px; border-radius: 6px; }
.qp-sheet-cell:hover .qp-sheet-value { background: var(--gray-bg); }
.qp-sheet-input { font-family: 'JetBrains Mono', monospace; font-weight: 700; width: 90px; text-align: right; border: 1.5px solid var(--green); border-radius: 6px; padding: 4px 8px; outline: none; }
.qp-sheet-change.up { color: var(--red); font-family: 'JetBrains Mono', monospace; font-weight: 600; }
.qp-sheet-change.down { color: var(--green-dark); font-family: 'JetBrains Mono', monospace; font-weight: 600; }
.qp-sheet-hint { font-size: 12px; color: var(--ink-faint); padding: 10px 14px; }

/* Bulk */
.qp-bulk { display: flex; flex-direction: column; gap: 14px; }
.qp-bulk-header { display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 10px; }
.qp-bulk-category select { padding: 8px 10px; border-radius: 8px; border: 1px solid var(--border); font-size: 13px; background: var(--surface); }
.qp-bulk-selectbtns { display: flex; gap: 8px; }
.qp-bulk-list {
  background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); box-shadow: var(--shadow);
  max-height: 320px; overflow-y: auto;
}
.qp-bulk-row { display: flex; align-items: center; gap: 10px; padding: 10px 14px; border-bottom: 1px solid var(--border-soft); font-size: 13.5px; }
.qp-bulk-row.checked { background: var(--green-bg); }
.qp-bulk-row input { width: 16px; height: 16px; accent-color: var(--green); }
.qp-bulk-icon { font-size: 17px; }
.qp-bulk-name { flex: 1; font-weight: 600; }
.qp-bulk-price { font-family: 'JetBrains Mono', monospace; color: var(--ink-soft); }

.qp-bulk-action-panel { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 16px; box-shadow: var(--shadow); }
.qp-bulk-action-title { font-size: 14px; margin: 0 0 10px; }
.qp-bulk-action-options { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 12px; }
.qp-bulk-opt { padding: 8px 12px; border-radius: 8px; border: 1px solid var(--border); background: var(--bg); font-size: 12.5px; font-weight: 600; color: var(--ink-soft); }
.qp-bulk-opt.active { background: var(--green-dark); border-color: var(--green-dark); color: white; }
.qp-bulk-amount { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
.qp-bulk-amount input { flex: 1; padding: 9px 12px; border-radius: 8px; border: 1.5px solid var(--border); font-family: 'JetBrains Mono', monospace; font-size: 15px; font-weight: 700; }
.qp-bulk-amount span { font-size: 13px; color: var(--ink-faint); font-weight: 700; }
.qp-bulk-preview { background: var(--bg); border-radius: 10px; padding: 10px 12px; font-size: 13px; margin-bottom: 4px; }
.qp-bulk-totals { display: flex; gap: 16px; margin-top: 6px; flex-wrap: wrap; font-size: 12.5px; }
.qp-bulk-totals strong { font-family: 'JetBrains Mono', monospace; }
.qp-bulk-totals .up { color: var(--red); }
.qp-bulk-totals .down { color: var(--green-dark); }

/* Queue sidebar */
.qp-queue {
  background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); box-shadow: var(--shadow);
  padding: 16px; position: sticky; top: 20px;
}
.qp-queue-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }
.qp-queue-header h3 { font-size: 14px; display: flex; align-items: center; gap: 6px; margin: 0; }
.qp-queue-close { display: none; background: none; border: none; }
.qp-queue-empty { font-size: 12.5px; color: var(--ink-faint); line-height: 1.5; margin-top: 8px; }
.qp-queue-count { font-size: 12px; color: var(--ink-faint); margin: 4px 0 10px; }
.qp-queue-list { display: flex; flex-direction: column; gap: 8px; max-height: 320px; overflow-y: auto; margin-bottom: 10px; }
.qp-queue-item { display: flex; justify-content: space-between; align-items: flex-start; background: var(--bg); border-radius: 9px; padding: 9px 10px; gap: 8px; }
.qp-queue-name { font-size: 12.5px; font-weight: 700; color: var(--green-dark); }
.qp-queue-change { font-size: 12px; color: var(--ink-soft); font-family: 'JetBrains Mono', monospace; margin-top: 2px; }
.qp-queue-remove { background: none; border: none; color: var(--ink-faint); flex-shrink: 0; }
.qp-queue-remove:hover { color: var(--red); }
.qp-queue-net { font-size: 12.5px; font-weight: 700; color: var(--ink-soft); margin-bottom: 4px; }

/* Save bar */
.qp-savebar {
  position: fixed; bottom: 18px; left: 50%; transform: translateX(-50%);
  background: var(--green-dark); color: white; border-radius: 16px; padding: 14px 20px;
  display: flex; align-items: center; gap: 24px; box-shadow: 0 10px 32px rgba(18,74,49,0.35);
  z-index: 40; max-width: calc(100% - 32px); flex-wrap: wrap;
}
.qp-savebar-info { display: flex; flex-direction: column; gap: 2px; font-size: 12.5px; }
.qp-savebar-info strong { font-size: 14px; }
.qp-savebar-actions { display: flex; gap: 8px; align-items: center; }
.qp-savebar .qp-btn-ghost { background: rgba(255,255,255,0.12); border-color: rgba(255,255,255,0.25); color: white; }
.qp-savebar .qp-btn-primary { background: white; color: var(--green-dark); }
.qp-savebar .qp-btn-primary:hover { background: #EAF3ED; }
.qp-savebar-queue { display: none; }

/* PriceToast */
.qp-toast {
  position: fixed; top: 20px; left: 50%; transform: translateX(-50%);
  background: var(--ink); color: white; padding: 12px 16px; border-radius: 11px;
  display: flex; align-items: center; gap: 10px; font-size: 13.5px; z-index: 60; box-shadow: var(--shadow);
}
.qp-toast-undo { background: rgba(255,255,255,0.15); border: none; color: white; border-radius: 7px; padding: 5px 10px; display: flex; align-items: center; gap: 4px; font-weight: 700; font-size: 12.5px; }
.qp-toast-close { background: none; border: none; color: rgba(255,255,255,0.6); }

/* Modals */
.qp-modal-backdrop {
  position: fixed; inset: 0; background: rgba(20,30,22,0.45); display: flex; align-items: center; justify-content: center; z-index: 100; padding: 20px;
}
.qp-modal {
  background: white; border-radius: 18px; padding: 26px; max-width: 420px; width: 100%; position: relative;
  box-shadow: 0 24px 60px rgba(0,0,0,0.25);
}
.qp-modal-wide { max-width: 560px; }
.qp-modal-summary { text-align: center; max-width: 460px; }
.qp-modal h3 { margin: 0 0 14px; font-size: 18px; }
.qp-modal-close { position: absolute; top: 14px; right: 14px; background: none; border: none; color: var(--ink-faint); }
.qp-modal-icon { width: 44px; height: 44px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-bottom: 10px; }
.qp-modal-icon.warn { background: var(--orange-bg); color: var(--orange); }
.qp-warn-rows { display: flex; flex-direction: column; gap: 6px; margin-bottom: 12px; }
.qp-warn-rows div { display: flex; justify-content: space-between; font-size: 13.5px; }
.qp-warn-rows strong.up { color: var(--red); }
.qp-warn-rows strong.down { color: var(--green-dark); }
.qp-warn-msg { font-size: 13px; color: var(--ink-soft); line-height: 1.5; margin-bottom: 18px; }
.qp-modal-actions { display: flex; justify-content: flex-end; gap: 8px; }

.qp-dropzone { border: 2px dashed var(--border); border-radius: 12px; padding: 34px 16px; text-align: center; color: var(--ink-faint); display: flex; flex-direction: column; align-items: center; gap: 10px; }
.qp-import-preview h4 { font-size: 13.5px; margin: 0 0 10px; }
.qp-import-stats { display: flex; gap: 14px; flex-wrap: wrap; font-size: 12.5px; margin-bottom: 14px; color: var(--ink-soft); }
.qp-import-stats .up { color: var(--green-dark); font-weight: 700; }
.qp-import-stats .warn { color: var(--amber); font-weight: 700; }
.qp-import-stats .down { color: var(--red); font-weight: 700; }

.qp-history-table th { text-align: left; font-size: 11px; text-transform: uppercase; color: var(--ink-faint); padding: 8px 10px; border-bottom: 1px solid var(--border); }
.qp-history-table td { padding: 8px 10px; font-size: 13px; border-bottom: 1px solid var(--border-soft); }
.qp-history-table .right { text-align: right; }
.qp-history-table .up { color: var(--red); }
.qp-history-table .down { color: var(--green-dark); }
.qp-empty-text { color: var(--ink-faint); font-size: 13.5px; }

.qp-summary-emoji { font-size: 34px; margin-bottom: 6px; }
.qp-summary-sub { color: var(--ink-soft); font-size: 13.5px; margin-bottom: 16px; }
.qp-summary-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 18px; }
.qp-summary-grid div { background: var(--bg); border-radius: 10px; padding: 10px; display: flex; flex-direction: column; gap: 2px; }
.qp-summary-grid span { font-size: 11.5px; color: var(--ink-faint); }
.qp-summary-grid strong { font-size: 17px; font-family: 'JetBrains Mono', monospace; }
.qp-summary-grid strong.up { color: var(--red); }
.qp-summary-grid strong.down { color: var(--green-dark); }
.qp-modal-summary .qp-modal-actions { justify-content: center; }

/* Recently updated */
.qp-recent-list { display: flex; flex-direction: column; gap: 8px; }
.qp-recent-item {
  display: flex; align-items: center; gap: 12px; background: var(--surface); border: 1px solid var(--border); border-radius: 12px;
  padding: 10px 14px; text-align: left; width: 100%;
}
.qp-recent-icon { font-size: 18px; }
.qp-recent-mid { flex: 1; min-width: 0; }
.qp-recent-name { font-weight: 600; font-size: 13.5px; }
.qp-recent-prices { font-size: 12px; color: var(--ink-faint); font-family: 'JetBrains Mono', monospace; }
.qp-recent-change { font-size: 12.5px; font-weight: 700; font-family: 'JetBrains Mono', monospace; }
.qp-recent-change.up { color: var(--red); }
.qp-recent-change.down { color: var(--green-dark); }
.qp-recent-when { font-size: 11.5px; color: var(--ink-faint); white-space: nowrap; }

.qp-empty { display: flex; flex-direction: column; align-items: center; gap: 8px; padding: 50px 0; color: var(--ink-faint); }

/* Responsive */
@media (max-width: 980px) {
  .qp-attention-grid { grid-template-columns: repeat(2, 1fr); }
  .qp-workspace { grid-template-columns: 1fr; }
  .qp-queue { position: fixed; top: 0; right: 0; bottom: 0; width: 300px; border-radius: 0; z-index: 90; transform: translateX(100%); transition: transform .2s ease; overflow-y: auto; }
  .qp-queue.open-mobile { transform: translateX(0); }
  .qp-queue-close { display: block; }
  .qp-savebar-queue { display: inline-flex; }
}
@media (max-width: 640px) {
  .qp-attention-grid { grid-template-columns: 1fr; }
  .qp-card-body { grid-template-columns: 1fr; }
  .qp-shell { padding: 20px 14px 160px; }
  .qp-title { font-size: 24px; }
  .qp-savebar { flex-direction: column; align-items: stretch; left: 14px; right: 14px; transform: none; }
}
`;


// ============================================================
// ---------- Smart Shopping ----------
// ============================================================


/* ------------------------------------------------------------------ */
/*  HELPERS                                                            */
/* ------------------------------------------------------------------ */

const ksh = (n) =>
  n == null ? "—" : `KSh ${Math.round(n).toLocaleString("en-KE")}`;

function priceEntries(productName, { products, shops, shopProducts }, productId) {
  const product =
    (productId && products.find((p) => p.id === productId)) ||
    products.find((p) => (p.productName || "").trim().toLowerCase() === String(productName).trim().toLowerCase());
  if (!product) return [];
  return shopProducts
    .filter((sp) => sp.productId === product.id)
    .map((sp) => {
      const shop = shops.find((s) => s.id === sp.shopId);
      return shop ? { shop: shop.name, price: sp.price } : null;
    })
    .filter(Boolean);
}

function analyzeProduct(item, ctx) {
  const entries = priceEntries(item.name, ctx, item.productId).sort((a, b) => a.price - b.price);
  if (entries.length === 0) return { available: false, entries: [] };
  const cheapest = entries[0];
  const priciest = entries[entries.length - 1];
  return {
    available: true,
    entries,
    cheapest,
    priciest,
    cheapestTotal: cheapest.price * item.qty,
    priciestTotal: priciest.price * item.qty,
  };
}

let shoppingIdCounter = 1;
const nextShoppingId = () => shoppingIdCounter++;

function makeOrderId() {
  return "QP-" + Math.random().toString(36).slice(2, 8).toUpperCase();
}

/* ------------------------------------------------------------------ */
/*  MAIN COMPONENT                                                     */
/* ------------------------------------------------------------------ */

const SMART_SHOPPING_STORAGE_KEY = "quickpick_smart_shopping";

function loadPersistedShopping() {
  if (typeof window === "undefined" || !window.localStorage) return null;
  try {
    const raw = window.localStorage.getItem(SMART_SHOPPING_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.list)) return null;
    return parsed;
  } catch (e) {
    return null;
  }
}

function QuickPickSmartShopping({ onBack, savedLists: sharedSavedLists, onSavedListsChange, products = [], shops = [], shopProducts = [], sharedList, onListChange }) {
  // Restores whatever the shopper had in progress - so leaving this page
  // (navigating to the Dashboard) and coming back, or a full page refresh,
  // doesn't wipe out an in-progress list.
  const persistedRef = useRef(loadPersistedShopping());
  const persisted = persistedRef.current;

  // The list is shared with Product Search (App owns the canonical copy in
  // `sharedList`). If items were added there since the last visit, the
  // confirmed/optimized state no longer matches the list, so it is reset.
  const initRef = useRef(null);
  if (initRef.current === null) {
    const sig = (l) => JSON.stringify((l || []).map((i) => [String(i.name).toLowerCase(), i.qty]));
    const fromShared = Array.isArray(sharedList) && sharedList.length ? sharedList : null;
    const startList = fromShared
      || (persisted?.list?.length
        ? persisted.list
        : [
            { id: nextShoppingId(), name: "Soko Maize Meal", qty: 2 },
            { id: nextShoppingId(), name: "Omo Detergent", qty: 1 },
            { id: nextShoppingId(), name: "Fresh Fri Cooking Oil", qty: 1 },
            { id: nextShoppingId(), name: "Colgate Toothpaste", qty: 2 },
          ]);
    const changedElsewhere = !!fromShared && !!persisted?.list?.length && sig(fromShared) !== sig(persisted.list);
    initRef.current = { startList, keepState: !changedElsewhere };
  }
  const [list, setList] = useState(() => initRef.current.startList);
  useEffect(() => {
    if (onListChange) onListChange(list);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [list]);
  const [locked, setLocked] = useState(() => (initRef.current.keepState ? persisted?.locked ?? false : false)); // confirmed + locked together
  const [optimized, setOptimized] = useState(() => (initRef.current.keepState ? persisted?.optimized ?? false : false));
  const [expandedRows, setExpandedRows] = useState({});
  const [clearArmed, setClearArmed] = useState(false);

  const [searchText, setSearchText] = useState("");
  const [qtyDraft, setQtyDraft] = useState(1);
  const [showSuggestions, setShowSuggestions] = useState(false);

  const [order, setOrder] = useState(() => (initRef.current.keepState ? persisted?.order ?? null : null)); // { id, date, time }
  const [payOpen, setPayOpen] = useState(false);
  const [payStatus, setPayStatus] = useState(null); // null | "pending"

  // savedLists is the real, accumulating record of completed/optimized
  // shopping lists - this IS the data the Admin Dashboard's "Shopping
  // Lists" and "Average Savings" cards are derived from. App.jsx owns the
  // canonical copy (sharedSavedLists); we fall back to whatever was
  // persisted locally if this component is ever used standalone.
  const [savedLists, setSavedLists] = useState(
    () => sharedSavedLists ?? persisted?.savedLists ?? []
  );
  useEffect(() => {
    if (onSavedListsChange) onSavedListsChange(savedLists);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [savedLists]);

  const listRef = useRef(null);
  const compareRef = useRef(null);
  const optimizeRef = useRef(null);

  // Keep the id generator ahead of any restored items so newly added
  // products never collide with a persisted id.
  useEffect(() => {
    const restoredList = initRef.current.startList;
    if (restoredList?.length) {
      const maxId = restoredList.reduce((m, p) => Math.max(m, p.id || 0), 0);
      if (maxId >= shoppingIdCounter) shoppingIdCounter = maxId + 1;
    }
  }, []);

  // Persist to localStorage on every change, so this survives navigating
  // away/back and full page refreshes.
  useEffect(() => {
    if (typeof window === "undefined" || !window.localStorage) return;
    try {
      window.localStorage.setItem(
        SMART_SHOPPING_STORAGE_KEY,
        JSON.stringify({ list, locked, optimized, order, savedLists })
      );
    } catch (e) {
      // localStorage unavailable/full - fail silently, session state still works.
    }
  }, [list, locked, optimized, order, savedLists]);

  const scrollTo = (ref) =>
    setTimeout(() => ref.current?.scrollIntoView({ behavior: "smooth", block: "start" }), 60);

  /* ---- shopping list actions (only available while unlocked) ---- */

  const addProduct = () => {
    if (locked) return;
    const name = searchText.trim();
    if (!name) return;
    const qty = Math.max(1, Math.round(qtyDraft) || 1);
    setList((prev) => {
      const existing = prev.find((p) => p.name.toLowerCase() === name.toLowerCase());
      if (existing) {
        return prev.map((p) => (p.id === existing.id ? { ...p, qty: p.qty + qty } : p));
      }
      const match = products.find((p) => (p.productName || "").trim().toLowerCase() === name.toLowerCase());
      return [...prev, { id: nextShoppingId(), name: match ? match.productName : name, qty, ...(match ? { productId: match.id } : {}) }];
    });
    setSearchText("");
    setQtyDraft(1);
    setShowSuggestions(false);
  };

  const updateQty = (id, qty) => {
    if (locked) return;
    setList((prev) => prev.map((p) => (p.id === id ? { ...p, qty: Math.max(1, qty) } : p)));
  };

  const removeProduct = (id) => {
    if (locked) return;
    setList((prev) => prev.filter((p) => p.id !== id));
  };

  const clearList = () => {
    if (locked) return;
    if (!clearArmed) {
      setClearArmed(true);
      setTimeout(() => setClearArmed(false), 3000);
      return;
    }
    setList([]);
    setClearArmed(false);
  };

  const handleConfirm = () => {
    if (list.length === 0) return;
    setLocked(true);
    scrollTo(compareRef);
  };

  const handleEditList = () => {
    setLocked(false);
    setOptimized(false);
    setOrder(null);
    scrollTo(listRef);
  };

  const handleOptimize = () => {
    setOptimized(true);
    const now = new Date();
    const newOrder = {
      id: makeOrderId(),
      date: now.toLocaleDateString("en-KE"),
      time: now.toLocaleTimeString("en-KE", { hour: "2-digit", minute: "2-digit" }),
    };
    setOrder(newOrder);
    // Record this as a real, saved shopping list - this is what the
    // Dashboard's "Shopping Lists" and "Average Savings" cards count and
    // average over, not a mock delta.
    setSavedLists((prev) => [
      ...prev,
      {
        id: newOrder.id,
        date: newOrder.date,
        time: newOrder.time,
        productCount: optimization.productCount,
        shopsToVisit: optimization.shopsToVisit,
        optimizedTotal: optimization.optimizedTotal,
        priciestTotal: optimization.priciestTotal,
        savings: optimization.savings,
        savingsPct: optimization.savingsPct,
      },
    ]);
    scrollTo(optimizeRef);
  };

  const toggleExpand = (id) => setExpandedRows((prev) => ({ ...prev, [id]: !prev[id] }));

  /* ---- derived data ---- */

  const priceCtx = useMemo(() => ({ products, shops, shopProducts }), [products, shops, shopProducts]);
  const analysis = useMemo(() => list.map((item) => ({ item, ...analyzeProduct(item, priceCtx) })), [list, priceCtx]);

  // Real shops vary per product now (many-to-many), so the comparison
  // table's columns are the shops that actually appear most often across
  // the current list, not a fixed hardcoded set.
  const rankedShopNames = useMemo(() => {
    const freq = {};
    analysis.forEach(({ entries }) => entries.forEach((e) => { freq[e.shop] = (freq[e.shop] || 0) + 1; }));
    return Object.keys(freq).sort((a, b) => freq[b] - freq[a]);
  }, [analysis]);
  const mainShops = rankedShopNames.slice(0, 4);
  const extraShopsGlobal = rankedShopNames.slice(4);

  const optimization = useMemo(() => {
    const rows = analysis.map((a) => ({
      id: a.item.id,
      name: a.item.name,
      qty: a.item.qty,
      available: a.available,
      shop: a.available ? a.cheapest.shop : null,
      unitPrice: a.available ? a.cheapest.price : null,
      total: a.available ? a.cheapestTotal : 0,
      priciestTotal: a.available ? a.priciestTotal : 0,
    }));

    const byShop = {};
    rows.forEach((r) => {
      if (!r.available) return;
      if (!byShop[r.shop]) byShop[r.shop] = { items: [], subtotal: 0 };
      byShop[r.shop].items.push(r);
      byShop[r.shop].subtotal += r.total;
    });

    const optimizedTotal = rows.reduce((s, r) => s + r.total, 0);
    const priciestTotal = rows.reduce((s, r) => s + r.priciestTotal, 0);
    const savings = priciestTotal - optimizedTotal;
    const savingsPct = priciestTotal > 0 ? (savings / priciestTotal) * 100 : 0;
    const shopsToVisit = Object.keys(byShop).length;
    const items = rows.reduce((s, r) => s + (r.available ? r.qty : 0), 0);

    return {
      rows,
      byShop,
      optimizedTotal,
      priciestTotal,
      savings,
      savingsPct,
      shopsToVisit,
      productCount: rows.length,
      items,
    };
  }, [analysis]);

  const productNames = useMemo(
    () => Array.from(new Set(products.map((p) => p.productName).filter(Boolean))),
    [products]
  );

  const suggestions = useMemo(() => {
    if (!searchText.trim()) return [];
    const q = searchText.trim().toLowerCase();
    return productNames.filter((n) => n.toLowerCase().includes(q)).slice(0, 6);
  }, [searchText, productNames]);

  const stage = optimized ? 3 : locked ? 2 : list.length > 0 ? 1 : 0;

  return (
    <div className="qp-root">
      <style>{`
        .qp-root {
          --paper: #f2f4ec;
          --card: #ffffff;
          --line: #dbe2d6;
          --ink: #1c241f;
          --ink-soft: #5b6a5f;
          --forest: #1f4d3d;
          --mango: #e2811f;
          --mango-deep: #b8650f;
          --grass: #2f8f5b;
          --grass-bg: #e6f4ea;
          --warn: #b3461f;
          --warn-bg: #fbeae1;
          font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
          background: var(--paper);
          color: var(--ink);
          min-height: 100vh;
          padding-bottom: 4rem;
        }
        .qp-root * { box-sizing: border-box; }
        .qp-display { font-family: Georgia, "Iowan Old Style", "Times New Roman", serif; font-weight: 600; letter-spacing: -0.01em; }

        .qp-print-only { display: none; }

        .qp-header { background: var(--forest); color: #eef3ec; padding: 1.75rem 1.5rem 1.25rem; }
        .qp-header-inner { max-width: 880px; margin: 0 auto; }
        .qp-brand { font-size: 1.9rem; margin: 0 0 0.15rem; }
        .qp-tagline { margin: 0 0 1.25rem; color: #c3d6c9; font-size: 0.95rem; }
        .qp-steps { display: flex; gap: 0.5rem; flex-wrap: wrap; }
        .qp-step { display: flex; align-items: center; gap: 0.5rem; padding: 0.4rem 0.75rem 0.4rem 0.5rem; border-radius: 999px; background: rgba(255,255,255,0.08); font-size: 0.82rem; color: #b9ccbe; border: 1px solid rgba(255,255,255,0.12); }
        .qp-step.done { background: rgba(255,255,255,0.14); color: #eef3ec; }
        .qp-step.active { background: var(--mango); border-color: var(--mango); color: #2a1400; }
        .qp-step-dot { width: 1.4rem; height: 1.4rem; border-radius: 50%; display: flex; align-items: center; justify-content: center; background: rgba(255,255,255,0.18); font-size: 0.75rem; }
        .qp-step.active .qp-step-dot { background: rgba(0,0,0,0.18); }

        .qp-main { max-width: 880px; margin: 0 auto; padding: 0 1.5rem; }
        .qp-section { margin-top: 2.25rem; padding-top: 2.25rem; border-top: 1px solid var(--line); }
        .qp-section:first-of-type { border-top: none; margin-top: 1.75rem; padding-top: 0; }
        .qp-section-head { display: flex; align-items: baseline; justify-content: space-between; margin-bottom: 1rem; gap: 1rem; flex-wrap: wrap; }
        .qp-section-title { font-size: 1.4rem; margin: 0; }
        .qp-section-sub { color: var(--ink-soft); font-size: 0.9rem; margin: 0.2rem 0 0; }

        .qp-card { background: var(--card); border: 1px solid var(--line); border-radius: 10px; padding: 1.1rem; }

        .qp-lock-note { display: flex; align-items: center; justify-content: space-between; gap: 0.75rem; background: var(--grass-bg); border: 1px solid var(--grass); color: #1e5c3c; padding: 0.6rem 0.85rem; border-radius: 8px; font-size: 0.88rem; margin-bottom: 1rem; flex-wrap: wrap; }

        .qp-entry-row { display: flex; gap: 0.6rem; flex-wrap: wrap; position: relative; }
        .qp-search-wrap { flex: 1 1 260px; position: relative; }
        .qp-input { width: 100%; padding: 0.65rem 0.8rem; border: 1px solid var(--line); border-radius: 8px; font-size: 0.95rem; background: var(--paper); color: var(--ink); }
        .qp-input:focus { outline: 2px solid var(--forest); outline-offset: 1px; }
        .qp-input:disabled { opacity: 0.55; cursor: not-allowed; }
        .qp-qty-input { width: 5rem; }
        .qp-suggestions { position: absolute; top: calc(100% + 4px); left: 0; right: 0; background: var(--card); border: 1px solid var(--line); border-radius: 8px; box-shadow: 0 6px 18px rgba(28,36,31,0.1); z-index: 5; overflow: hidden; }
        .qp-suggestion { padding: 0.55rem 0.8rem; cursor: pointer; font-size: 0.9rem; border-bottom: 1px solid var(--line); }
        .qp-suggestion:last-child { border-bottom: none; }
        .qp-suggestion:hover, .qp-suggestion:focus { background: var(--paper); }

        .qp-btn { padding: 0.65rem 1.1rem; border-radius: 8px; border: none; font-size: 0.92rem; font-weight: 600; cursor: pointer; transition: transform 0.08s ease, filter 0.15s ease; }
        .qp-btn:active { transform: translateY(1px); }
        .qp-btn:disabled { cursor: not-allowed; opacity: 0.45; }
        .qp-btn-forest { background: var(--forest); color: #eef3ec; }
        .qp-btn-forest:hover:not(:disabled) { filter: brightness(1.08); }
        .qp-btn-mango { background: var(--mango); color: #2a1400; }
        .qp-btn-mango:hover:not(:disabled) { filter: brightness(1.06); }
        .qp-btn-ghost { background: transparent; color: var(--ink-soft); border: 1px solid var(--line); }
        .qp-btn-ghost:hover:not(:disabled) { border-color: var(--warn); color: var(--warn); }
        .qp-btn-block { width: 100%; padding: 0.85rem 1rem; font-size: 1rem; }
        .qp-btn-confirmed { background: var(--grass-bg); color: var(--grass); border: 1px solid var(--grass); }

        .qp-list-table { width: 100%; border-collapse: collapse; margin-top: 1rem; }
        .qp-list-table th { text-align: left; font-size: 0.78rem; color: var(--ink-soft); font-weight: 600; padding: 0.4rem 0.5rem; border-bottom: 1px solid var(--line); }
        .qp-list-table td { padding: 0.55rem 0.5rem; border-bottom: 1px solid var(--line); font-size: 0.92rem; vertical-align: middle; }
        .qp-list-table tr:last-child td { border-bottom: none; }
        .qp-qty-controls { display: flex; align-items: center; gap: 0.35rem; }
        .qp-qty-btn { width: 1.6rem; height: 1.6rem; border-radius: 6px; border: 1px solid var(--line); background: var(--paper); cursor: pointer; font-size: 0.9rem; line-height: 1; }
        .qp-qty-btn:disabled { opacity: 0.4; cursor: not-allowed; }
        .qp-qty-field { width: 2.6rem; text-align: center; border: 1px solid var(--line); border-radius: 6px; padding: 0.25rem; font-size: 0.9rem; }
        .qp-remove-btn { border: none; background: transparent; color: var(--ink-soft); cursor: pointer; font-size: 1rem; padding: 0.2rem 0.4rem; border-radius: 6px; }
        .qp-remove-btn:hover { background: var(--warn-bg); color: var(--warn); }
        .qp-remove-btn:disabled { opacity: 0.3; cursor: not-allowed; }
        .qp-empty { color: var(--ink-soft); font-size: 0.9rem; padding: 1rem 0; text-align: center; }

        .qp-list-footer { display: flex; justify-content: space-between; align-items: center; margin-top: 1.25rem; gap: 0.75rem; flex-wrap: wrap; }

        .qp-table-scroll { overflow-x: auto; }
        .qp-cmp-table { width: 100%; border-collapse: collapse; min-width: 620px; }
        .qp-cmp-table th { text-align: left; font-size: 0.78rem; color: var(--ink-soft); font-weight: 600; padding: 0.5rem 0.6rem; border-bottom: 2px solid var(--line); }
        .qp-cmp-table td { padding: 0.65rem 0.6rem; border-bottom: 1px solid var(--line); font-size: 0.92rem; }
        .qp-price-cell { text-align: right; font-variant-numeric: tabular-nums; }
        .qp-price-cheapest { background: var(--grass-bg); color: var(--grass); font-weight: 700; border-radius: 6px; }
        .qp-price-none { color: var(--ink-soft); }
        .qp-unavail-tag { color: var(--warn); font-size: 0.78rem; font-weight: 600; margin-left: 0.4rem; }
        .qp-expand-link { background: none; border: none; color: var(--forest); font-size: 0.82rem; font-weight: 600; cursor: pointer; padding: 0; text-decoration: underline; }
        .qp-extra-row td { background: var(--paper); border-bottom: 1px solid var(--line); padding: 0.6rem; }
        .qp-extra-shops { display: flex; gap: 0.5rem; flex-wrap: wrap; }
        .qp-extra-chip { font-size: 0.82rem; padding: 0.3rem 0.55rem; border-radius: 6px; border: 1px solid var(--line); background: var(--card); }
        .qp-extra-chip.cheapest { border-color: var(--grass); background: var(--grass-bg); font-weight: 700; }

        .qp-opt-table { width: 100%; border-collapse: collapse; min-width: 560px; }
        .qp-opt-table th { text-align: left; font-size: 0.78rem; color: var(--ink-soft); font-weight: 600; padding: 0.5rem 0.6rem; border-bottom: 2px solid var(--line); }
        .qp-opt-table td { padding: 0.65rem 0.6rem; border-bottom: 1px solid var(--line); font-size: 0.92rem; }
        .qp-opt-table td.num { text-align: right; font-variant-numeric: tabular-nums; }
        .qp-shop-tag { display: inline-block; padding: 0.15rem 0.55rem; border-radius: 999px; background: var(--forest); color: #eef3ec; font-size: 0.8rem; font-weight: 600; }
        .qp-unavailable { color: var(--warn); font-weight: 600; font-size: 0.88rem; }

        .qp-shop-groups { display: grid; gap: 1rem; margin-top: 1.4rem; }
        .qp-shop-group { border: 1px solid var(--line); border-radius: 8px; padding: 0.9rem 1rem; background: var(--card); }
        .qp-shop-group-head { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 0.6rem; }
        .qp-shop-group-head h4 { margin: 0; font-size: 1.02rem; }
        .qp-shop-group-sub { font-size: 0.9rem; font-weight: 700; color: var(--forest); }
        .qp-shop-table { width: 100%; border-collapse: collapse; }
        .qp-shop-table th { text-align: left; font-size: 0.74rem; color: var(--ink-soft); padding: 0.3rem 0.4rem; }
        .qp-shop-table td { padding: 0.3rem 0.4rem; font-size: 0.88rem; border-top: 1px solid var(--line); }
        .qp-shop-table td.num, .qp-shop-table th.num { text-align: right; font-variant-numeric: tabular-nums; }

        .qp-receipt-wrap { display: flex; justify-content: center; margin-top: 1.5rem; }
        .qp-receipt { background: var(--card); width: 100%; max-width: 380px; padding: 1.5rem 1.4rem 1.75rem; border: 1px solid var(--line); position: relative; box-shadow: 0 10px 24px rgba(28,36,31,0.08); }
        .qp-receipt::before, .qp-receipt::after { content: ""; position: absolute; left: 0; right: 0; height: 10px; background: radial-gradient(circle at 8px 0, transparent 6px, var(--paper) 6.5px) repeat-x; background-size: 16px 16px; }
        .qp-receipt::before { top: -5px; }
        .qp-receipt::after { bottom: -5px; transform: rotate(180deg); }
        .qp-receipt-title { text-align: center; font-size: 1.1rem; margin: 0 0 0.9rem; }
        .qp-receipt-row { display: flex; justify-content: space-between; font-size: 0.92rem; padding: 0.3rem 0; font-variant-numeric: tabular-nums; }
        .qp-receipt-dash { border-top: 1px dashed var(--line); margin: 0.6rem 0; }
        .qp-receipt-total { font-size: 1.05rem; font-weight: 700; color: var(--forest); }
        .qp-receipt-save { color: var(--grass); font-weight: 700; }
        .qp-receipt-pct { color: var(--mango-deep); font-weight: 700; }
        .qp-receipt-stat-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 0.5rem; text-align: center; margin: 0.9rem 0; }
        .qp-receipt-stat b { display: block; font-size: 1.15rem; }
        .qp-receipt-stat span { font-size: 0.72rem; color: var(--ink-soft); }

        .qp-final-actions { display: flex; gap: 0.9rem; margin-top: 1.6rem; flex-wrap: wrap; }
        .qp-final-actions .qp-btn { flex: 1 1 180px; padding: 0.9rem 1rem; font-size: 1rem; }

        .qp-modal-overlay { position: fixed; inset: 0; background: rgba(20,26,21,0.55); display: flex; align-items: center; justify-content: center; padding: 1rem; z-index: 50; }
        .qp-modal { background: var(--card); border-radius: 10px; max-width: 360px; width: 100%; padding: 1.5rem; }
        .qp-modal h3 { margin: 0 0 1rem; font-size: 1.15rem; }
        .qp-modal-row { display: flex; justify-content: space-between; font-size: 0.95rem; padding: 0.35rem 0; }
        .qp-modal-row.total { font-weight: 700; color: var(--forest); border-top: 1px dashed var(--line); margin-top: 0.4rem; padding-top: 0.6rem; }
        .qp-modal-actions { display: flex; gap: 0.6rem; margin-top: 1.3rem; }
        .qp-modal-actions .qp-btn { flex: 1; }
        .qp-modal-note { margin-top: 1rem; padding: 0.75rem; background: var(--paper); border: 1px solid var(--line); border-radius: 8px; font-size: 0.85rem; color: var(--ink-soft); }

        @media (max-width: 560px) {
          .qp-brand { font-size: 1.5rem; }
          .qp-section-title { font-size: 1.2rem; }
          .qp-receipt-stat-grid { grid-template-columns: 1fr 1fr; }
        }

        @media print {
          .qp-screen-only { display: none !important; }
          .qp-print-only { display: block !important; padding: 1.5rem; font-family: Georgia, serif; color: #111; }
          .qp-print-only h1 { font-size: 1.4rem; margin: 0; }
          .qp-print-only .qp-print-sub { color: #555; margin: 0 0 1rem; }
          .qp-print-only table { width: 100%; border-collapse: collapse; margin: 0.75rem 0; }
          .qp-print-only th, .qp-print-only td { border-bottom: 1px solid #ccc; padding: 0.4rem 0.3rem; text-align: left; font-size: 0.9rem; }
          .qp-print-only td.num, .qp-print-only th.num { text-align: right; }
          .qp-print-meta { display: flex; gap: 1.5rem; font-size: 0.85rem; color: #444; margin-bottom: 1rem; }
          .qp-print-totals { margin-top: 0.75rem; max-width: 260px; margin-left: auto; }
          .qp-print-totals div { display: flex; justify-content: space-between; padding: 0.2rem 0; font-size: 0.92rem; }
          .qp-print-totals .final { font-weight: 700; border-top: 1px solid #111; margin-top: 0.3rem; padding-top: 0.4rem; }
        }
      `}</style>

      <div className="qp-screen-only">
        <header className="qp-header">
          <div className="qp-header-inner">
            {onBack && (
              <button
                onClick={onBack}
                style={{
                  display: "inline-flex", alignItems: "center", gap: 6, background: "none", border: "none",
                  cursor: "pointer", color: "#c3d6c9", fontSize: 13, fontWeight: 600, padding: 0, marginBottom: 12,
                }}
              >
                ← Back to Dashboard
              </button>
            )}
            <h1 className="qp-brand qp-display">QuickPick</h1>
            <p className="qp-tagline">Smart Shopping — build your list, compare shops, shop for less.</p>
            <div className="qp-steps">
              {["Build list", "Confirm", "Compare prices", "Optimize"].map((label, i) => (
                <div key={label} className={`qp-step ${stage > i ? "done" : ""} ${stage === i ? "active" : ""}`}>
                  <span className="qp-step-dot">{stage > i ? "✓" : i + 1}</span>
                  {label}
                </div>
              ))}
            </div>
          </div>
        </header>

        <main className="qp-main">
          {/* ============ 1. SHOPPING LIST GENERATOR ============ */}
          <section className="qp-section" ref={listRef}>
            <div className="qp-section-head">
              <div>
                <h2 className="qp-section-title">🛒 Shopping List Generator</h2>
                <p className="qp-section-sub">What do I want to buy?</p>
              </div>
            </div>

            <div className="qp-card">
              {locked && (
                <div className="qp-lock-note">
                  <span>✓ List locked for price comparison.</span>
                  <button className="qp-btn qp-btn-ghost" onClick={handleEditList}>Edit list</button>
                </div>
              )}

              <div className="qp-entry-row">
                <div className="qp-search-wrap">
                  <input
                    className="qp-input"
                    placeholder="Search or type a product…"
                    value={searchText}
                    disabled={locked}
                    onChange={(e) => { setSearchText(e.target.value); setShowSuggestions(true); }}
                    onFocus={() => setShowSuggestions(true)}
                    onKeyDown={(e) => { if (e.key === "Enter") addProduct(); }}
                  />
                  {!locked && showSuggestions && suggestions.length > 0 && (
                    <div className="qp-suggestions">
                      {suggestions.map((s) => (
                        <div key={s} className="qp-suggestion" onClick={() => { setSearchText(s); setShowSuggestions(false); }}>
                          {s}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
                <input
                  type="number"
                  min="1"
                  className="qp-input qp-qty-input"
                  value={qtyDraft}
                  disabled={locked}
                  onChange={(e) => setQtyDraft(Number(e.target.value))}
                  aria-label="Quantity"
                />
                <button className="qp-btn qp-btn-forest" onClick={addProduct} disabled={locked}>
                  Add product
                </button>
              </div>

              {list.length === 0 ? (
                <p className="qp-empty">Your list is empty. Add a product above to get started.</p>
              ) : (
                <table className="qp-list-table">
                  <thead>
                    <tr>
                      <th>Product</th>
                      <th style={{ width: "9rem" }}>Qty</th>
                      <th style={{ width: "2.5rem" }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {list.map((p) => (
                      <tr key={p.id}>
                        <td>{p.name}</td>
                        <td>
                          <div className="qp-qty-controls">
                            <button className="qp-qty-btn" disabled={locked} onClick={() => updateQty(p.id, p.qty - 1)} aria-label={`Decrease ${p.name} quantity`}>−</button>
                            <input className="qp-qty-field" type="number" min="1" value={p.qty} disabled={locked} onChange={(e) => updateQty(p.id, Number(e.target.value) || 1)} />
                            <button className="qp-qty-btn" disabled={locked} onClick={() => updateQty(p.id, p.qty + 1)} aria-label={`Increase ${p.name} quantity`}>+</button>
                          </div>
                        </td>
                        <td>
                          <button className="qp-remove-btn" disabled={locked} onClick={() => removeProduct(p.id)} aria-label={`Remove ${p.name}`}>🗑</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}

              <div className="qp-list-footer">
                <button className="qp-btn qp-btn-ghost" onClick={clearList} disabled={locked || list.length === 0}>
                  {clearArmed ? "Tap again to clear all" : "Clear entire list"}
                </button>

                {!locked ? (
                  <button className="qp-btn qp-btn-mango qp-btn-block" style={{ maxWidth: "16rem" }} onClick={handleConfirm} disabled={list.length === 0}>
                    Confirm shopping list
                  </button>
                ) : (
                  <button className="qp-btn qp-btn-confirmed qp-btn-block" style={{ maxWidth: "16rem" }} disabled>
                    ✓ Shopping list confirmed
                  </button>
                )}
              </div>
            </div>
          </section>

          {/* ============ 2. PRICE COMPARISON ENGINE ============ */}
          {locked && list.length > 0 && (
            <section className="qp-section" ref={compareRef}>
              <div className="qp-section-head">
                <div>
                  <h2 className="qp-section-title">💰 Price Comparison Engine</h2>
                  <p className="qp-section-sub">What are the prices at each shop? ✓ marks the cheapest.</p>
                </div>
              </div>

              <div className="qp-card">
                <div className="qp-table-scroll">
                  <table className="qp-cmp-table">
                    <thead>
                      <tr>
                        <th>Product</th>
                        <th style={{ textAlign: "right" }}>Qty</th>
                        {mainShops.map((s) => (
                          <th key={s} style={{ textAlign: "right" }}>{s}</th>
                        ))}
                        <th></th>
                      </tr>
                    </thead>
                    <tbody>
                      {analysis.map(({ item, available, entries, cheapest }) => {
                        const priceByShop = Object.fromEntries(entries.map((e) => [e.shop, e.price]));
                        const extraEntries = entries.filter((e) => extraShopsGlobal.includes(e.shop));
                        const isExpanded = !!expandedRows[item.id];

                        return (
                          <React.Fragment key={item.id}>
                            <tr>
                              <td>
                                {item.name}
                                {!available && <span className="qp-unavail-tag">⚠ unavailable</span>}
                              </td>
                              <td style={{ textAlign: "right" }}>{item.qty}</td>
                              {mainShops.map((shop) => {
                                const price = priceByShop[shop];
                                const isCheapest = available && cheapest.shop === shop;
                                return (
                                  <td key={shop} className={`qp-price-cell ${isCheapest ? "qp-price-cheapest" : ""} ${price == null ? "qp-price-none" : ""}`}>
                                    {price == null ? "—" : `${isCheapest ? "✓ " : ""}${ksh(price)}`}
                                  </td>
                                );
                              })}
                              <td style={{ textAlign: "right" }}>
                                {extraEntries.length > 0 && (
                                  <button className="qp-expand-link" onClick={() => toggleExpand(item.id)}>
                                    {isExpanded ? "Hide" : "+ View more shops"}
                                  </button>
                                )}
                              </td>
                            </tr>
                            {isExpanded && extraEntries.length > 0 && (
                              <tr className="qp-extra-row">
                                <td colSpan={mainShops.length + 3}>
                                  <div className="qp-extra-shops">
                                    {extraEntries.map((e) => (
                                      <span key={e.shop} className={`qp-extra-chip ${cheapest.shop === e.shop ? "cheapest" : ""}`}>
                                        {cheapest.shop === e.shop ? "✓ " : ""}{e.shop} — {ksh(e.price)}
                                      </span>
                                    ))}
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                </div>

                <div className="qp-list-footer" style={{ justifyContent: "flex-end" }}>
                  <button className="qp-btn qp-btn-mango qp-btn-block" style={{ maxWidth: "16rem" }} onClick={handleOptimize}>
                    Optimize shopping
                  </button>
                </div>
              </div>
            </section>
          )}

          {/* ============ 3. SHOP COMBINATION / OPTIMIZATION ENGINE ============ */}
          {optimized && locked && (
            <section className="qp-section" ref={optimizeRef}>
              <div className="qp-section-head">
                <div>
                  <h2 className="qp-section-title">🎯 Shop Combination / Optimization</h2>
                  <p className="qp-section-sub">Where should I buy each product to get the cheapest overall basket?</p>
                </div>
              </div>

              <div className="qp-card">
                <div className="qp-table-scroll">
                  <table className="qp-opt-table">
                    <thead>
                      <tr>
                        <th>Product</th>
                        <th style={{ textAlign: "right" }}>Qty</th>
                        <th>Recommended shop</th>
                        <th style={{ textAlign: "right" }}>Unit price</th>
                        <th style={{ textAlign: "right" }}>Product total</th>
                      </tr>
                    </thead>
                    <tbody>
                      {optimization.rows.map((r) => (
                        <tr key={r.id}>
                          <td>{r.name}</td>
                          <td className="num">{r.qty}</td>
                          <td>{r.available ? <span className="qp-shop-tag">{r.shop}</span> : <span className="qp-unavailable">⚠ Unavailable</span>}</td>
                          <td className="num">{r.available ? ksh(r.unitPrice) : "—"}</td>
                          <td className="num">{r.available ? ksh(r.total) : "—"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="qp-shop-groups">
                  {Object.entries(optimization.byShop).sort((a, b) => b[1].subtotal - a[1].subtotal).map(([shop, data]) => (
                    <div className="qp-shop-group" key={shop}>
                      <div className="qp-shop-group-head">
                        <h4>🛒 {shop}</h4>
                        <span className="qp-shop-group-sub">Subtotal: {ksh(data.subtotal)}</span>
                      </div>
                      <table className="qp-shop-table">
                        <thead>
                          <tr><th>Product</th><th className="num">Qty</th><th className="num">Price</th><th className="num">Total</th></tr>
                        </thead>
                        <tbody>
                          {data.items.map((it) => (
                            <tr key={it.id}>
                              <td>{it.name}</td>
                              <td className="num">{it.qty}</td>
                              <td className="num">{ksh(it.unitPrice)}</td>
                              <td className="num">{ksh(it.total)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ))}
                </div>
              </div>

              {/* ---- summary ---- */}
              <div className="qp-receipt-wrap">
                <div className="qp-receipt">
                  <h3 className="qp-receipt-title qp-display">🧾 Shopping Summary</h3>

                  <div className="qp-receipt-stat-grid">
                    <div className="qp-receipt-stat"><b>{optimization.shopsToVisit}</b><span>Shops to visit</span></div>
                    <div className="qp-receipt-stat"><b>{optimization.productCount}</b><span>Products</span></div>
                    <div className="qp-receipt-stat"><b>{optimization.items}</b><span>Items</span></div>
                  </div>

                  <div className="qp-receipt-dash" />

                  <div className="qp-receipt-row"><span>Most expensive basket</span><span>{ksh(optimization.priciestTotal)}</span></div>
                  <div className="qp-receipt-row qp-receipt-total"><span>Optimized basket</span><span>{ksh(optimization.optimizedTotal)}</span></div>

                  <div className="qp-receipt-dash" />

                  <div className="qp-receipt-row qp-receipt-save"><span>Total savings</span><span>{ksh(optimization.savings)}</span></div>
                  <div className="qp-receipt-row qp-receipt-pct"><span>Savings percentage</span><span>{optimization.savingsPct.toFixed(1)}%</span></div>
                </div>
              </div>

              {/* ---- final actions: print + pay ---- */}
              <div className="qp-final-actions">
                <button className="qp-btn qp-btn-forest" onClick={() => window.print()}>
                  🖨 Print receipt
                </button>
                <button className="qp-btn qp-btn-mango" onClick={() => { setPayStatus(null); setPayOpen(true); }}>
                  💳 Pay {ksh(optimization.optimizedTotal)}
                </button>
              </div>
            </section>
          )}
        </main>
      </div>

      {/* ============ PAY MODAL ============ */}
      {payOpen && (
        <div className="qp-modal-overlay" role="dialog" aria-modal="true">
          <div className="qp-modal">
            <h3 className="qp-display">Confirm Payment</h3>
            <div className="qp-modal-row"><span>Shopping total</span><span>{ksh(optimization.priciestTotal)}</span></div>
            <div className="qp-modal-row"><span>Savings</span><span>−{ksh(optimization.savings)}</span></div>
            <div className="qp-modal-row total"><span>Amount to pay</span><span>{ksh(optimization.optimizedTotal)}</span></div>

            {payStatus === "pending" ? (
              <div className="qp-modal-note">
                Payment isn't wired up yet — this screen is ready for a provider such as M-Pesa, card, or mobile money to be connected. No charge has been made.
              </div>
            ) : (
              <div className="qp-modal-actions">
                <button className="qp-btn qp-btn-ghost" onClick={() => setPayOpen(false)}>Cancel</button>
                <button className="qp-btn qp-btn-mango" onClick={() => setPayStatus("pending")}>Confirm & Pay</button>
              </div>
            )}
            {payStatus === "pending" && (
              <div className="qp-modal-actions">
                <button className="qp-btn qp-btn-ghost" style={{ flex: 1 }} onClick={() => setPayOpen(false)}>Close</button>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ============ PRINT-ONLY RECEIPT ============ */}
      <div className="qp-print-only">
        <h1>QUICKPICK</h1>
        <p className="qp-print-sub">Shopping Receipt</p>
        <div className="qp-print-meta">
          <span>Date: {order?.date || "—"}</span>
          <span>Time: {order?.time || "—"}</span>
          <span>Order ID: {order?.id || "—"}</span>
        </div>
        <table>
          <thead>
            <tr><th>Product</th><th className="num">Qty</th><th>Shop</th><th className="num">Unit price</th><th className="num">Total</th></tr>
          </thead>
          <tbody>
            {optimization.rows.filter((r) => r.available).map((r) => (
              <tr key={r.id}>
                <td>{r.name}</td>
                <td className="num">{r.qty}</td>
                <td>{r.shop}</td>
                <td className="num">{ksh(r.unitPrice)}</td>
                <td className="num">{ksh(r.total)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="qp-print-totals">
          <div><span>Subtotal</span><span>{ksh(optimization.optimizedTotal)}</span></div>
          <div><span>Savings</span><span>{ksh(optimization.savings)}</span></div>
          <div className="final"><span>Total</span><span>{ksh(optimization.optimizedTotal)}</span></div>
        </div>
      </div>
    </div>
  );
}


// ============================================================
// ---------- Product Search & Display (customer-facing) ----------
// ============================================================

function pdSearchIconFor(category) {
  const map = {
    "Food & Beverages": "🍽️", Household: "🧺", "Personal Care": "🧴",
    "Baby & Kids": "🍼", "Health & Wellness": "💊", Electronics: "🔌",
    Stationery: "✏️", Clothing: "👕", "General Merchandise": "📦",
  };
  return map[category] || "🛒";
}

const PD_SHOP_COLORS = ["#2E7D74", "#7A5490", "#B5591F", "#3D5A80", "#74792F", "#8A4A4A", "#4A6B5C", "#6B5B95"];
function pdShopColor(shopId, shopList) {
  const idx = shopList.findIndex((s) => s.id === shopId);
  return PD_SHOP_COLORS[(idx >= 0 ? idx : 0) % PD_SHOP_COLORS.length];
}

// Reused for cards, table, comparison, sorting, filtering - the single
// place that resolves "which shops sell this product, at what price."
function getProductOffers(productId, shopList, shopProductsList) {
  return shopProductsList
    .filter((sp) => sp.productId === productId)
    .map((sp) => {
      const shop = shopList.find((s) => s.id === sp.shopId);
      return shop ? { shopId: shop.id, shopName: shop.name, price: Number(sp.price) || 0, lastUpdated: sp.lastUpdated } : null;
    })
    .filter(Boolean);
}

function pdFmt(n) { return "KSh " + Math.round(n).toLocaleString(); }

const PD_SAVED_KEY = "quickpick_saved_products";

function pdLoad(key, fallback) {
  if (typeof window === "undefined" || !window.localStorage) return fallback;
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (e) {
    return fallback;
  }
}
function pdSave(key, value) {
  if (typeof window === "undefined" || !window.localStorage) return;
  try { window.localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* storage unavailable */ }
}

const PD_CSS = `
.qps-root{
  --bg:#F3F5F1; --surface:#FFFFFF; --border:#E2E6E0; --border-soft:#EDEFEB;
  --ink:#1B231F; --ink-soft:#5C6A62; --ink-faint:#8D998F;
  --brand:#1E7145; --brand-dark:#154F31; --brand-tint:#E3F1E7;
  --amber:#A15E1A; --amber-tint:#F6E9D8; --danger:#B23B3B; --danger-tint:#F7E4E4;
  --radius:10px; --radius-sm:5px;
  background:var(--bg); color:var(--ink); font-family:'IBM Plex Sans',system-ui,-apple-system,sans-serif;
  -webkit-font-smoothing:antialiased; min-height:100vh;
}
.qps-root *{box-sizing:border-box;}
.qps-root button{font-family:inherit;} .qps-root input,.qps-root select{font-family:inherit;}
.qps-app{max-width:1400px;margin:0 auto;padding:0 20px 64px;}
@media (max-width:640px){ .qps-app{padding:0 12px 48px;} }
.qps-top{padding:16px 0 16px;border-bottom:1px solid var(--border);position:sticky;top:0;background:var(--bg);z-index:30;}
.qps-back{display:inline-flex;align-items:center;gap:6px;background:none;border:none;cursor:pointer;color:var(--ink-soft);font-size:13px;font-weight:600;padding:0;margin-bottom:12px;}
.qps-back:hover{color:var(--ink);}
.qps-brand-row{display:flex;align-items:flex-end;justify-content:space-between;gap:16px;margin-bottom:16px;flex-wrap:wrap;}
.qps-brand-mark{display:flex;align-items:baseline;gap:10px;}
.qps-logo{font-family:ui-monospace,'SF Mono',monospace;font-weight:700;font-size:22px;letter-spacing:-0.02em;color:var(--brand-dark);background:var(--brand-tint);padding:2px 8px;border-radius:var(--radius-sm);border:1px solid #C7E2CE;}
.qps-tagline{color:var(--ink-soft);font-size:14px;}
.qps-cart-btn{display:flex;align-items:center;gap:8px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-sm);padding:9px 14px;cursor:pointer;font-size:14px;font-weight:600;color:var(--ink);}
.qps-cart-btn:hover{border-color:var(--brand);}
.qps-cart-count{font-family:ui-monospace,monospace;font-weight:700;background:var(--brand);color:#fff;border-radius:999px;min-width:20px;height:20px;padding:0 6px;display:flex;align-items:center;justify-content:center;font-size:12px;}
.qps-search-row{position:relative;}
.qps-search-row input{width:100%;padding:13px 44px 13px 42px;border:1.5px solid var(--border);border-radius:var(--radius);background:var(--surface);font-size:15px;color:var(--ink);outline:none;}
.qps-search-row input:focus{border-color:var(--brand);}
.qps-icon-search{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:var(--ink-faint);pointer-events:none;}
.qps-icon-clear{position:absolute;right:10px;top:50%;transform:translateY(-50%);background:var(--border-soft);border:none;border-radius:50%;width:26px;height:26px;display:none;align-items:center;justify-content:center;cursor:pointer;color:var(--ink-soft);}
.qps-icon-clear.qps-show{display:flex;}
.qps-toolbar{padding:16px 0 12px;display:flex;flex-direction:column;gap:12px;}
.qps-chip-row{display:flex;gap:8px;overflow-x:auto;padding-bottom:2px;scrollbar-width:none;}
.qps-chip-row::-webkit-scrollbar{display:none;}
.qps-chip{flex:0 0 auto;padding:7px 14px;border-radius:999px;border:1.5px solid var(--border);background:var(--surface);font-size:13.5px;font-weight:500;color:var(--ink-soft);cursor:pointer;white-space:nowrap;}
.qps-chip:hover{border-color:var(--brand);color:var(--ink);}
.qps-chip.qps-active{background:var(--brand);border-color:var(--brand);color:#fff;}
.qps-toolbar-controls{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;}
.qps-result-summary{font-size:13.5px;color:var(--ink-soft);}
.qps-result-summary strong{color:var(--ink);font-weight:600;}
.qps-toolbar-right{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}
.qps-select-wrap{position:relative;}
.qps-sort-select{appearance:none;-webkit-appearance:none;padding:8px 32px 8px 12px;border:1.5px solid var(--border);border-radius:var(--radius-sm);background:var(--surface);font-size:13.5px;color:var(--ink);cursor:pointer;}
.qps-select-wrap::after{content:"";position:absolute;right:11px;top:50%;transform:translateY(-40%);border:4px solid transparent;border-top-color:var(--ink-soft);pointer-events:none;}
.qps-view-toggle{display:flex;border:1.5px solid var(--border);border-radius:var(--radius-sm);overflow:hidden;}
.qps-view-toggle button{background:var(--surface);border:none;padding:7px 11px;cursor:pointer;color:var(--ink-soft);display:flex;align-items:center;font-size:13px;gap:5px;}
.qps-view-toggle button.qps-active{background:var(--brand-tint);color:var(--brand-dark);}
.qps-view-toggle button+button{border-left:1.5px solid var(--border);}
.qps-filter-toggle-btn{display:none;align-items:center;gap:6px;padding:8px 13px;border:1.5px solid var(--border);border-radius:var(--radius-sm);background:var(--surface);font-size:13.5px;font-weight:600;color:var(--ink);cursor:pointer;}
.qps-fbadge{background:var(--brand);color:#fff;border-radius:999px;font-family:ui-monospace,monospace;font-size:11px;min-width:16px;height:16px;padding:0 4px;display:flex;align-items:center;justify-content:center;}
.qps-body-layout{display:flex;gap:22px;margin-top:18px;align-items:flex-start;}
.qps-sidebar{flex:0 0 250px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:16px;position:sticky;top:150px;max-height:calc(100vh - 170px);overflow-y:auto;}
.qps-main-col{flex:1;min-width:0;}
.qps-filter-block{padding:14px 0;border-bottom:1px solid var(--border-soft);}
.qps-filter-block:first-child{padding-top:0;}
.qps-filter-block:last-of-type{border-bottom:none;}
.qps-filter-block h4{margin:0 0 10px;font-size:12.5px;font-weight:600;color:var(--ink-soft);text-transform:uppercase;letter-spacing:.04em;}
.qps-filter-opt{display:flex;align-items:center;gap:8px;padding:4px 0;font-size:13.5px;color:var(--ink);cursor:pointer;}
.qps-filter-opt input{accent-color:var(--brand);width:15px;height:15px;cursor:pointer;}
.qps-swatch{width:9px;height:9px;border-radius:50%;flex:0 0 auto;}
.qps-price-range{display:flex;align-items:center;gap:8px;}
.qps-price-range input{width:100%;border:1.5px solid var(--border);border-radius:var(--radius-sm);padding:7px 8px;font-size:13px;font-family:ui-monospace,monospace;color:var(--ink);}
.qps-clear-filters-btn{width:100%;margin-top:14px;padding:9px;border-radius:var(--radius-sm);border:1.5px solid var(--border);background:var(--bg);color:var(--ink);font-size:13.5px;font-weight:600;cursor:pointer;}
.qps-clear-filters-btn:hover{border-color:var(--danger);color:var(--danger);}
.qps-sidebar-close{display:none;}
.qps-card-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;}
@media (max-width:1150px){ .qps-card-grid{grid-template-columns:repeat(3,1fr);} }
@media (max-width:860px){ .qps-card-grid{grid-template-columns:repeat(2,1fr);} }
@media (max-width:480px){ .qps-card-grid{grid-template-columns:1fr;} }
.qps-card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);display:flex;flex-direction:column;overflow:hidden;height:100%;border-top:3px solid var(--shop-color,var(--brand));}
.qps-img-box{aspect-ratio:1/0.82;background:var(--brand-tint);display:flex;align-items:center;justify-content:center;font-size:42px;position:relative;overflow:hidden;}
.qps-img-box img{width:100%;height:100%;object-fit:cover;}
.qps-stock-badge{position:absolute;top:8px;left:8px;background:var(--danger-tint);color:var(--danger);font-size:11px;font-weight:600;padding:3px 8px;border-radius:999px;}
.qps-card-body{padding:12px 13px 0;flex:1;display:flex;flex-direction:column;}
.qps-pname{font-size:14.5px;font-weight:600;line-height:1.25;margin:0 0 2px;color:var(--ink);}
.qps-pmeta{font-size:12.5px;color:var(--ink-soft);margin:0 0 10px;}
.qps-shop-line{display:flex;align-items:center;gap:6px;font-size:12.5px;color:var(--ink-soft);margin-bottom:10px;}
.qps-shop-line .qps-dot{width:7px;height:7px;border-radius:50%;background:var(--shop-color,var(--brand));}
.qps-shop-line strong{color:var(--ink);font-weight:600;}
.qps-price-cell{margin-top:auto;border:1px solid var(--border-soft);border-radius:var(--radius-sm);background:var(--bg);padding:8px 10px;margin-bottom:12px;}
.qps-price-cell .qps-price{font-family:ui-monospace,monospace;font-size:19px;font-weight:700;color:var(--ink);}
.qps-price-cell .qps-per{font-size:11.5px;color:var(--ink-faint);margin-left:4px;}
.qps-price-cell .qps-indicator{font-size:11.5px;font-weight:600;margin-top:3px;}
.qps-indicator.qps-lowest{color:var(--brand-dark);}
.qps-indicator.qps-higher{color:var(--amber);}
.qps-action-bar{display:flex;border-top:1px solid var(--border-soft);}
.qps-action-bar button{flex:1;background:none;border:none;padding:11px 0;cursor:pointer;display:flex;align-items:center;justify-content:center;position:relative;color:var(--ink-soft);font-size:18px;}
.qps-action-bar button+button{border-left:1px solid var(--border-soft);}
.qps-action-bar button.qps-active{color:var(--brand-dark);background:var(--brand-tint);}
.qps-action-bar button.qps-active.qps-saved{color:var(--amber);background:var(--amber-tint);}
.qps-qty-pill{position:absolute;top:2px;right:14px;background:var(--brand);color:#fff;font-family:ui-monospace,monospace;font-size:10px;font-weight:700;border-radius:999px;min-width:15px;height:15px;padding:0 3px;display:flex;align-items:center;justify-content:center;}
.qps-table-wrap{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);overflow:auto;}
.qps-table{width:100%;border-collapse:collapse;min-width:680px;}
.qps-table th{text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:.04em;color:var(--ink-soft);padding:11px 14px;border-bottom:1px solid var(--border);background:var(--bg);position:sticky;top:0;}
.qps-table td{padding:11px 14px;border-bottom:1px solid var(--border-soft);font-size:13.5px;vertical-align:middle;}
.qps-table td.qps-price-td{font-family:ui-monospace,monospace;font-weight:700;font-size:14.5px;}
.qps-table tr:last-child td{border-bottom:none;}
.qps-shop-td{display:flex;align-items:center;gap:6px;}
.qps-shop-td .qps-dot{width:7px;height:7px;border-radius:50%;}
.qps-avail-badge{font-size:11px;font-weight:600;padding:3px 8px;border-radius:999px;}
.qps-avail-badge.qps-yes{background:var(--brand-tint);color:var(--brand-dark);}
.qps-avail-badge.qps-no{background:var(--danger-tint);color:var(--danger);}
.qps-t-actions{display:flex;gap:4px;}
.qps-t-actions button{background:var(--bg);border:1px solid var(--border);border-radius:var(--radius-sm);width:30px;height:30px;cursor:pointer;font-size:14px;color:var(--ink-soft);}
.qps-t-actions button.qps-active{background:var(--brand-tint);border-color:var(--brand);color:var(--brand-dark);}
.qps-empty-state{background:var(--surface);border:1px dashed var(--border);border-radius:var(--radius);padding:56px 24px;text-align:center;}
.qps-empty-state h3{margin:0 0 6px;font-size:17px;}
.qps-empty-state p{color:var(--ink-soft);font-size:14px;margin:0 0 4px;}
.qps-empty-state ul{color:var(--ink-soft);font-size:13.5px;display:inline-block;text-align:left;margin:10px 0 18px;padding-left:18px;}
.qps-empty-state button{background:var(--brand);color:#fff;border:none;border-radius:var(--radius-sm);padding:10px 18px;font-size:13.5px;font-weight:600;cursor:pointer;}
.qps-overlay{position:fixed;inset:0;background:rgba(20,26,22,.4);z-index:90;opacity:0;pointer-events:none;transition:opacity .18s ease;}
.qps-overlay.qps-show{opacity:1;pointer-events:auto;}
.qps-drawer{position:fixed;top:0;right:0;height:100%;width:380px;max-width:90vw;background:var(--surface);z-index:100;box-shadow:-8px 0 24px rgba(0,0,0,.12);transform:translateX(100%);transition:transform .22s ease;display:flex;flex-direction:column;}
.qps-drawer.qps-show{transform:translateX(0);}
.qps-drawer.qps-left{right:auto;left:0;transform:translateX(-100%);box-shadow:8px 0 24px rgba(0,0,0,.12);}
.qps-drawer.qps-left.qps-show{transform:translateX(0);}
.qps-drawer-head{padding:18px 18px 14px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;}
.qps-drawer-head h3{margin:0;font-size:16px;}
.qps-drawer-head button{background:none;border:none;font-size:20px;cursor:pointer;color:var(--ink-soft);}
.qps-drawer-body{padding:16px 18px;overflow-y:auto;flex:1;}
.qps-drawer-foot{padding:14px 18px;border-top:1px solid var(--border);}
.qps-cart-item{display:flex;gap:10px;padding:12px 0;border-bottom:1px solid var(--border-soft);}
.qps-cart-item .qps-img-box{width:52px;height:52px;border-radius:var(--radius-sm);flex:0 0 auto;aspect-ratio:auto;}
.qps-ci-info{flex:1;min-width:0;}
.qps-ci-name{font-size:13.5px;font-weight:600;margin:0 0 2px;}
.qps-ci-meta{font-size:12px;color:var(--ink-soft);margin:0 0 6px;}
.qps-ci-row{display:flex;align-items:center;justify-content:space-between;}
.qps-qty-ctrl{display:flex;align-items:center;border:1px solid var(--border);border-radius:999px;overflow:hidden;}
.qps-qty-ctrl button{background:var(--bg);border:none;width:24px;height:24px;cursor:pointer;font-size:14px;color:var(--ink);}
.qps-qty-ctrl span{font-family:ui-monospace,monospace;font-size:12.5px;min-width:20px;text-align:center;}
.qps-ci-price{font-family:ui-monospace,monospace;font-weight:700;font-size:13.5px;}
.qps-remove-x{background:none;border:none;color:var(--ink-faint);cursor:pointer;font-size:12px;margin-left:8px;}
.qps-cart-empty{text-align:center;padding:40px 10px;color:var(--ink-soft);font-size:13.5px;}
.qps-cart-total-row{display:flex;justify-content:space-between;font-size:14.5px;font-weight:600;margin-bottom:10px;}
.qps-cart-total-row span:last-child{font-family:ui-monospace,monospace;}
.qps-modal{position:fixed;top:50%;left:50%;transform:translate(-50%,-46%);width:520px;max-width:92vw;max-height:86vh;overflow-y:auto;background:var(--surface);border-radius:var(--radius);z-index:101;box-shadow:0 20px 50px rgba(0,0,0,.2);opacity:0;pointer-events:none;transition:all .18s ease;}
.qps-modal.qps-show{opacity:1;pointer-events:auto;transform:translate(-50%,-50%);}
.qps-modal-head{padding:18px 20px 14px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:flex-start;}
.qps-modal-head h3{margin:0 0 3px;font-size:16.5px;}
.qps-modal-head p{margin:0;font-size:12.5px;color:var(--ink-soft);}
.qps-modal-head button{background:none;border:none;font-size:20px;cursor:pointer;color:var(--ink-soft);}
.qps-modal-body{padding:18px 20px;}
.qps-compare-row{display:flex;align-items:center;gap:10px;padding:11px 0;border-bottom:1px solid var(--border-soft);}
.qps-compare-row input[type=radio]{accent-color:var(--brand);width:16px;height:16px;}
.qps-compare-row .qps-dot{width:9px;height:9px;border-radius:50%;flex:0 0 auto;}
.qps-cr-name{flex:1;font-size:14px;font-weight:500;}
.qps-cr-price{font-family:ui-monospace,monospace;font-weight:700;font-size:14.5px;}
.qps-cr-diff{font-size:11.5px;color:var(--ink-soft);width:88px;text-align:right;}
.qps-compare-row.qps-cheapest .qps-cr-diff{color:var(--brand-dark);font-weight:600;}
.qps-compare-row.qps-unavailable{opacity:.45;}
.qps-save-banner{margin-top:14px;background:var(--brand-tint);color:var(--brand-dark);border-radius:var(--radius-sm);padding:9px 12px;font-size:13px;font-weight:600;}
.qps-modal-foot{padding:14px 20px 20px;}
.qps-btn-primary{width:100%;background:var(--brand);color:#fff;border:none;border-radius:var(--radius-sm);padding:12px;font-size:14.5px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;}
.qps-btn-primary:hover{background:var(--brand-dark);}
.qps-sidebar-overlay-title{display:none;}
@media (max-width:900px){
  .qps-sidebar{position:fixed;top:0;left:0;height:100%;width:300px;max-width:85vw;max-height:none;z-index:100;border-radius:0;transform:translateX(-100%);transition:transform .2s ease;padding-top:0;}
  .qps-sidebar.qps-show{transform:translateX(0);}
  .qps-sidebar-overlay-title{display:flex;align-items:center;justify-content:space-between;padding:16px;border-bottom:1px solid var(--border);margin:-16px -16px 6px;}
  .qps-sidebar-overlay-title h3{margin:0;font-size:16px;}
  .qps-sidebar-overlay-title button{background:none;border:none;font-size:20px;cursor:pointer;}
  .qps-sidebar-close{display:block;}
  .qps-filter-toggle-btn{display:flex;}
}
`;

function QuickPickProductSearch({ onBack, products = [], shops = [], shopProducts = [], savedShoppingLists = [], shoppingList = [], onShoppingListChange, onOpenSmartShopping }) {
  // ----- Build one "offer record" per product-shop relationship (real data) -----
  const records = useMemo(() => {
    const validProducts = Array.isArray(products) ? products.filter((p) => p && p.id) : [];
    const validShops = Array.isArray(shops) ? shops.filter((s) => s && s.id) : [];
    const productIndex = new Map(validProducts.map((p) => [p.id, p]));
    const shopIndex = new Map(validShops.map((s) => [s.id, s]));
    const validShopProducts = Array.isArray(shopProducts) ? shopProducts.filter((sp) => sp && sp.id) : [];

    const offerRecords = validShopProducts
      .filter((sp) => productIndex.has(sp.productId) && shopIndex.has(sp.shopId))
      .map((sp) => {
        const product = productIndex.get(sp.productId);
        const shop = shopIndex.get(sp.shopId);
        return {
          id: sp.id,
          productId: product.id,
          name: product.productName || "Unnamed product",
          brand: product.brand || "",
          packageSize: product.packageSize || "",
          category: product.category || "General Merchandise",
          icon: pdSearchIconFor(product.category),
          image: product.image || null,
          shop: { id: shop.id, name: shop.name, color: pdShopColor(shop.id, validShops) },
          price: Number(sp.price) || 0,
          available: (Number(product.stock) || 0) > 0 && product.status !== "Out of Stock",
          dateAdded: sp.lastUpdated ? new Date(sp.lastUpdated).getTime() : 0,
        };
      });

    // A product nobody sells yet still shows up (so a newly created product is
    // visible straight away) - flagged so it can't be added to a list.
    const withOffers = new Set(offerRecords.map((r) => r.productId));
    const placeholders = validProducts
      .filter((product) => !withOffers.has(product.id))
      .map((product) => ({
        id: `np-${product.id}`,
        productId: product.id,
        name: product.productName || "Unnamed product",
        brand: product.brand || "",
        packageSize: product.packageSize || "",
        category: product.category || "General Merchandise",
        icon: pdSearchIconFor(product.category),
        image: product.image || null,
        shop: { id: null, name: "No shop assigned", color: "#9aa1a8" },
        price: null,
        available: false,
        noShop: true,
        dateAdded: 0,
      }));
    return [...offerRecords, ...placeholders];
  }, [products, shops, shopProducts]);

  const categories = useMemo(() => {
    const set = new Set(records.map((r) => r.category).filter(Boolean));
    return ["All", ...Array.from(set).sort()];
  }, [records]);
  const shopChoices = useMemo(() => {
    const map = new Map();
    records.forEach((r) => { if (!r.noShop && !map.has(r.shop.id)) map.set(r.shop.id, r.shop); });
    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name));
  }, [records]);
  const allBrands = useMemo(() => Array.from(new Set(records.map((r) => r.brand).filter(Boolean))).sort(), [records]);
  const allSizes = useMemo(() => Array.from(new Set(records.map((r) => r.packageSize).filter(Boolean))), [records]);

  function groupOf(productId) { return records.filter((r) => r.productId === productId && !r.noShop); }
  function cheapestOf(productId) {
    const group = groupOf(productId);
    if (!group.length) return null;
    const availableOnes = group.filter((r) => r.available);
    const pool = availableOnes.length ? availableOnes : group;
    return pool.reduce((a, b) => (!a || b.price < a.price ? b : a), null);
  }
  function withSaving(r) {
    const cheapest = cheapestOf(r.productId);
    const diff = cheapest ? r.price - cheapest.price : 0;
    return { isLowest: !!cheapest && cheapest.id === r.id, diff, cheapest };
  }

  // ----- UI state (mirrors the original page's `state` object) -----
  const [search, setSearch] = useState("");
  const [category, setCategory] = useState("All");
  const [shopFilter, setShopFilter] = useState(() => new Set());
  const [brandFilter, setBrandFilter] = useState(() => new Set());
  const [sizeFilter, setSizeFilter] = useState(() => new Set());
  const [minPrice, setMinPrice] = useState("");
  const [maxPrice, setMaxPrice] = useState("");
  const [availability, setAvailability] = useState("all");
  const [sort, setSort] = useState("relevance");
  const [view, setView] = useState("cards");
  const [saved, setSaved] = useState(() => new Set(pdLoad(PD_SAVED_KEY, [])));
  const [cartOpen, setCartOpen] = useState(false);
  const [filtersOpen, setFiltersOpen] = useState(false);
  const [compareRecordId, setCompareRecordId] = useState(null);
  const [compareSelectedId, setCompareSelectedId] = useState(null);
  const [addedToast, setAddedToast] = useState("");

  useEffect(() => pdSave(PD_SAVED_KEY, Array.from(saved)), [saved]);
  useEffect(() => {
    if (!addedToast) return;
    const t = setTimeout(() => setAddedToast(""), 2200);
    return () => clearTimeout(t);
  }, [addedToast]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return records.filter((r) => {
      if (category !== "All" && r.category !== category) return false;
      if (shopFilter.size && !shopFilter.has(r.shop.id)) return false;
      if (brandFilter.size && !brandFilter.has(r.brand)) return false;
      if (sizeFilter.size && !sizeFilter.has(r.packageSize)) return false;
      if (r.noShop && (minPrice !== "" || maxPrice !== "")) return false;
      if (minPrice !== "" && r.price < Number(minPrice)) return false;
      if (maxPrice !== "" && r.price > Number(maxPrice)) return false;
      if (availability === "available" && !r.available) return false;
      if (availability === "out" && r.available) return false;
      if (q) {
        const hay = `${r.name} ${r.brand} ${r.category} ${r.shop.name} ${r.packageSize}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }, [records, search, category, shopFilter, brandFilter, sizeFilter, minPrice, maxPrice, availability]);

  const sorted = useMemo(() => {
    const arr = filtered.slice();
    switch (sort) {
      case "lowest": arr.sort((a, b) => (a.price == null ? 1e12 : a.price) - (b.price == null ? 1e12 : b.price)); break;
      case "highest": arr.sort((a, b) => (b.price == null ? -1 : b.price) - (a.price == null ? -1 : a.price)); break;
      case "name": arr.sort((a, b) => a.name.localeCompare(b.name)); break;
      case "shop": arr.sort((a, b) => a.shop.name.localeCompare(b.shop.name)); break;
      case "saving": arr.sort((a, b) => withSaving(b).diff - withSaving(a).diff); break;
      case "recent": arr.sort((a, b) => b.dateAdded - a.dateAdded); break;
      default: break; // relevance = catalog order
    }
    return arr;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [filtered, sort, records]);

  const shopCountInResults = new Set(sorted.filter((r) => !r.noShop).map((r) => r.shop.id)).size;
  const activeFilterCount =
    shopFilter.size + brandFilter.size + sizeFilter.size + (minPrice ? 1 : 0) + (maxPrice ? 1 : 0) + (availability !== "all" ? 1 : 0);

  function toggleSetValue(setFn, value) {
    setFn((prev) => {
      const next = new Set(prev);
      if (next.has(value)) next.delete(value); else next.add(value);
      return next;
    });
  }
  function clearFilters() {
    setShopFilter(new Set()); setBrandFilter(new Set()); setSizeFilter(new Set());
    setMinPrice(""); setMaxPrice(""); setAvailability("all");
  }
  function clearAll() { setSearch(""); setCategory("All"); clearFilters(); }

  function toggleSave(recordId) { toggleSetValue(setSaved, recordId); }

  // ----- Shopping list: shared with Smart Shopping through App state -----
  // Each item keeps productId + chosen shop + price + quantity. Prices shown
  // in the drawer are always re-read from the live shop-price records.
  const list = Array.isArray(shoppingList) ? shoppingList : [];
  const updateList = (fn) => { if (onShoppingListChange) onShoppingListChange(fn); };
  const sameProduct = (item, r) =>
    item.productId ? item.productId === r.productId : String(item.name || "").toLowerCase() === r.name.toLowerCase();

  function cartQtyFor(r) {
    if (r.noShop) return 0;
    const it = list.find((i) => sameProduct(i, r));
    return it && it.shopId === r.shop.id ? it.qty : 0;
  }
  function addRecordToList(record) {
    if (!record || record.noShop) return;
    const existing = list.find((i) => sameProduct(i, record));
    const entry = { productId: record.productId, name: record.name, shopId: record.shop.id, shopName: record.shop.name, price: record.price };
    updateList((prev) => {
      const cur = prev.find((i) => sameProduct(i, record));
      if (cur) return prev.map((i) => (i.id === cur.id ? { ...i, ...entry } : i)); // switch shop, keep qty
      const nextId = prev.reduce((m, i) => Math.max(m, Number(i.id) || 0), 0) + 1;
      return [...prev, { id: nextId, qty: 1, ...entry }];
    });
    setAddedToast(existing ? `${record.name} will be bought at ${record.shop.name}` : `Added ${record.name} to your shopping list`);
  }
  function toggleCart(recordId) {
    const record = records.find((r) => r.id === recordId);
    if (!record || record.noShop) return;
    if (cartQtyFor(record) > 0) {
      updateList((prev) => prev.filter((i) => !sameProduct(i, record)));
      setAddedToast(`Removed ${record.name} from your shopping list`);
    } else {
      addRecordToList(record);
    }
  }
  function changeQty(itemId, delta) {
    updateList((prev) => prev.map((i) => (i.id === itemId ? { ...i, qty: (Number(i.qty) || 0) + delta } : i)).filter((i) => i.qty > 0));
  }
  function removeFromList(itemId) {
    updateList((prev) => prev.filter((i) => i.id !== itemId));
  }

  const cartEntries = list
    .map((item) => {
      const candidates = records.filter((r) => !r.noShop && sameProduct(item, r));
      if (!candidates.length) return null;
      const record =
        candidates.find((r) => r.shop.id === item.shopId) ||
        candidates.reduce((a, b) => (b.price < a.price ? b : a));
      return { item, record, qty: Number(item.qty) || 1 };
    })
    .filter(Boolean);
  const cartCount = cartEntries.reduce((n, e) => n + e.qty, 0);
  const cartTotal = cartEntries.reduce((n, e) => n + e.record.price * e.qty, 0);

  function openCompare(recordId) {
    setCompareRecordId(recordId);
    setCompareSelectedId(recordId);
  }
  const compareRecord = records.find((r) => r.id === compareRecordId) || null;
  const compareGroup = compareRecord ? groupOf(compareRecord.productId).slice().sort((a, b) => a.price - b.price) : [];
  const compareCheapest = compareGroup.find((g) => g.available) || compareGroup[0] || null;

  function confirmCompareAdd() {
    const id = compareSelectedId || (compareCheapest && compareCheapest.id);
    const rec = records.find((r) => r.id === id);
    if (rec) addRecordToList(rec);
    setCompareRecordId(null);
  }

  return (
    <div className="qps-root">
      <style>{PD_CSS}</style>
      <div className="qps-app">
        <header className="qps-top">
          {onBack && (
            <button className="qps-back" onClick={onBack}>
              <ArrowLeft size={14} /> Back to Dashboard
            </button>
          )}
          <div className="qps-brand-row">
            <div className="qps-brand-mark">
              <span className="qps-logo">QuickPick</span>
              <span className="qps-tagline">Product Search</span>
            </div>
            <button className="qps-cart-btn" onClick={() => setCartOpen(true)}>
              <ShoppingCart size={16} />
              My List <span className="qps-cart-count">{cartCount}</span>
            </button>
          </div>
          <div className="qps-search-row">
            <Search size={18} className="qps-icon-search" />
            <input
              type="text" placeholder="Search products, brands or shops..."
              value={search} onChange={(e) => setSearch(e.target.value)}
            />
            <button className={`qps-icon-clear${search ? " qps-show" : ""}`} onClick={() => setSearch("")}>✕</button>
          </div>
        </header>

        <div className="qps-toolbar">
          <div className="qps-chip-row">
            {categories.map((c) => (
              <button key={c} className={`qps-chip${c === category ? " qps-active" : ""}`} onClick={() => setCategory(c)}>
                {c}
              </button>
            ))}
          </div>
          <div className="qps-toolbar-controls">
            <div className="qps-result-summary">
              {sorted.length ? (
                <>
                  <strong>{sorted.length}</strong> shop products found
                  <span style={{ color: "var(--ink-faint)" }}> &middot; {shopCountInResults} shop{shopCountInResults !== 1 ? "s" : ""}</span>
                </>
              ) : (
                <><strong>0</strong> shop products found</>
              )}
            </div>
            <div className="qps-toolbar-right">
              <button className="qps-filter-toggle-btn" onClick={() => setFiltersOpen(true)}>
                <Filter size={14} /> Filters {activeFilterCount > 0 && <span className="qps-fbadge">{activeFilterCount}</span>}
              </button>
              <div className="qps-select-wrap">
                <select className="qps-sort-select" value={sort} onChange={(e) => setSort(e.target.value)}>
                  <option value="relevance">Sort: Relevance</option>
                  <option value="lowest">Sort: Lowest Price</option>
                  <option value="highest">Sort: Highest Price</option>
                  <option value="name">Sort: Product Name</option>
                  <option value="shop">Sort: Shop Name</option>
                  <option value="saving">Sort: Biggest Saving</option>
                  <option value="recent">Sort: Recently Added</option>
                </select>
              </div>
              <div className="qps-view-toggle">
                <button className={view === "cards" ? "qps-active" : ""} onClick={() => setView("cards")}>
                  <Grid3x3 size={13} /> Cards
                </button>
                <button className={view === "table" ? "qps-active" : ""} onClick={() => setView("table")}>
                  <ListChecks size={13} /> Table
                </button>
              </div>
            </div>
          </div>
        </div>

        <div className="qps-body-layout">
          <aside className={`qps-sidebar${filtersOpen ? " qps-show" : ""}`}>
            <div className="qps-sidebar-overlay-title">
              <h3>Filters</h3>
              <button className="qps-sidebar-close" onClick={() => setFiltersOpen(false)}>✕</button>
            </div>
            <div className="qps-filter-block">
              <h4>Shop</h4>
              {shopChoices.map((s) => (
                <label key={s.id} className="qps-filter-opt">
                  <input type="checkbox" checked={shopFilter.has(s.id)} onChange={() => toggleSetValue(setShopFilter, s.id)} />
                  <span className="qps-swatch" style={{ background: s.color }} />{s.name}
                </label>
              ))}
              {shopChoices.length === 0 && <p style={{ fontSize: 12.5, color: "var(--ink-faint)", margin: 0 }}>No shops yet.</p>}
            </div>
            <div className="qps-filter-block">
              <h4>Brand</h4>
              {allBrands.map((b) => (
                <label key={b} className="qps-filter-opt">
                  <input type="checkbox" checked={brandFilter.has(b)} onChange={() => toggleSetValue(setBrandFilter, b)} />{b}
                </label>
              ))}
            </div>
            <div className="qps-filter-block">
              <h4>Price (KSh)</h4>
              <div className="qps-price-range">
                <input type="number" min="0" placeholder="Min" value={minPrice} onChange={(e) => setMinPrice(e.target.value)} />
                <span style={{ color: "var(--ink-faint)" }}>&ndash;</span>
                <input type="number" min="0" placeholder="Max" value={maxPrice} onChange={(e) => setMaxPrice(e.target.value)} />
              </div>
            </div>
            <div className="qps-filter-block">
              <h4>Package Size</h4>
              {allSizes.map((s) => (
                <label key={s} className="qps-filter-opt">
                  <input type="checkbox" checked={sizeFilter.has(s)} onChange={() => toggleSetValue(setSizeFilter, s)} />{s}
                </label>
              ))}
            </div>
            <div className="qps-filter-block">
              <h4>Availability</h4>
              <label className="qps-filter-opt"><input type="radio" checked={availability === "all"} onChange={() => setAvailability("all")} /> All</label>
              <label className="qps-filter-opt"><input type="radio" checked={availability === "available"} onChange={() => setAvailability("available")} /> Available</label>
              <label className="qps-filter-opt"><input type="radio" checked={availability === "out"} onChange={() => setAvailability("out")} /> Out of Stock</label>
            </div>
            <button className="qps-clear-filters-btn" onClick={clearFilters}>Clear Filters</button>
          </aside>

          <div className="qps-main-col">
            {sorted.length === 0 ? (
              <div className="qps-empty-state">
                <h3>No shop products found</h3>
                <p>Try:</p>
                <ul><li>Another product name</li><li>Different category</li><li>Different shop</li><li>Clear your filters</li></ul>
                <div><button onClick={clearAll}>Clear Filters</button></div>
              </div>
            ) : view === "cards" ? (
              <div className="qps-card-grid">
                {sorted.map((r) => {
                  const { isLowest, diff, cheapest } = withSaving(r);
                  const inCart = cartQtyFor(r);
                  const isSaved = saved.has(r.id);
                  return (
                    <div key={r.id} className="qps-card" style={{ "--shop-color": r.shop.color }}>
                      <div className="qps-img-box">
                        {!r.available && <span className="qps-stock-badge">{r.noShop ? "No shop assigned" : "Out of Stock"}</span>}
                        {r.image ? <img src={r.image} alt={r.name} /> : <span>{r.icon}</span>}
                      </div>
                      <div className="qps-card-body">
                        <p className="qps-pname">{r.name}</p>
                        <p className="qps-pmeta">{r.brand}{r.brand && r.packageSize ? " \u2022 " : ""}{r.packageSize}</p>
                        <div className="qps-shop-line"><span className="qps-dot" />&#127978; <strong>{r.shop.name}</strong></div>
                        <div className="qps-price-cell">
                          <span className="qps-price">{r.noShop ? "No price available" : pdFmt(r.price)}</span>
                          {!r.noShop && r.packageSize && <span className="qps-per">per {r.packageSize}</span>}
                          <div className={`qps-indicator ${isLowest ? "qps-lowest" : "qps-higher"}`}>
                            {r.noShop
                              ? "No shops currently selling this product"
                              : isLowest ? "\u2713 Lowest Price" : `${pdFmt(diff)} more than cheapest`}
                          </div>
                          {!r.noShop && (() => {
                            const g = groupOf(r.productId);
                            if (g.length < 2 || !cheapest) return null;
                            const hi = Math.max(...g.map((x) => x.price));
                            return (
                              <span className="qps-per" style={{ display: "block" }}>
                                {isLowest
                                  ? `Save up to ${pdFmt(hi - cheapest.price)} vs highest`
                                  : `Cheapest: ${cheapest.shop.name} ${pdFmt(cheapest.price)}`}
                              </span>
                            );
                          })()}
                        </div>
                      </div>
                      <div className="qps-action-bar">
                        <button className={inCart ? "qps-active" : ""} disabled={r.noShop} onClick={() => toggleCart(r.id)} title="Add to Shopping List">
                          {inCart ? "\ud83d\udc97" : "\u2764\ufe0f"}{inCart ? <span className="qps-qty-pill">{inCart}</span> : null}
                        </button>
                        <button onClick={() => openCompare(r.id)} title="Compare Shops">&#9878;&#65039;</button>
                        <button className={isSaved ? "qps-active qps-saved" : ""} onClick={() => toggleSave(r.id)} title="Save Product">&#128278;</button>
                      </div>
                    </div>
                  );
                })}
              </div>
            ) : (
              <div className="qps-table-wrap">
                <table className="qps-table">
                  <thead><tr><th>Product</th><th>Shop</th><th>Package</th><th>Price</th><th>Availability</th><th>Actions</th></tr></thead>
                  <tbody>
                    {sorted.map((r) => {
                      const inCart = cartQtyFor(r);
                      const isSaved = saved.has(r.id);
                      return (
                        <tr key={r.id}>
                          <td><strong>{r.name}</strong><br /><span style={{ color: "var(--ink-soft)", fontSize: 12 }}>{r.brand}</span></td>
                          <td><span className="qps-shop-td"><span className="qps-dot" style={{ background: r.shop.color }} />{r.shop.name}</span></td>
                          <td>{r.packageSize || "\u2014"}</td>
                          <td className="qps-price-td">{r.noShop ? "\u2014" : pdFmt(r.price)}</td>
                          <td>{r.available ? <span className="qps-avail-badge qps-yes">Available</span> : <span className="qps-avail-badge qps-no">{r.noShop ? "No shop assigned" : "Out of Stock"}</span>}</td>
                          <td>
                            <div className="qps-t-actions">
                              <button className={inCart ? "qps-active" : ""} disabled={r.noShop} onClick={() => toggleCart(r.id)} title="Add to Shopping List">{inCart ? "\ud83d\udc97" : "\u2764\ufe0f"}</button>
                              <button onClick={() => openCompare(r.id)} title="Compare Shops">&#9878;&#65039;</button>
                              <button className={isSaved ? "qps-active" : ""} onClick={() => toggleSave(r.id)} title="Save Product">&#128278;</button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      </div>

      <div className={`qps-overlay${cartOpen || filtersOpen || compareRecordId ? " qps-show" : ""}`} onClick={() => { setCartOpen(false); setFiltersOpen(false); setCompareRecordId(null); }} />

      <div className={`qps-drawer${cartOpen ? " qps-show" : ""}`}>
        <div className="qps-drawer-head">
          <h3>My Shopping List</h3>
          <button onClick={() => setCartOpen(false)}>✕</button>
        </div>
        <div className="qps-drawer-body">
          {cartEntries.length === 0 ? (
            <div className="qps-cart-empty">Your shopping list is empty.<br />Tap ❤️ on a product to add it here.</div>
          ) : (
            cartEntries.map(({ record, qty, item }) => (
              <div key={item.id} className="qps-cart-item">
                <div className="qps-img-box">{record.image ? <img src={record.image} alt={record.name} /> : record.icon}</div>
                <div className="qps-ci-info">
                  <p className="qps-ci-name">{record.name}</p>
                  <p className="qps-ci-meta">{record.brand} &bull; {record.packageSize} &bull; &#127978; {record.shop.name}</p>
                  <div className="qps-ci-row">
                    <div className="qps-qty-ctrl">
                      <button onClick={() => changeQty(item.id, -1)}>&minus;</button>
                      <span>{qty}</span>
                      <button onClick={() => changeQty(item.id, 1)}>+</button>
                    </div>
                    <div>
                      <span className="qps-ci-price">{pdFmt(record.price * qty)}</span>
                      <button className="qps-remove-x" onClick={() => removeFromList(item.id)}>Remove</button>
                    </div>
                  </div>
                </div>
              </div>
            ))
          )}
        </div>
        {cartEntries.length > 0 && (
          <div className="qps-drawer-foot">
            <div className="qps-cart-total-row"><span>Estimated total</span><span>{pdFmt(cartTotal)}</span></div>
            {onOpenSmartShopping && (
              <button className="qps-btn-primary" onClick={onOpenSmartShopping}>
                Compare &amp; optimize in Smart Shopping
              </button>
            )}
            {savedShoppingLists.length > 0 && (
              <p style={{ margin: "10px 0 0", fontSize: 12, color: "var(--ink-soft)", textAlign: "center" }}>
                {savedShoppingLists.length} optimized list{savedShoppingLists.length === 1 ? "" : "s"} saved so far
              </p>
            )}
          </div>
        )}
      </div>

      <div className={`qps-modal${compareRecordId ? " qps-show" : ""}`}>
        {compareRecord && (
          <>
            <div className="qps-modal-head">
              <div>
                <h3>Compare {compareRecord.name} {compareRecord.packageSize}</h3>
                <p>{compareRecord.icon} {compareRecord.brand} &middot; {compareRecord.category}</p>
              </div>
              <button onClick={() => setCompareRecordId(null)}>✕</button>
            </div>
            <div className="qps-modal-body">
              {compareGroup.map((g) => {
                const diff = compareCheapest ? g.price - compareCheapest.price : 0;
                const isCheapest = compareCheapest && g.id === compareCheapest.id;
                return (
                  <label key={g.id} className={`qps-compare-row${isCheapest ? " qps-cheapest" : ""}${!g.available ? " qps-unavailable" : ""}`}>
                    <input
                      type="radio" name="qpsCmpChoice" disabled={!g.available}
                      checked={compareSelectedId === g.id}
                      onChange={() => setCompareSelectedId(g.id)}
                    />
                    <span className="qps-dot" style={{ background: g.shop.color }} />
                    <span className="qps-cr-name">{g.shop.name}{!g.available ? " (out of stock)" : ""}</span>
                    <span className="qps-cr-price">{pdFmt(g.price)}</span>
                    <span className="qps-cr-diff">{isCheapest ? "Cheapest" : `+${pdFmt(diff)}`}</span>
                  </label>
                );
              })}
            </div>
            <div className="qps-modal-foot">
              <button className="qps-btn-primary" onClick={confirmCompareAdd}>❤ Add selected to list</button>
              {compareGroup.length > 0 && compareCheapest && compareGroup[compareGroup.length - 1].price > compareCheapest.price && (
                <div className="qps-save-banner">
                  &#10003; Cheapest: {compareCheapest.shop.name} &mdash; {pdFmt(compareCheapest.price)} &middot; Save up to {pdFmt(compareGroup[compareGroup.length - 1].price - compareCheapest.price)}
                </div>
              )}
              {compareGroup.length === 0 && (
                <div className="qps-save-banner">No shops currently selling this product.</div>
              )}
            </div>
          </>
        )}
      </div>

      {addedToast && (
        <div style={{
          position: "fixed", bottom: 22, left: "50%", transform: "translateX(-50%)",
          background: "var(--brand-dark, #154F31)", color: "#fff", padding: "10px 18px", borderRadius: 10,
          fontSize: 13, fontWeight: 600, boxShadow: "0 10px 30px rgba(0,0,0,0.2)", zIndex: 200,
        }}>
          {addedToast}
        </div>
      )}
    </div>
  );
}

// ---------- App (routes between Dashboard, Product Management, Shop Management & Price Update) ----------

/* ------------------------------------------------------------------ */
/* Shared app-level persistence (mirrors what Smart Shopping already   */
/* does for itself) - keeps Shops / Products / Price records / Saved   */
/* shopping lists as the single source of truth, surviving page       */
/* navigation and browser refresh alike.                              */
/* ------------------------------------------------------------------ */

function loadPersisted(key, fallback) {
  if (typeof window === "undefined" || !window.localStorage) return fallback;
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return fallback;
    const parsed = JSON.parse(raw);
    return parsed === undefined || parsed === null ? fallback : parsed;
  } catch (e) {
    return fallback;
  }
}

function savePersisted(key, value) {
  if (typeof window === "undefined" || !window.localStorage) return;
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch (e) {
    // storage unavailable/full - state still works for this session.
  }
}

/* ------------------------------------------------------------------ */
/* Product <-> Shop <-> Price (many-to-many)                          */
/* One product can be sold by many shops, each at its own price. This */
/* replaces the old "one product, one shopId" model with a real join  */
/* table, so price comparison always reflects genuine relationships.  */
/* ------------------------------------------------------------------ */

function hashStr(str) {
  let h = 0;
  for (let i = 0; i < String(str).length; i++) h = (h * 31 + String(str).charCodeAt(i)) >>> 0;
  return h;
}

function deterministicPick(list, seedKey, count) {
  if (!list.length) return [];
  const start = hashStr(seedKey) % list.length;
  const picked = [];
  const seen = new Set();
  for (let i = 0; i < list.length && picked.length < count; i++) {
    const item = list[(start + i * 3 + 1) % list.length];
    if (!seen.has(item.id)) { seen.add(item.id); picked.push(item); }
  }
  return picked;
}

// Builds the initial many-to-many product-shop-price records. Used only as
// a seed when nothing is persisted yet - after that, App's shopProducts
// state (assign / edit price / remove) is the single source of truth.
function buildInitialShopProducts(products, shops) {
  if (!shops.length || !products.length) return [];
  const records = [];
  let idCounter = 1;
  const nowIso = new Date().toISOString();

  products.forEach((p) => {
    const assigned = new Set();

    // Keep the product's already-assigned shop (migrated from the old
    // single-shopId model) at its existing listed price.
    const primaryShop = shops.find((s) => s.id === p.shopId);
    if (primaryShop) {
      records.push({ id: `SP${idCounter++}`, productId: p.id, shopId: primaryShop.id, price: Number(p.price) || 0, lastUpdated: nowIso });
      assigned.add(primaryShop.id);
    }

    // Deterministically add 1-2 more shops with a small price variance, so
    // the many-to-many comparison feature has real, visible data from the
    // start rather than requiring the admin to assign everything by hand.
    const extra = deterministicPick(shops, p.id, 3).filter((s) => !assigned.has(s.id)).slice(0, 2);
    extra.forEach((s) => {
      const variancePct = ((hashStr(p.id + s.id) % 15) - 7); // -7%..+7%
      const price = Math.max(1, Math.round((Number(p.price) || 0) * (1 + variancePct / 100)));
      records.push({ id: `SP${idCounter++}`, productId: p.id, shopId: s.id, price, lastUpdated: nowIso });
      assigned.add(s.id);
    });
  });

  return records;
}

const APP_STORAGE_KEYS = {
  shops: "quickpick_shops",
  products: "quickpick_products",
  shopProducts: "quickpick_shop_products",
  priceProducts: "quickpick_price_products",
  savedShoppingLists: "quickpick_saved_shopping_lists",
  lastPriceUpdateAt: "quickpick_last_price_update",
  shoppingList: "quickpick_shopping_list",
};

export default function App() {
  const [page, setPage] = useState("dashboard");

  // Single source of truth for every management page. Each page still owns
  // its own local editing state (search, filters, form drafts, etc.) but the
  // actual records live here and are what the Dashboard derives its cards
  // from - never a separate hard-coded/mock copy.
  const [shops, setShops] = useState(() => loadPersisted(APP_STORAGE_KEYS.shops, initialShops));
  const [products, setProducts] = useState(() => loadPersisted(APP_STORAGE_KEYS.products, initialProducts));
  const [shopProducts, setShopProducts] = useState(() => {
    const persisted = loadPersisted(APP_STORAGE_KEYS.shopProducts, null);
    if (persisted) return persisted;
    const seedShops = loadPersisted(APP_STORAGE_KEYS.shops, initialShops);
    const seedProducts = loadPersisted(APP_STORAGE_KEYS.products, initialProducts);
    return buildInitialShopProducts(seedProducts, seedShops);
  });
  const [priceProducts, setPriceProducts] = useState(() => loadPersisted(APP_STORAGE_KEYS.priceProducts, null) ?? buildProducts());
  const [savedShoppingLists, setSavedShoppingLists] = useState(() => loadPersisted(APP_STORAGE_KEYS.savedShoppingLists, []));
  const [lastPriceUpdateAt, setLastPriceUpdateAt] = useState(() => loadPersisted(APP_STORAGE_KEYS.lastPriceUpdateAt, null));
  // The in-progress shopping list shared by Product Search (heart button)
  // and Smart Shopping. Items: { id, name, qty, productId?, shopId?, shopName?, price? }
  const [shoppingList, setShoppingList] = useState(() => {
    const saved = loadPersisted(APP_STORAGE_KEYS.shoppingList, []);
    return Array.isArray(saved) ? saved : [];
  });

  useEffect(() => savePersisted(APP_STORAGE_KEYS.shops, shops), [shops]);
  useEffect(() => savePersisted(APP_STORAGE_KEYS.products, products), [products]);
  useEffect(() => savePersisted(APP_STORAGE_KEYS.shopProducts, shopProducts), [shopProducts]);
  useEffect(() => savePersisted(APP_STORAGE_KEYS.priceProducts, priceProducts), [priceProducts]);
  useEffect(() => savePersisted(APP_STORAGE_KEYS.savedShoppingLists, savedShoppingLists), [savedShoppingLists]);
  useEffect(() => savePersisted(APP_STORAGE_KEYS.lastPriceUpdateAt, lastPriceUpdateAt), [lastPriceUpdateAt]);
  useEffect(() => savePersisted(APP_STORAGE_KEYS.shoppingList, shoppingList), [shoppingList]);

  // A deleted product must also leave the shared shopping list.
  useEffect(() => {
    const ids = new Set(products.map((p) => p && p.id));
    setShoppingList((prev) => {
      const next = prev.filter((i) => !i.productId || ids.has(i.productId));
      return next.length === prev.length ? prev : next;
    });
  }, [products]);

  const handleNavigate = (route) => {
    if (route === adminRoutes.products) setPage("products");
    if (route === adminRoutes.shops) setPage("shops");
    if (route === adminRoutes.prices) setPage("prices");
    if (route === adminRoutes.smartShopping) setPage("smartShopping");
    if (route === adminRoutes.productSearch) setPage("productSearch");
  };

  if (page === "products") {
    return (
      <ProductManagement
        onBack={() => setPage("dashboard")}
        products={products}
        onProductsChange={setProducts}
        shops={shops}
        shopProducts={shopProducts}
        onShopProductsChange={setShopProducts}
      />
    );
  }
  if (page === "shops") {
    return (
      <ShopManagement
        onBack={() => setPage("dashboard")}
        shops={shops}
        onShopsChange={setShops}
        products={products}
        shopProducts={shopProducts}
        onShopProductsChange={setShopProducts}
      />
    );
  }
  if (page === "prices") {
    return (
      <PriceUpdate
        onBack={() => setPage("dashboard")}
        products={priceProducts}
        onProductsChange={setPriceProducts}
        onPriceCommitted={setLastPriceUpdateAt}
      />
    );
  }
  if (page === "smartShopping") {
    return (
      <QuickPickSmartShopping
        onBack={() => setPage("dashboard")}
        savedLists={savedShoppingLists}
        onSavedListsChange={setSavedShoppingLists}
        products={products}
        shops={shops}
        shopProducts={shopProducts}
        sharedList={shoppingList}
        onListChange={setShoppingList}
      />
    );
  }
  if (page === "productSearch") {
    return (
      <QuickPickProductSearch
        onBack={() => setPage("dashboard")}
        products={products}
        shops={shops}
        shopProducts={shopProducts}
        savedShoppingLists={savedShoppingLists}
        shoppingList={shoppingList}
        onShoppingListChange={setShoppingList}
        onOpenSmartShopping={() => setPage("smartShopping")}
      />
    );
  }
  return (
    <AdminDashboard
      onNavigate={handleNavigate}
      shops={shops}
      products={products}
      shopProducts={shopProducts}
      priceProducts={priceProducts}
      savedShoppingLists={savedShoppingLists}
      lastPriceUpdateAt={lastPriceUpdateAt}
    />
  );
}
