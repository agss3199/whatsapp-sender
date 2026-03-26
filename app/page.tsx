'use client';

/**
 * WhatsApp Bulk Sender — single-page app, multi-sheet Excel support
 *
 * Flow:
 *  1. User uploads an Excel file (.xlsx / .xls)
 *  2. SheetJS parses ALL sheets at once — each sheet gets its own row array
 *  3. Phone-number column is auto-detected independently per sheet
 *  4. User can override the column per sheet and toggle sheets on/off
 *  5. All enabled sheets are combined (in order) into one flat phone list
 *  6. User steps through numbers with NEXT / SAME / GO TO ROW
 *  7. Each action opens  https://wa.me/<number>?text=<encoded_message>
 *  8. Progress + settings persisted in localStorage
 */

import { useState, useRef, useCallback, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** Configuration + data for a single Excel sheet */
interface SheetConfig {
  name: string;
  rows: Record<string, unknown>[];
  headers: string[];
  /** The column header selected as the phone-number source */
  phoneCol: string;
  /** Whether this sheet is included in the combined phone list */
  enabled: boolean;
}

/** One valid, normalised phone number ready to send */
interface PhoneEntry {
  /** 1-based row number within that sheet */
  rowNum: number;
  /** Raw cell value as read from Excel */
  raw: string;
  /** E.164 digits without "+" — used in the wa.me URL */
  normalized: string;
  /** Name of the sheet this entry came from */
  sheetName: string;
}

type Notice = { text: string; kind: 'error' | 'success' | 'info' };

// ─────────────────────────────────────────────────────────────────────────────
// Phone normalisation
//
// Accepted input formats:
//   9876543210       → 10-digit local  → prepend default CC
//   +919876543210    → already E.164   → strip "+"
//   919876543210     → 12 digits       → assume CC present
//   09876543210      → leading "0"     → strip 0, prepend CC
//   98765 43210      → spaces          → strip, then apply rules
//   98765-43210      → dashes          → strip, then apply rules
//   00919876543210   → IDD "00"        → strip "00"
//
// Returns null for anything too short / clearly not numeric.
// ─────────────────────────────────────────────────────────────────────────────

function normalizePhone(raw: unknown, countryCode: string): string | null {
  if (raw === null || raw === undefined) return null;

  const str = String(raw).trim();
  if (!str) return null;

  // Detect a leading "+" before we strip everything
  const hasPlus = str.startsWith('+');

  // Strip formatting: spaces, dashes, dots, parens, commas
  const digits = str.replace(/[\s\-().,]/g, '').replace(/\D/g, '');

  if (!digits || digits.length < 7) return null; // too short

  if (hasPlus) return digits;                     // already country-coded

  // IDD "00" prefix → strip, rest is already country-coded
  if (digits.startsWith('00') && digits.length > 6) return digits.slice(2);

  const cc = countryCode.replace(/\D/g, '');

  // Leading "0" STD prefix → replace with CC
  if (digits.startsWith('0') && digits.length <= 11) return cc + digits.slice(1);

  // Classic 10-digit local number → prepend CC
  if (digits.length === 10) return cc + digits;

  // Anything else → assume it already carries a country code
  return digits;
}

// ─────────────────────────────────────────────────────────────────────────────
// Auto-detect the column most likely to hold phone numbers.
//
// Scoring:
//   +15  header contains a phone-related keyword
//   +1   per sampled cell whose stripped digits are 7–15 chars
// ─────────────────────────────────────────────────────────────────────────────

function detectPhoneColumn(data: Record<string, unknown>[]): string | null {
  if (!data.length) return null;

  const headers = Object.keys(data[0]);
  const sample  = data.slice(0, Math.min(30, data.length));

  let bestCol: string | null = null;
  let bestScore = 0;

  for (const header of headers) {
    let score = 0;

    const h = header.toLowerCase();
    if (/phone|mobile|number|cell|contact|whatsapp|tel|ph|no\.?$/.test(h)) score += 15;

    for (const row of sample) {
      const val = String(row[header] ?? '').replace(/[\s\-().,+]/g, '');
      if (/^\d{7,15}$/.test(val)) score += 1;
    }

    if (score > bestScore) { bestScore = score; bestCol = header; }
  }

  return bestCol;
}

// ─────────────────────────────────────────────────────────────────────────────
// localStorage keys
// ─────────────────────────────────────────────────────────────────────────────

const LS = {
  countryCode: 'wa_country_code',
  index:       'wa_current_index',
  fileId:      'wa_file_id',
  message:     'wa_message',
} as const;

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

const clamp = (n: number, lo: number, hi: number) => Math.min(Math.max(n, lo), hi);

// ─────────────────────────────────────────────────────────────────────────────
// Component
// ─────────────────────────────────────────────────────────────────────────────

export default function Home() {

  // ── Per-sheet data (parsed once, mutated only on column/toggle changes) ─────
  const [sheets,       setSheets]       = useState<SheetConfig[]>([]);
  const [fileId,       setFileId]       = useState<string>('');

  // ── User settings (persisted to localStorage) ──────────────────────────────
  const [countryCode,  setCountryCode]  = useState<string>('+91');
  const [message,      setMessage]      = useState<string>('');

  // ── Navigation ──────────────────────────────────────────────────────────────
  const [currentIndex, setCurrentIndex] = useState<number>(0);
  const [goToInput,    setGoToInput]    = useState<string>('');

  // ── UI ──────────────────────────────────────────────────────────────────────
  const [isLoading,    setIsLoading]    = useState<boolean>(false);
  const [notice,       setNotice]       = useState<Notice | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── Restore persisted settings on mount ─────────────────────────────────────
  useEffect(() => {
    const cc  = localStorage.getItem(LS.countryCode);
    const msg = localStorage.getItem(LS.message);
    if (cc)  setCountryCode(cc);
    if (msg) setMessage(msg);
  }, []);

  useEffect(() => { localStorage.setItem(LS.countryCode, countryCode); }, [countryCode]);
  useEffect(() => { localStorage.setItem(LS.message, message);         }, [message]);
  useEffect(() => {
    if (fileId) localStorage.setItem(LS.index, String(currentIndex));
  }, [currentIndex, fileId]);

  // ─────────────────────────────────────────────────────────────────────────────
  // Combined phone list — built from all ENABLED sheets, in sheet order.
  //
  // useMemo: only reruns when sheets array or countryCode changes.
  // Typing in the message box does NOT trigger this.
  // ─────────────────────────────────────────────────────────────────────────────
  const phoneNumbers = useMemo<PhoneEntry[]>(() => {
    const result: PhoneEntry[] = [];

    for (const sheet of sheets) {
      // Skip disabled sheets or sheets with no column chosen
      if (!sheet.enabled || !sheet.phoneCol || !sheet.rows.length) continue;

      for (let i = 0; i < sheet.rows.length; i++) {
        const raw        = String(sheet.rows[i][sheet.phoneCol] ?? '');
        const normalized = normalizePhone(sheet.rows[i][sheet.phoneCol], countryCode);
        if (normalized) {
          result.push({ rowNum: i + 1, raw, normalized, sheetName: sheet.name });
        }
      }
    }

    return result;
  }, [sheets, countryCode]);

  // ─────────────────────────────────────────────────────────────────────────────
  // File upload — parses every sheet independently
  // ─────────────────────────────────────────────────────────────────────────────
  const handleFileUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    setNotice(null);

    // Fingerprint: name + byte-size (good enough to detect a different file)
    const newFileId = `${file.name}_${file.size}`;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const workbook = XLSX.read(evt.target?.result, { type: 'binary' });

        if (!workbook.SheetNames.length) {
          setNotice({ text: 'The workbook contains no sheets.', kind: 'error' });
          setIsLoading(false);
          return;
        }

        // Parse every sheet; auto-detect phone column for each independently
        const parsed: SheetConfig[] = workbook.SheetNames.map(name => {
          const ws = workbook.Sheets[name];

          // raw: false → formatted text strings, preserving leading 0s and "+" signs
          const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, {
            defval: '',
            raw: false,
          });

          const headers  = rows.length ? Object.keys(rows[0]) : [];
          const phoneCol = detectPhoneColumn(rows) ?? headers[0] ?? '';

          return { name, rows, headers, phoneCol, enabled: rows.length > 0 };
        });

        // If every sheet is empty disable all (user will see a clear error)
        const totalRows = parsed.reduce((s, sh) => s + sh.rows.length, 0);
        if (totalRows === 0) {
          setNotice({ text: 'All sheets appear to be empty.', kind: 'error' });
          setIsLoading(false);
          return;
        }

        setSheets(parsed);
        setFileId(newFileId);

        // Restore or reset progress
        const prevFileId = localStorage.getItem(LS.fileId);
        if (prevFileId === newFileId) {
          const saved = parseInt(localStorage.getItem(LS.index) ?? '0', 10);
          setCurrentIndex(isNaN(saved) ? 0 : saved);
        } else {
          localStorage.setItem(LS.fileId, newFileId);
          localStorage.setItem(LS.index, '0');
          setCurrentIndex(0);
        }

        // Build a summary notice, e.g. "3 sheets · Sheet1 (150), Sheet2 (89), Sheet3 (212)"
        const summary = parsed
          .map(s => `${s.name} (${s.rows.length.toLocaleString()})`)
          .join(' · ');
        setNotice({
          text: `${parsed.length} sheet${parsed.length > 1 ? 's' : ''} loaded — ${summary}`,
          kind: 'success',
        });
      } catch {
        setNotice({ text: 'Could not parse the file. Please upload a valid .xlsx or .xls.', kind: 'error' });
      } finally {
        setIsLoading(false);
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };
    reader.readAsBinaryString(file);
  }, []);

  // ─────────────────────────────────────────────────────────────────────────────
  // Per-sheet update helper — patches one sheet by index, resets progress
  // ─────────────────────────────────────────────────────────────────────────────
  const updateSheet = useCallback((idx: number, patch: Partial<SheetConfig>) => {
    setSheets(prev => prev.map((s, i) => i === idx ? { ...s, ...patch } : s));
    setCurrentIndex(0); // any column/toggle change resets the pointer
    setNotice(null);
  }, []);

  // ─────────────────────────────────────────────────────────────────────────────
  // WhatsApp link
  // ─────────────────────────────────────────────────────────────────────────────
  const openWhatsApp = useCallback((normalized: string) => {
    const url = `https://wa.me/${normalized}?text=${encodeURIComponent(message)}`;
    window.open(url, '_blank', 'noopener,noreferrer');
  }, [message]);

  // ─────────────────────────────────────────────────────────────────────────────
  // Button handlers
  // ─────────────────────────────────────────────────────────────────────────────

  const current = phoneNumbers[currentIndex] ?? null;

  /** Open WhatsApp for the current number, then advance the pointer */
  const handleNext = useCallback(() => {
    if (!current) return;
    openWhatsApp(current.normalized);
    setCurrentIndex(prev => clamp(prev + 1, 0, phoneNumbers.length - 1));
  }, [current, phoneNumbers.length, openWhatsApp]);

  /** Re-open the same number WITHOUT moving the pointer */
  const handleSame = useCallback(() => {
    if (!current) return;
    openWhatsApp(current.normalized);
  }, [current, openWhatsApp]);

  /** Jump to a 1-based index in the combined phone list */
  const handleGoTo = useCallback(() => {
    const n = parseInt(goToInput, 10);
    if (isNaN(n) || n < 1 || n > phoneNumbers.length) {
      setNotice({ text: `Enter a number between 1 and ${phoneNumbers.length}.`, kind: 'error' });
      return;
    }
    setCurrentIndex(n - 1);
    setGoToInput('');
    setNotice(null);
  }, [goToInput, phoneNumbers.length]);

  // ─────────────────────────────────────────────────────────────────────────────
  // Derived values
  // ─────────────────────────────────────────────────────────────────────────────

  const total       = phoneNumbers.length;
  const processed   = currentIndex;
  const remaining   = Math.max(0, total - processed);
  const progressPct = total > 0 ? Math.round((processed / total) * 100) : 0;
  const hasFile     = sheets.length > 0;
  const allDone     = hasFile && total > 0 && currentIndex >= total;

  // ─────────────────────────────────────────────────────────────────────────────
  // Tailwind shorthand strings
  // ─────────────────────────────────────────────────────────────────────────────

  const card         = 'bg-white rounded-2xl shadow-sm border border-slate-100 p-5';
  const inputCls     = 'w-full border border-slate-200 rounded-xl px-3 py-3 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-green-400 transition';
  const sectionTitle = 'font-semibold text-xs uppercase tracking-widest text-slate-400 mb-4';
  const labelCls     = 'block text-sm font-medium text-slate-600 mb-1.5';

  // ─────────────────────────────────────────────────────────────────────────────
  // Render
  // ─────────────────────────────────────────────────────────────────────────────

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <div className="max-w-md mx-auto px-4 py-8 space-y-4">

        {/* ── Header ───────────────────────────────────────────────────────── */}
        <div className="text-center pb-1">
          <div className="inline-flex items-center justify-center w-12 h-12 rounded-2xl bg-green-500 mb-3 shadow">
            <svg viewBox="0 0 24 24" className="w-7 h-7 fill-white" xmlns="http://www.w3.org/2000/svg">
              <path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/>
            </svg>
          </div>
          <h1 className="text-2xl font-bold text-slate-800 tracking-tight">WhatsApp Bulk Sender</h1>
          <p className="text-sm text-slate-400 mt-1">Multi-sheet Excel · Normalise numbers · Send one by one</p>
        </div>

        {/* ── 1. Upload ─────────────────────────────────────────────────────── */}
        <section className={card}>
          <p className={sectionTitle}>1 · Upload Excel File</p>

          <input
            ref={fileInputRef}
            id="file-upload"
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileUpload}
            className="hidden"
          />
          <label
            htmlFor="file-upload"
            className={[
              'flex flex-col items-center justify-center w-full py-6 rounded-xl',
              'border-2 border-dashed cursor-pointer select-none transition-colors',
              hasFile
                ? 'border-green-400 bg-green-50 hover:bg-green-100'
                : 'border-slate-300 bg-slate-50 hover:bg-slate-100',
            ].join(' ')}
          >
            {isLoading ? (
              <span className="text-slate-500 text-sm font-medium">⏳ Parsing all sheets…</span>
            ) : hasFile ? (
              <>
                <span className="text-green-600 font-bold text-sm">
                  ✓ {sheets.length} sheet{sheets.length > 1 ? 's' : ''} loaded
                </span>
                <span className="text-slate-400 text-xs mt-1">
                  {sheets.reduce((s, sh) => s + sh.rows.length, 0).toLocaleString()} total rows · tap to replace
                </span>
              </>
            ) : (
              <>
                <span className="text-slate-600 font-semibold text-sm">Tap to upload .xlsx or .xls</span>
                <span className="text-slate-400 text-xs mt-1">All sheets are parsed automatically</span>
              </>
            )}
          </label>

          {notice && (
            <div className={[
              'mt-3 text-xs rounded-xl px-3 py-2.5 leading-relaxed',
              notice.kind === 'error'   ? 'bg-red-50 text-red-600 border border-red-100'     :
              notice.kind === 'success' ? 'bg-green-50 text-green-700 border border-green-100' :
                                          'bg-blue-50 text-blue-600 border border-blue-100',
            ].join(' ')}>
              {notice.text}
            </div>
          )}
        </section>

        {/* ── 2. Per-sheet column config ────────────────────────────────────── */}
        {hasFile && (
          <section className={card}>
            <p className={sectionTitle}>
              2 · Phone Number Column
              {sheets.length > 1 && (
                <span className="ml-2 normal-case font-normal text-slate-400">
                  — configure each sheet independently
                </span>
              )}
            </p>

            <div className="space-y-4">
              {sheets.map((sheet, idx) => {
                // Valid-number count for this sheet (derived locally for display)
                const sheetValidCount = sheet.enabled && sheet.phoneCol
                  ? sheet.rows.filter(
                      row => normalizePhone(row[sheet.phoneCol], countryCode) !== null
                    ).length
                  : 0;

                // Preview: first 3 raw values in the chosen column
                const sheetPreview = sheet.rows.slice(0, 3).map((row, i) => ({
                  rowNum: i + 1,
                  raw: String(row[sheet.phoneCol] ?? ''),
                }));

                return (
                  <div
                    key={sheet.name}
                    className={[
                      'rounded-xl border p-3 transition-colors',
                      sheet.enabled
                        ? 'border-slate-200 bg-white'
                        : 'border-slate-100 bg-slate-50 opacity-60',
                    ].join(' ')}
                  >
                    {/* Sheet header row */}
                    <div className="flex items-center justify-between mb-2.5">
                      <div className="flex items-center gap-2 min-w-0">
                        {/* Enable / disable toggle */}
                        <button
                          onClick={() => updateSheet(idx, { enabled: !sheet.enabled })}
                          className={[
                            'shrink-0 w-9 h-5 rounded-full transition-colors relative',
                            sheet.enabled ? 'bg-green-500' : 'bg-slate-300',
                          ].join(' ')}
                          title={sheet.enabled ? 'Disable this sheet' : 'Enable this sheet'}
                        >
                          <span className={[
                            'absolute top-0.5 w-4 h-4 rounded-full bg-white shadow transition-all',
                            sheet.enabled ? 'left-[18px]' : 'left-0.5',
                          ].join(' ')} />
                        </button>

                        {/* Sheet name */}
                        <span className="font-semibold text-sm text-slate-700 truncate">
                          {sheet.name}
                        </span>
                      </div>

                      {/* Row + valid-number summary */}
                      <span className="shrink-0 text-xs text-slate-400 ml-2">
                        {sheet.rows.length.toLocaleString()} rows
                        {sheet.enabled && (
                          <> · <span className="text-green-600 font-medium">{sheetValidCount.toLocaleString()} valid</span></>
                        )}
                      </span>
                    </div>

                    {/* Column selector — only shown when sheet is enabled */}
                    {sheet.enabled && (
                      <>
                        {sheet.headers.length > 0 ? (
                          <select
                            value={sheet.phoneCol}
                            onChange={e => updateSheet(idx, { phoneCol: e.target.value })}
                            className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-green-400 transition"
                          >
                            {sheet.headers.map(h => (
                              <option key={h} value={h}>{h}</option>
                            ))}
                          </select>
                        ) : (
                          <p className="text-xs text-slate-400 italic">Sheet has no columns</p>
                        )}

                        {/* Preview first 3 rows of selected column */}
                        {sheetPreview.length > 0 && (
                          <div className="mt-2.5 bg-slate-50 rounded-lg p-2.5 space-y-1.5">
                            <p className="text-xs font-medium text-slate-400">Preview:</p>
                            {sheetPreview.map(p => (
                              <div key={p.rowNum} className="flex items-baseline gap-2 text-xs">
                                <span className="text-slate-400 w-10 shrink-0">R{p.rowNum}</span>
                                <span className="font-mono text-slate-700 truncate">
                                  {p.raw || <span className="italic text-slate-300">(empty)</span>}
                                </span>
                              </div>
                            ))}
                          </div>
                        )}
                      </>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Total across all enabled sheets */}
            <p className="mt-3 text-xs text-slate-400">
              Combined total:{' '}
              <span className="font-semibold text-slate-600">{total.toLocaleString()}</span> valid
              number{total !== 1 ? 's' : ''} across{' '}
              {sheets.filter(s => s.enabled).length} enabled sheet{sheets.filter(s => s.enabled).length !== 1 ? 's' : ''}
            </p>
          </section>
        )}

        {/* ── 3. Settings ───────────────────────────────────────────────────── */}
        <section className={card}>
          <p className={sectionTitle}>3 · Settings</p>

          <div className="mb-4">
            <label className={labelCls}>
              Default Country Code
              <span className="ml-1 font-normal text-slate-400">(applied to numbers without one)</span>
            </label>
            <input
              type="text"
              value={countryCode}
              onChange={e => setCountryCode(e.target.value)}
              placeholder="+91"
              className={inputCls}
            />
          </div>

          <div>
            <label className={labelCls}>
              Message
              <span className="ml-1 font-normal text-slate-400">(same for every number)</span>
            </label>
            <textarea
              value={message}
              onChange={e => setMessage(e.target.value)}
              placeholder="Type your WhatsApp message here…"
              rows={4}
              className={`${inputCls} resize-none`}
            />
            {message.length > 0 && (
              <p className="text-xs text-slate-400 mt-1 text-right">{message.length} chars</p>
            )}
          </div>
        </section>

        {/* ── 4. Send controls ──────────────────────────────────────────────── */}
        {hasFile && total > 0 && (
          <section className={`${card} space-y-5`}>
            <p className={sectionTitle}>4 · Send Messages</p>

            {/* Progress bar */}
            <div>
              <div className="flex justify-between items-center mb-1.5">
                <span className="text-xs text-slate-500">Progress</span>
                <span className="text-xs font-semibold text-slate-700">{progressPct}%</span>
              </div>
              <div className="h-2.5 bg-slate-100 rounded-full overflow-hidden">
                <div
                  className="h-full bg-green-500 rounded-full transition-all duration-500 ease-out"
                  style={{ width: `${progressPct}%` }}
                />
              </div>
            </div>

            {/* Stats grid */}
            <div className="grid grid-cols-3 gap-2 text-center">
              {([
                { label: 'Processed', value: processed, bg: 'bg-blue-50',   text: 'text-blue-600'  },
                { label: 'Remaining', value: remaining, bg: 'bg-amber-50',  text: 'text-amber-600' },
                { label: 'Total',     value: total,     bg: 'bg-slate-100', text: 'text-slate-700' },
              ] as const).map(s => (
                <div key={s.label} className={`${s.bg} rounded-xl py-3.5 px-2`}>
                  <p className={`text-2xl font-bold tabular-nums ${s.text}`}>
                    {s.value.toLocaleString()}
                  </p>
                  <p className="text-xs text-slate-500 mt-0.5">{s.label}</p>
                </div>
              ))}
            </div>

            {/* Current number card */}
            {!allDone && current && (
              <div className="bg-green-50 border border-green-100 rounded-xl p-4 text-center">
                {/* Position + sheet badge */}
                <div className="flex items-center justify-center gap-2 mb-1">
                  <span className="text-xs text-slate-500">
                    #{(currentIndex + 1).toLocaleString()} of {total.toLocaleString()}
                  </span>
                  {sheets.length > 1 && (
                    <span className="text-xs font-medium bg-green-100 text-green-700 rounded-full px-2 py-0.5">
                      {current.sheetName}
                    </span>
                  )}
                </div>

                {/* Normalised number */}
                <p className="text-2xl font-bold font-mono tracking-widest text-green-700 break-all">
                  +{current.normalized}
                </p>

                {/* Show original only when it differs */}
                {current.raw !== current.normalized && (
                  <p className="text-xs text-slate-400 mt-1.5">
                    Original: <span className="font-mono">{current.raw}</span>
                  </p>
                )}

                {/* Row in sheet */}
                <p className="text-xs text-slate-400 mt-0.5">
                  Sheet row {current.rowNum}
                </p>
              </div>
            )}

            {/* All-done state */}
            {allDone && (
              <div className="bg-green-50 border border-green-100 rounded-xl p-5 text-center space-y-1">
                <p className="text-xl font-bold text-green-700">All numbers processed!</p>
                <p className="text-sm text-slate-500">
                  Upload a new file or use GO TO ROW to revisit any number.
                </p>
              </div>
            )}

            {/* Action buttons */}
            <div className="space-y-3">
              <button
                onClick={handleNext}
                disabled={!current || allDone}
                className={[
                  'w-full py-4 rounded-2xl font-bold text-lg tracking-wide transition-all',
                  'shadow-sm active:scale-[0.98]',
                  !current || allDone
                    ? 'bg-slate-200 text-slate-400 cursor-not-allowed'
                    : 'bg-green-500 hover:bg-green-600 active:bg-green-700 text-white shadow-green-200 shadow-md',
                ].join(' ')}
              >
                NEXT NUMBER →
              </button>

              <button
                onClick={handleSame}
                disabled={!current}
                className={[
                  'w-full py-3.5 rounded-2xl font-semibold text-base transition-all',
                  'shadow-sm active:scale-[0.98]',
                  !current
                    ? 'bg-slate-200 text-slate-400 cursor-not-allowed'
                    : 'bg-blue-500 hover:bg-blue-600 active:bg-blue-700 text-white',
                ].join(' ')}
              >
                SAME NUMBER ↺
              </button>
            </div>

            {/* Go to row */}
            <div>
              <p className="text-xs text-slate-400 mb-2">
                Jump to any position in the combined list:
              </p>
              <div className="flex gap-2">
                <input
                  type="number"
                  value={goToInput}
                  onChange={e => setGoToInput(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && handleGoTo()}
                  placeholder={`1 – ${total.toLocaleString()}`}
                  min={1}
                  max={total}
                  className={`${inputCls} flex-1 min-w-0`}
                />
                <button
                  onClick={handleGoTo}
                  className="shrink-0 px-4 py-3 rounded-xl bg-slate-700 hover:bg-slate-800 active:bg-slate-900 text-white font-semibold text-sm transition-colors"
                >
                  GO TO ROW
                </button>
              </div>
            </div>
          </section>
        )}

        {/* ── No valid numbers warning ──────────────────────────────────────── */}
        {hasFile && total === 0 && !isLoading && (
          <div className="bg-amber-50 border border-amber-100 rounded-2xl p-4 text-center">
            <p className="font-semibold text-amber-700 text-sm">No valid phone numbers found</p>
            <p className="text-xs text-slate-500 mt-1">
              Check that at least one sheet is enabled and the correct column is selected.
            </p>
          </div>
        )}

        {/* ── Footer ───────────────────────────────────────────────────────── */}
        <p className="text-center text-xs text-slate-400 pb-4">
          All processing happens in your browser — no data leaves your device.
        </p>

      </div>
    </div>
  );
}
