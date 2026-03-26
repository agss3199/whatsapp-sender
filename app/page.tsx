'use client';

/**
 * WhatsApp Bulk Sender — single-page app
 *
 * Flow:
 *  1. User uploads an Excel file (.xlsx / .xls)
 *  2. SheetJS parses it once → rows stored in state
 *  3. Phone-number column is auto-detected (user can override)
 *  4. Numbers are normalised to E.164 digits via useMemo (runs only when
 *     rows / column / country-code change — safe for 10 000 rows)
 *  5. User steps through numbers with NEXT / SAME / GO TO ROW
 *  6. Each action opens  https://wa.me/<number>?text=<encoded_message>
 *  7. Progress + settings persisted in localStorage
 */

import { useState, useRef, useCallback, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

interface PhoneEntry {
  /** 1-based row number in the original sheet (for display) */
  rowNum: number;
  /** Raw cell value exactly as read from Excel */
  raw: string;
  /** E.164 digits without "+" — used in the wa.me URL */
  normalized: string;
}

type Notice = { text: string; kind: 'error' | 'success' | 'info' };

// ─────────────────────────────────────────────────────────────────────────────
// Phone normalisation
//
// Accepted formats:
//   9876543210        → 10-digit local → prepend default CC
//   +919876543210     → already E.164  → strip "+"
//   919876543210      → 12 digits, assume CC already present
//   09876543210       → leading 0      → strip 0, prepend CC
//   98765 43210       → spaces         → strip, then apply rules
//   98765-43210       → dashes         → strip, then apply rules
//   00919876543210    → IDD "00"       → strip "00"
//
// Returns null for anything too short / clearly not a number.
// ─────────────────────────────────────────────────────────────────────────────

function normalizePhone(raw: unknown, countryCode: string): string | null {
  if (raw === null || raw === undefined) return null;

  const str = String(raw).trim();
  if (!str) return null;

  // Detect a leading "+" before stripping non-digits
  const hasPlus = str.startsWith('+');

  // Strip every formatting character: spaces, dashes, dots, parens, commas
  const digits = str.replace(/[\s\-().,]/g, '').replace(/\D/g, '');

  // Reject anything implausibly short
  if (!digits || digits.length < 7) return null;

  // Already country-coded with "+"  →  trust the digits as-is
  if (hasPlus) return digits;

  // IDD prefix "00"  →  strip it, remainder is country-coded
  if (digits.startsWith('00') && digits.length > 6) return digits.slice(2);

  // Sanitise the default country code (accept "+91" or "91")
  const cc = countryCode.replace(/\D/g, '');

  // Leading "0" STD prefix (common in many countries)  →  replace with CC
  if (digits.startsWith('0') && digits.length <= 11) return cc + digits.slice(1);

  // Classic 10-digit local number  →  prepend CC
  if (digits.length === 10) return cc + digits;

  // Anything else: assume the number already carries a country code
  return digits;
}

// ─────────────────────────────────────────────────────────────────────────────
// Auto-detect the column most likely to hold phone numbers.
//
// Scoring:
//   +15  if the header name contains a phone-related keyword
//   +1   for every cell (up to 30 sampled) whose digits are 7–15 chars long
// ─────────────────────────────────────────────────────────────────────────────

function detectPhoneColumn(data: Record<string, unknown>[]): string | null {
  if (!data.length) return null;

  const headers = Object.keys(data[0]);
  const sample = data.slice(0, Math.min(30, data.length));

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

    if (score > bestScore) {
      bestScore = score;
      bestCol = header;
    }
  }

  return bestCol;
}

// ─────────────────────────────────────────────────────────────────────────────
// localStorage key constants
// ─────────────────────────────────────────────────────────────────────────────

const LS = {
  countryCode: 'wa_country_code',
  index:       'wa_current_index',
  fileId:      'wa_file_id',
  message:     'wa_message',
} as const;

// ─────────────────────────────────────────────────────────────────────────────
// Helper: clamp a number between min and max
// ─────────────────────────────────────────────────────────────────────────────

const clamp = (n: number, min: number, max: number) => Math.min(Math.max(n, min), max);

// ─────────────────────────────────────────────────────────────────────────────
// Component
// ─────────────────────────────────────────────────────────────────────────────

export default function Home() {

  // ── Raw sheet data (replaced only on new file upload) ──────────────────────
  const [headers,    setHeaders]    = useState<string[]>([]);
  const [rows,       setRows]       = useState<Record<string, unknown>[]>([]);
  const [phoneCol,   setPhoneCol]   = useState<string>('');
  const [fileId,     setFileId]     = useState<string>('');

  // ── User settings (persisted) ───────────────────────────────────────────────
  const [countryCode, setCountryCode] = useState<string>('+91');
  const [message,     setMessage]     = useState<string>('');

  // ── Navigation state ────────────────────────────────────────────────────────
  const [currentIndex, setCurrentIndex] = useState<number>(0);
  const [goToInput,    setGoToInput]    = useState<string>('');

  // ── UI feedback ─────────────────────────────────────────────────────────────
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [notice,    setNotice]    = useState<Notice | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── Restore persisted settings on first render ─────────────────────────────
  useEffect(() => {
    const cc  = localStorage.getItem(LS.countryCode);
    const msg = localStorage.getItem(LS.message);
    if (cc)  setCountryCode(cc);
    if (msg) setMessage(msg);
  }, []);

  // ── Persist settings whenever they change ──────────────────────────────────
  useEffect(() => { localStorage.setItem(LS.countryCode, countryCode); }, [countryCode]);
  useEffect(() => { localStorage.setItem(LS.message,     message);     }, [message]);

  // ── Persist current index whenever it changes (only while a file is loaded) ─
  useEffect(() => {
    if (fileId) localStorage.setItem(LS.index, String(currentIndex));
  }, [currentIndex, fileId]);

  // ─────────────────────────────────────────────────────────────────────────────
  // Derive the list of valid PhoneEntry objects from raw rows.
  //
  // useMemo ensures this O(n) pass only reruns when its three inputs change —
  // typing in the message textarea will NOT trigger it.
  // ─────────────────────────────────────────────────────────────────────────────
  const phoneNumbers = useMemo<PhoneEntry[]>(() => {
    if (!phoneCol || !rows.length) return [];

    const result: PhoneEntry[] = [];
    for (let i = 0; i < rows.length; i++) {
      const raw        = String(rows[i][phoneCol] ?? '');
      const normalized = normalizePhone(rows[i][phoneCol], countryCode);
      if (normalized) result.push({ rowNum: i + 1, raw, normalized });
    }
    return result;
  }, [rows, phoneCol, countryCode]);

  // ─────────────────────────────────────────────────────────────────────────────
  // File upload handler
  // ─────────────────────────────────────────────────────────────────────────────
  const handleFileUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    setNotice(null);

    // Simple file identity fingerprint: name + byte-size.
    // Used to decide whether to restore saved progress or reset it.
    const newFileId = `${file.name}_${file.size}`;

    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const workbook = XLSX.read(evt.target?.result, { type: 'binary' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        // raw: false → SheetJS returns formatted text strings.
        // This preserves leading zeros and "+" signs that Excel would otherwise
        // drop when storing a phone number as a plain numeric cell.
        const jsonData = XLSX.utils.sheet_to_json<Record<string, unknown>>(firstSheet, {
          defval: '',
          raw: false,
        });

        if (!jsonData.length) {
          setNotice({ text: 'The file is empty or has no data rows.', kind: 'error' });
          setIsLoading(false);
          return;
        }

        const hdrs        = Object.keys(jsonData[0]);
        const detectedCol = detectPhoneColumn(jsonData) ?? hdrs[0] ?? '';

        // Commit sheet data to state — this is the single source of truth
        setHeaders(hdrs);
        setRows(jsonData);
        setPhoneCol(detectedCol);
        setFileId(newFileId);

        // Restore progress for the same file; otherwise start fresh
        const prevFileId = localStorage.getItem(LS.fileId);
        if (prevFileId === newFileId) {
          const saved = parseInt(localStorage.getItem(LS.index) ?? '0', 10);
          setCurrentIndex(isNaN(saved) ? 0 : saved);
        } else {
          localStorage.setItem(LS.fileId, newFileId);
          localStorage.setItem(LS.index, '0');
          setCurrentIndex(0);
        }

        setNotice({
          text: `Loaded ${jsonData.length.toLocaleString()} rows · auto-detected column: "${detectedCol}"`,
          kind: 'success',
        });
      } catch {
        setNotice({ text: 'Could not parse the file. Please upload a valid .xlsx or .xls.', kind: 'error' });
      } finally {
        setIsLoading(false);
        // Reset the input so the same file can be re-uploaded if needed
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };

    reader.readAsBinaryString(file);
  }, []);

  // ─────────────────────────────────────────────────────────────────────────────
  // Open WhatsApp deep-link in a new tab
  // ─────────────────────────────────────────────────────────────────────────────
  const openWhatsApp = useCallback((normalized: string) => {
    const url = `https://wa.me/${normalized}?text=${encodeURIComponent(message)}`;
    window.open(url, '_blank', 'noopener,noreferrer');
  }, [message]);

  // ─────────────────────────────────────────────────────────────────────────────
  // Button handlers
  // ─────────────────────────────────────────────────────────────────────────────

  const current = phoneNumbers[currentIndex] ?? null;

  /** Open current number in WhatsApp, then advance the pointer */
  const handleNext = useCallback(() => {
    if (!current) return;
    openWhatsApp(current.normalized);
    setCurrentIndex(prev => clamp(prev + 1, 0, phoneNumbers.length - 1));
  }, [current, phoneNumbers.length, openWhatsApp]);

  /** Re-open current number in WhatsApp WITHOUT advancing */
  const handleSame = useCallback(() => {
    if (!current) return;
    openWhatsApp(current.normalized);
  }, [current, openWhatsApp]);

  /** Jump directly to a user-specified 1-based row in the valid-number list */
  const handleGoTo = useCallback(() => {
    const n = parseInt(goToInput, 10);
    if (isNaN(n) || n < 1 || n > phoneNumbers.length) {
      setNotice({ text: `Please enter a number between 1 and ${phoneNumbers.length}.`, kind: 'error' });
      return;
    }
    setCurrentIndex(n - 1);
    setGoToInput('');
    setNotice(null);
  }, [goToInput, phoneNumbers.length]);

  // ─────────────────────────────────────────────────────────────────────────────
  // Derived display values
  // ─────────────────────────────────────────────────────────────────────────────

  const total       = phoneNumbers.length;
  const processed   = currentIndex;
  const remaining   = Math.max(0, total - processed);
  const progressPct = total > 0 ? Math.round((processed / total) * 100) : 0;
  const allDone     = rows.length > 0 && total > 0 && currentIndex >= total;

  /** Raw values of the first 5 rows in the chosen column — for user confirmation */
  const preview = useMemo(() => {
    if (!phoneCol || !rows.length) return [];
    return rows.slice(0, 5).map((row, i) => ({
      rowNum: i + 1,
      raw: String(row[phoneCol] ?? ''),
    }));
  }, [rows, phoneCol]);

  const hasFile = rows.length > 0;

  // ─────────────────────────────────────────────────────────────────────────────
  // Tailwind helper strings
  // ─────────────────────────────────────────────────────────────────────────────

  const card   = 'bg-white rounded-2xl shadow-sm border border-slate-100 p-5';
  const label  = 'block text-sm font-medium text-slate-600 mb-1.5';
  const input  = 'w-full border border-slate-200 rounded-xl px-3 py-3 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-green-400 transition';
  const sectionTitle = 'font-semibold text-xs uppercase tracking-widest text-slate-400 mb-4';

  // ─────────────────────────────────────────────────────────────────────────────
  // Render
  // ─────────────────────────────────────────────────────────────────────────────

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <div className="max-w-md mx-auto px-4 py-8 space-y-4">

        {/* ── App header ───────────────────────────────────────────────────── */}
        <div className="text-center pb-1">
          <div className="inline-flex items-center justify-center w-12 h-12 rounded-2xl bg-green-500 mb-3 shadow">
            <svg viewBox="0 0 24 24" className="w-7 h-7 fill-white" xmlns="http://www.w3.org/2000/svg">
              <path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/>
            </svg>
          </div>
          <h1 className="text-2xl font-bold text-slate-800 tracking-tight">WhatsApp Bulk Sender</h1>
          <p className="text-sm text-slate-400 mt-1">Upload Excel · Normalise numbers · Send one by one</p>
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
              <span className="text-slate-500 text-sm font-medium">⏳ Parsing…</span>
            ) : hasFile ? (
              <>
                <span className="text-green-600 font-bold text-sm">✓ File loaded</span>
                <span className="text-slate-400 text-xs mt-1">
                  {rows.length.toLocaleString()} rows · tap to replace
                </span>
              </>
            ) : (
              <>
                <span className="text-slate-600 font-semibold text-sm">Tap to upload .xlsx or .xls</span>
                <span className="text-slate-400 text-xs mt-1">Supports up to 10,000 rows</span>
              </>
            )}
          </label>

          {/* Inline notice / error */}
          {notice && (
            <div className={[
              'mt-3 text-xs rounded-xl px-3 py-2.5 leading-relaxed',
              notice.kind === 'error'   ? 'bg-red-50 text-red-600 border border-red-100' :
              notice.kind === 'success' ? 'bg-green-50 text-green-700 border border-green-100' :
                                          'bg-blue-50 text-blue-600 border border-blue-100',
            ].join(' ')}>
              {notice.text}
            </div>
          )}
        </section>

        {/* ── 2. Column selection + preview ─────────────────────────────────── */}
        {hasFile && (
          <section className={card}>
            <p className={sectionTitle}>2 · Phone Number Column</p>

            <select
              value={phoneCol}
              onChange={e => {
                setPhoneCol(e.target.value);
                setCurrentIndex(0); // reset progress when column changes
                setNotice(null);
              }}
              className={input}
            >
              {headers.map(h => (
                <option key={h} value={h}>{h}</option>
              ))}
            </select>

            {/* First-5-rows preview so the user can confirm the right column */}
            {preview.length > 0 && (
              <div className="mt-4 bg-slate-50 rounded-xl p-3 space-y-2">
                <p className="text-xs font-semibold text-slate-400">Preview — first 5 rows:</p>
                {preview.map(p => (
                  <div key={p.rowNum} className="flex items-baseline gap-3 text-xs">
                    <span className="text-slate-400 w-12 shrink-0">Row {p.rowNum}</span>
                    <span className="font-mono text-slate-700 truncate">
                      {p.raw || <span className="italic text-slate-300">(empty)</span>}
                    </span>
                  </div>
                ))}
              </div>
            )}

            <p className="mt-3 text-xs text-slate-400">
              <span className="font-semibold text-slate-600">{total.toLocaleString()}</span> valid
              number{total !== 1 ? 's' : ''} detected
              {rows.length !== total && (
                <> out of {rows.length.toLocaleString()} rows
                  {' '}({(rows.length - total).toLocaleString()} skipped — empty or invalid)</>
              )}
            </p>
          </section>
        )}

        {/* ── 3. Country code + message ──────────────────────────────────────── */}
        <section className={card}>
          <p className={sectionTitle}>3 · Settings</p>

          <div className="mb-4">
            <label className={label}>
              Default Country Code
              <span className="ml-1 font-normal text-slate-400">(applied to numbers that have none)</span>
            </label>
            <input
              type="text"
              value={countryCode}
              onChange={e => setCountryCode(e.target.value)}
              placeholder="+91"
              className={input}
            />
          </div>

          <div>
            <label className={label}>
              Message
              <span className="ml-1 font-normal text-slate-400">(same for every number)</span>
            </label>
            <textarea
              value={message}
              onChange={e => setMessage(e.target.value)}
              placeholder="Type your WhatsApp message here…"
              rows={4}
              className={`${input} resize-none`}
            />
            {message.length > 0 && (
              <p className="text-xs text-slate-400 mt-1 text-right">
                {message.length} chars
              </p>
            )}
          </div>
        </section>

        {/* ── 4. Progress + send controls ───────────────────────────────────── */}
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

            {/* Stats: Processed / Remaining / Total */}
            <div className="grid grid-cols-3 gap-2 text-center">
              {([
                { label: 'Processed', value: processed, bg: 'bg-blue-50',   text: 'text-blue-600'   },
                { label: 'Remaining', value: remaining, bg: 'bg-amber-50',  text: 'text-amber-600'  },
                { label: 'Total',     value: total,     bg: 'bg-slate-100', text: 'text-slate-700'  },
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
                <p className="text-xs text-slate-500 mb-1">
                  Current · #{(currentIndex + 1).toLocaleString()} of {total.toLocaleString()}
                </p>
                <p className="text-2xl font-bold font-mono tracking-widest text-green-700 break-all">
                  +{current.normalized}
                </p>
                {/* Show the original cell value when it differs from the normalised form */}
                {current.raw !== current.normalized && (
                  <p className="text-xs text-slate-400 mt-1.5">
                    Original: <span className="font-mono">{current.raw}</span>
                  </p>
                )}
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

            {/* ── Action buttons ─────────────────────────────────────────── */}
            <div className="space-y-3">

              {/* NEXT NUMBER — primary CTA, large touch target */}
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

              {/* SAME NUMBER — secondary */}
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

            {/* GO TO ROW */}
            <div>
              <p className="text-xs text-slate-400 mb-2">Jump to a specific row in the valid-number list:</p>
              <div className="flex gap-2">
                <input
                  type="number"
                  value={goToInput}
                  onChange={e => setGoToInput(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && handleGoTo()}
                  placeholder={`1 – ${total.toLocaleString()}`}
                  min={1}
                  max={total}
                  className={`${input} flex-1 min-w-0`}
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
              Try selecting a different column or check your country code.
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
