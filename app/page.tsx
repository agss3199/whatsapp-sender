'use client';

import { useState, useRef, useCallback, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';

type SendMode = 'whatsapp' | 'email';

interface SheetConfig {
  name: string;
  rows: Record<string, unknown>[];
  headers: string[];
  phoneCol: string;
  emailCol: string;
  enabled: boolean;
}

interface PhoneEntry {
  kind: 'whatsapp';
  rowNum: number;
  raw: string;
  normalized: string;
  sheetName: string;
}

interface EmailEntry {
  kind: 'email';
  rowNum: number;
  raw: string;
  email: string;
  sheetName: string;
}

type RecipientEntry = PhoneEntry | EmailEntry;

type Notice = { text: string; kind: 'error' | 'success' | 'info' };

function normalizePhone(raw: unknown, countryCode: string): string | null {
  if (raw === null || raw === undefined) return null;

  const str = String(raw).trim();
  if (!str) return null;

  const hasPlus = str.startsWith('+');
  const digits = str.replace(/[\s\-().,]/g, '').replace(/\D/g, '');

  if (!digits || digits.length < 7) return null;
  if (hasPlus) return digits;

  if (digits.startsWith('00') && digits.length > 6) return digits.slice(2);

  const cc = countryCode.replace(/\D/g, '');

  if (digits.startsWith('0') && digits.length <= 11) return cc + digits.slice(1);
  if (digits.length === 10) return cc + digits;

  return digits;
}

function normalizeEmail(raw: unknown): string | null {
  if (raw === null || raw === undefined) return null;
  const value = String(raw).trim().toLowerCase();
  if (!value) return null;

  const cleaned = value.replace(/^mailto:/i, '');
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  return emailRegex.test(cleaned) ? cleaned : null;
}

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

function detectEmailColumn(data: Record<string, unknown>[]): string | null {
  if (!data.length) return null;

  const headers = Object.keys(data[0]);
  const sample = data.slice(0, Math.min(30, data.length));

  let bestCol: string | null = null;
  let bestScore = 0;

  for (const header of headers) {
    let score = 0;

    const h = header.toLowerCase();
    if (/email|e-mail|mail|email id|email address/.test(h)) score += 15;

    for (const row of sample) {
      const val = normalizeEmail(row[header]);
      if (val) score += 1;
    }

    if (score > bestScore) {
      bestScore = score;
      bestCol = header;
    }
  }

  return bestCol;
}

const LS = {
  mode: 'sender_mode',
  countryCode: 'wa_country_code',
  index: 'wa_current_index',
  fileId: 'wa_file_id',
  message: 'wa_message',
  emailSubject: 'email_subject',
  emailBody: 'email_body',
} as const;

const clamp = (n: number, lo: number, hi: number) => Math.min(Math.max(n, lo), hi);

export default function Home() {
  const [mode, setMode] = useState<SendMode>('whatsapp');

  const [sheets, setSheets] = useState<SheetConfig[]>([]);
  const [fileId, setFileId] = useState<string>('');

  const [countryCode, setCountryCode] = useState<string>('+91');
  const [message, setMessage] = useState<string>('');
  const [emailSubject, setEmailSubject] = useState<string>('');
  const [emailBody, setEmailBody] = useState<string>('');

  const [currentIndex, setCurrentIndex] = useState<number>(0);
  const [goToInput, setGoToInput] = useState<string>('');

  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [notice, setNotice] = useState<Notice | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const savedMode = localStorage.getItem(LS.mode);
    const cc = localStorage.getItem(LS.countryCode);
    const msg = localStorage.getItem(LS.message);
    const subj = localStorage.getItem(LS.emailSubject);
    const body = localStorage.getItem(LS.emailBody);

    if (savedMode === 'whatsapp' || savedMode === 'email') setMode(savedMode);
    if (cc) setCountryCode(cc);
    if (msg) setMessage(msg);
    if (subj) setEmailSubject(subj);
    if (body) setEmailBody(body);
  }, []);

  useEffect(() => { localStorage.setItem(LS.mode, mode); }, [mode]);
  useEffect(() => { localStorage.setItem(LS.countryCode, countryCode); }, [countryCode]);
  useEffect(() => { localStorage.setItem(LS.message, message); }, [message]);
  useEffect(() => { localStorage.setItem(LS.emailSubject, emailSubject); }, [emailSubject]);
  useEffect(() => { localStorage.setItem(LS.emailBody, emailBody); }, [emailBody]);

  useEffect(() => {
    if (fileId) localStorage.setItem(LS.index, String(currentIndex));
  }, [currentIndex, fileId]);

  const recipients = useMemo<RecipientEntry[]>(() => {
    const result: RecipientEntry[] = [];

    for (const sheet of sheets) {
      if (!sheet.enabled || !sheet.rows.length) continue;

      if (mode === 'whatsapp') {
        if (!sheet.phoneCol) continue;

        for (let i = 0; i < sheet.rows.length; i++) {
          const raw = String(sheet.rows[i][sheet.phoneCol] ?? '');
          const normalized = normalizePhone(sheet.rows[i][sheet.phoneCol], countryCode);
          if (normalized) {
            result.push({
              kind: 'whatsapp',
              rowNum: i + 1,
              raw,
              normalized,
              sheetName: sheet.name,
            });
          }
        }
      } else {
        if (!sheet.emailCol) continue;

        for (let i = 0; i < sheet.rows.length; i++) {
          const raw = String(sheet.rows[i][sheet.emailCol] ?? '');
          const email = normalizeEmail(sheet.rows[i][sheet.emailCol]);
          if (email) {
            result.push({
              kind: 'email',
              rowNum: i + 1,
              raw,
              email,
              sheetName: sheet.name,
            });
          }
        }
      }
    }

    return result;
  }, [sheets, mode, countryCode]);

  const handleFileUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    setNotice(null);

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

        const parsed: SheetConfig[] = workbook.SheetNames.map((name) => {
          const ws = workbook.Sheets[name];
          const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, {
            defval: '',
            raw: false,
          });

          const headers = rows.length ? Object.keys(rows[0]) : [];
          const phoneCol = detectPhoneColumn(rows) ?? headers[0] ?? '';
          const emailCol = detectEmailColumn(rows) ?? headers[0] ?? '';

          return {
            name,
            rows,
            headers,
            phoneCol,
            emailCol,
            enabled: rows.length > 0,
          };
        });

        const totalRows = parsed.reduce((s, sh) => s + sh.rows.length, 0);
        if (totalRows === 0) {
          setNotice({ text: 'All sheets appear to be empty.', kind: 'error' });
          setIsLoading(false);
          return;
        }

        setSheets(parsed);
        setFileId(newFileId);

        const prevFileId = localStorage.getItem(LS.fileId);
        if (prevFileId === newFileId) {
          const saved = parseInt(localStorage.getItem(LS.index) ?? '0', 10);
          setCurrentIndex(isNaN(saved) ? 0 : saved);
        } else {
          localStorage.setItem(LS.fileId, newFileId);
          localStorage.setItem(LS.index, '0');
          setCurrentIndex(0);
        }

        const summary = parsed.map(s => `${s.name} (${s.rows.length.toLocaleString()})`).join(' · ');
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

  const updateSheet = useCallback((idx: number, patch: Partial<SheetConfig>) => {
    setSheets(prev => prev.map((s, i) => (i === idx ? { ...s, ...patch } : s)));
    setCurrentIndex(0);
    setNotice(null);
  }, []);

  const openWhatsApp = useCallback((normalized: string) => {
    const url = `https://wa.me/${normalized}?text=${encodeURIComponent(message)}`;
    window.open(url, '_blank', 'noopener,noreferrer');
  }, [message]);

  const openGmailCompose = useCallback((email: string) => {
    const url =
      `https://mail.google.com/mail/?view=cm&fs=1` +
      `&to=${encodeURIComponent(email)}` +
      `&su=${encodeURIComponent(emailSubject)}` +
      `&body=${encodeURIComponent(emailBody)}`;

    window.open(url, '_blank', 'noopener,noreferrer');
  }, [emailSubject, emailBody]);

  const current = recipients[currentIndex] ?? null;

  const handleNext = useCallback(() => {
    if (!current) return;

    if (current.kind === 'whatsapp') {
      openWhatsApp(current.normalized);
    } else {
      openGmailCompose(current.email);
    }

    setCurrentIndex(prev => clamp(prev + 1, 0, recipients.length));
  }, [current, recipients.length, openWhatsApp, openGmailCompose]);

  const handleSame = useCallback(() => {
    if (!current) return;

    if (current.kind === 'whatsapp') {
      openWhatsApp(current.normalized);
    } else {
      openGmailCompose(current.email);
    }
  }, [current, openWhatsApp, openGmailCompose]);

  const handleGoTo = useCallback(() => {
    const n = parseInt(goToInput, 10);
    if (isNaN(n) || n < 1 || n > recipients.length) {
      setNotice({ text: `Enter a number between 1 and ${recipients.length}.`, kind: 'error' });
      return;
    }
    setCurrentIndex(n - 1);
    setGoToInput('');
    setNotice(null);
  }, [goToInput, recipients.length]);

  const total = recipients.length;
  const processed = Math.min(currentIndex, total);
  const remaining = Math.max(0, total - processed);
  const progressPct = total > 0 ? Math.round((processed / total) * 100) : 0;
  const hasFile = sheets.length > 0;
  const allDone = hasFile && total > 0 && currentIndex >= total;

  const card = 'bg-white rounded-2xl shadow-sm border border-slate-100 p-5';
  const inputCls = 'w-full border border-slate-200 rounded-xl px-3 py-3 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-green-400 transition';
  const sectionTitle = 'font-semibold text-xs uppercase tracking-widest text-slate-400 mb-4';
  const labelCls = 'block text-sm font-medium text-slate-600 mb-1.5';

  const appTitle = mode === 'whatsapp' ? 'WhatsApp Bulk Sender' : 'Email Bulk Sender';
  const appSubtitle =
    mode === 'whatsapp'
      ? 'Multi-sheet Excel · Normalise numbers · Send one by one'
      : 'Multi-sheet Excel · Gmail compose prefill · Send one by one';

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <div className="max-w-md mx-auto px-4 py-8 space-y-4">

        <div className="text-center pb-1">
          <div className="inline-flex items-center justify-center w-12 h-12 rounded-2xl bg-green-500 mb-3 shadow text-white text-xl font-bold">
            {mode === 'whatsapp' ? 'W' : 'E'}
          </div>
          <h1 className="text-2xl font-bold text-slate-800 tracking-tight">{appTitle}</h1>
          <p className="text-sm text-slate-400 mt-1">{appSubtitle}</p>
        </div>

        <section className={card}>
          <p className={sectionTitle}>Mode</p>

          <div className="grid grid-cols-2 gap-2">
            <button
              onClick={() => {
                setMode('whatsapp');
                setCurrentIndex(0);
              }}
              className={[
                'py-3 rounded-xl font-semibold transition',
                mode === 'whatsapp'
                  ? 'bg-green-500 text-white'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200',
              ].join(' ')}
            >
              WhatsApp
            </button>

            <button
              onClick={() => {
                setMode('email');
                setCurrentIndex(0);
              }}
              className={[
                'py-3 rounded-xl font-semibold transition',
                mode === 'email'
                  ? 'bg-green-500 text-white'
                  : 'bg-slate-100 text-slate-600 hover:bg-slate-200',
              ].join(' ')}
            >
              Email
            </button>
          </div>
        </section>

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
            <div
              className={[
                'mt-3 text-xs rounded-xl px-3 py-2.5 leading-relaxed',
                notice.kind === 'error'
                  ? 'bg-red-50 text-red-600 border border-red-100'
                  : notice.kind === 'success'
                    ? 'bg-green-50 text-green-700 border border-green-100'
                    : 'bg-blue-50 text-blue-600 border border-blue-100',
              ].join(' ')}
            >
              {notice.text}
            </div>
          )}
        </section>

        {hasFile && (
          <section className={card}>
            <p className={sectionTitle}>
              2 · {mode === 'whatsapp' ? 'Phone Number Column' : 'Email Column'}
              {sheets.length > 1 && (
                <span className="ml-2 normal-case font-normal text-slate-400">
                  — configure each sheet independently
                </span>
              )}
            </p>

            <div className="space-y-4">
              {sheets.map((sheet, idx) => {
                const selectedCol = mode === 'whatsapp' ? sheet.phoneCol : sheet.emailCol;

                const sheetValidCount = sheet.enabled && selectedCol
                  ? sheet.rows.filter((row) =>
                      mode === 'whatsapp'
                        ? normalizePhone(row[selectedCol], countryCode) !== null
                        : normalizeEmail(row[selectedCol]) !== null
                    ).length
                  : 0;

                const sheetPreview = sheet.rows.slice(0, 3).map((row, i) => ({
                  rowNum: i + 1,
                  raw: String(row[selectedCol] ?? ''),
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
                    <div className="flex items-center justify-between mb-2.5">
                      <div className="flex items-center gap-2 min-w-0">
                        <button
                          onClick={() => updateSheet(idx, { enabled: !sheet.enabled })}
                          className={[
                            'shrink-0 w-9 h-5 rounded-full transition-colors relative',
                            sheet.enabled ? 'bg-green-500' : 'bg-slate-300',
                          ].join(' ')}
                          title={sheet.enabled ? 'Disable this sheet' : 'Enable this sheet'}
                        >
                          <span
                            className={[
                              'absolute top-0.5 w-4 h-4 rounded-full bg-white shadow transition-all',
                              sheet.enabled ? 'left-[18px]' : 'left-0.5',
                            ].join(' ')}
                          />
                        </button>

                        <span className="font-semibold text-sm text-slate-700 truncate">
                          {sheet.name}
                        </span>
                      </div>

                      <span className="shrink-0 text-xs text-slate-400 ml-2">
                        {sheet.rows.length.toLocaleString()} rows
                        {sheet.enabled && (
                          <> · <span className="text-green-600 font-medium">{sheetValidCount.toLocaleString()} valid</span></>
                        )}
                      </span>
                    </div>

                    {sheet.enabled && (
                      <>
                        {sheet.headers.length > 0 ? (
                          <select
                            value={selectedCol}
                            onChange={(e) =>
                              updateSheet(
                                idx,
                                mode === 'whatsapp'
                                  ? { phoneCol: e.target.value }
                                  : { emailCol: e.target.value }
                              )
                            }
                            className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-green-400 transition"
                          >
                            {sheet.headers.map((h) => (
                              <option key={h} value={h}>{h}</option>
                            ))}
                          </select>
                        ) : (
                          <p className="text-xs text-slate-400 italic">Sheet has no columns</p>
                        )}

                        {sheetPreview.length > 0 && (
                          <div className="mt-2.5 bg-slate-50 rounded-lg p-2.5 space-y-1.5">
                            <p className="text-xs font-medium text-slate-400">Preview:</p>
                            {sheetPreview.map((p) => (
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

            <p className="mt-3 text-xs text-slate-400">
              Combined total:{' '}
              <span className="font-semibold text-slate-600">{total.toLocaleString()}</span>{' '}
              valid {mode === 'whatsapp' ? 'number' : 'email'}
              {total !== 1 ? 's' : ''} across {sheets.filter(s => s.enabled).length} enabled sheet
              {sheets.filter(s => s.enabled).length !== 1 ? 's' : ''}
            </p>
          </section>
        )}

        <section className={card}>
          <p className={sectionTitle}>3 · Settings</p>

          {mode === 'whatsapp' ? (
            <>
              <div className="mb-4">
                <label className={labelCls}>
                  Default Country Code
                  <span className="ml-1 font-normal text-slate-400">(applied to numbers without one)</span>
                </label>
                <input
                  type="text"
                  value={countryCode}
                  onChange={(e) => setCountryCode(e.target.value)}
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
                  onChange={(e) => setMessage(e.target.value)}
                  placeholder="Type your WhatsApp message here…"
                  rows={4}
                  className={`${inputCls} resize-none`}
                />
                {message.length > 0 && (
                  <p className="text-xs text-slate-400 mt-1 text-right">{message.length} chars</p>
                )}
              </div>
            </>
          ) : (
            <>
              <div className="mb-4">
                <label className={labelCls}>Email Subject</label>
                <input
                  type="text"
                  value={emailSubject}
                  onChange={(e) => setEmailSubject(e.target.value)}
                  placeholder="Type email subject…"
                  className={inputCls}
                />
              </div>

              <div>
                <label className={labelCls}>Email Body</label>
                <textarea
                  value={emailBody}
                  onChange={(e) => setEmailBody(e.target.value)}
                  placeholder="Type email body here…"
                  rows={6}
                  className={`${inputCls} resize-none`}
                />
                {emailBody.length > 0 && (
                  <p className="text-xs text-slate-400 mt-1 text-right">{emailBody.length} chars</p>
                )}
              </div>

              <p className="text-xs text-slate-400 mt-3">
                Clicking Next will open Gmail compose in a new tab with recipient, subject, and body prefilled.
              </p>
            </>
          )}
        </section>

        {hasFile && total > 0 && (
          <section className={`${card} space-y-5`}>
            <p className={sectionTitle}>
              4 · {mode === 'whatsapp' ? 'Send Messages' : 'Open Gmail Drafts'}
            </p>

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

            <div className="grid grid-cols-3 gap-2 text-center">
              {([
                { label: 'Processed', value: processed, bg: 'bg-blue-50', text: 'text-blue-600' },
                { label: 'Remaining', value: remaining, bg: 'bg-amber-50', text: 'text-amber-600' },
                { label: 'Total', value: total, bg: 'bg-slate-100', text: 'text-slate-700' },
              ] as const).map((s) => (
                <div key={s.label} className={`${s.bg} rounded-xl py-3.5 px-2`}>
                  <p className={`text-2xl font-bold tabular-nums ${s.text}`}>
                    {s.value.toLocaleString()}
                  </p>
                  <p className="text-xs text-slate-500 mt-0.5">{s.label}</p>
                </div>
              ))}
            </div>

            {!allDone && current && (
              <div className="bg-green-50 border border-green-100 rounded-xl p-4 text-center">
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

                {current.kind === 'whatsapp' ? (
                  <>
                    <p className="text-2xl font-bold font-mono tracking-widest text-green-700 break-all">
                      +{current.normalized}
                    </p>

                    {current.raw !== current.normalized && (
                      <p className="text-xs text-slate-400 mt-1.5">
                        Original: <span className="font-mono">{current.raw}</span>
                      </p>
                    )}
                  </>
                ) : (
                  <>
                    <p className="text-lg font-bold text-green-700 break-all">
                      {current.email}
                    </p>

                    {current.raw !== current.email && (
                      <p className="text-xs text-slate-400 mt-1.5">
                        Original: <span className="font-mono">{current.raw}</span>
                      </p>
                    )}
                  </>
                )}

                <p className="text-xs text-slate-400 mt-0.5">
                  Sheet row {current.rowNum}
                </p>
              </div>
            )}

            {allDone && (
              <div className="bg-green-50 border border-green-100 rounded-xl p-5 text-center space-y-1">
                <p className="text-xl font-bold text-green-700">All entries processed!</p>
                <p className="text-sm text-slate-500">
                  Upload a new file or use GO TO ROW to revisit any entry.
                </p>
              </div>
            )}

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
                {mode === 'whatsapp' ? 'NEXT NUMBER →' : 'NEXT EMAIL →'}
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
                {mode === 'whatsapp' ? 'SAME NUMBER ↺' : 'SAME EMAIL ↺'}
              </button>
            </div>

            <div>
              <p className="text-xs text-slate-400 mb-2">
                Jump to any position in the combined list:
              </p>
              <div className="flex gap-2">
                <input
                  type="number"
                  value={goToInput}
                  onChange={(e) => setGoToInput(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && handleGoTo()}
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

        {hasFile && total === 0 && !isLoading && (
          <div className="bg-amber-50 border border-amber-100 rounded-2xl p-4 text-center">
            <p className="font-semibold text-amber-700 text-sm">
              No valid {mode === 'whatsapp' ? 'phone numbers' : 'email addresses'} found
            </p>
            <p className="text-xs text-slate-500 mt-1">
              Check that at least one sheet is enabled and the correct column is selected.
            </p>
          </div>
        )}

        <p className="text-center text-xs text-slate-400 pb-4">
          All processing happens in your browser — no file is uploaded anywhere.
        </p>
      </div>
    </div>
  );
}
