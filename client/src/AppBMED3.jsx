import { useState, useRef, useCallback } from 'react';
import { generateReportBMED3, getDownloadUrl } from './services/api';

// ── Icons ─────────────────────────────────────────────────────────────────────
const UploadIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="w-10 h-10" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5" />
  </svg>
);

const FileIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
  </svg>
);

const CheckIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M4.5 12.75l6 6 9-13.5" />
  </svg>
);

const DownloadIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3" />
  </svg>
);

const XIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
  </svg>
);

const PlusIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M12 4.5v15m7.5-7.5h-15" />
  </svg>
);

// ── StatCard ──────────────────────────────────────────────────────────────────
function StatCard({ label, value, color }) {
  const colors = {
    red:    'bg-red-50   border-red-200   text-red-700',
    amber:  'bg-amber-50 border-amber-200 text-amber-700',
    blue:   'bg-blue-50  border-blue-200  text-blue-700',
    purple: 'bg-purple-50 border-purple-200 text-purple-700',
  };
  return (
    <div className={`rounded-xl border px-5 py-4 flex flex-col gap-1 ${colors[color]}`}>
      <span className="text-3xl font-bold">{value}</span>
      <span className="text-sm font-medium opacity-80">{label}</span>
    </div>
  );
}

// ── InputField ────────────────────────────────────────────────────────────────
function InputField({ label, name, value, onChange, type = 'text', required, placeholder }) {
  return (
    <div className="flex flex-col gap-1.5">
      <label className="text-sm font-medium text-slate-600">
        {label}{required && <span className="text-red-400 ml-0.5">*</span>}
      </label>
      <input
        type={type}
        name={name}
        value={value}
        onChange={onChange}
        required={required}
        className="px-3.5 py-2.5 rounded-lg border border-slate-200 bg-white text-slate-800 text-sm
                   focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-transparent
                   placeholder:text-slate-350 transition"
        placeholder={placeholder || `Enter ${label.toLowerCase()}`}
      />
    </div>
  );
}

// ── AppBMED3 ─────────────────────────────────────────────────────────────────
export default function AppBMED3({ onBack }) {
  const [files, setFiles]       = useState([]);
  const [mappingFile, setMappingFile] = useState(null);
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading]   = useState(false);
  const [result, setResult]     = useState(null);
  const [error, setError]       = useState('');
  const fileRef = useRef();

  const [form, setForm] = useState({
    examName:    'Summer – 2026',
    schoolName:  'Department of Electronics Engineering',
    programme:   'Biomedical Engineering',
    semester:    'VI Semester',
    date:        new Date().toISOString().split('T')[0],
    condonation: '',
  });

  const handleForm = e => setForm(f => ({ ...f, [e.target.name]: e.target.value }));

  // ── File validation helper ──────────────────────────────────────────────────
  const isValidExcel = f => f.name.endsWith('.xls') || f.name.endsWith('.xlsx');

  // ── Add files (dedup by name) ───────────────────────────────────────────────
  const addFiles = useCallback((incoming) => {
    const valid   = Array.from(incoming).filter(isValidExcel);
    const invalid = Array.from(incoming).filter(f => !isValidExcel(f));
    if (invalid.length > 0) {
      setError(`Skipped ${invalid.length} non-Excel file(s). Only .xls and .xlsx are allowed.`);
    } else {
      setError('');
    }
    setFiles(prev => {
      const existingNames = new Set(prev.map(f => f.name));
      const newOnes = valid.filter(f => !existingNames.has(f.name));
      return [...prev, ...newOnes];
    });
  }, []);

  // ── Drag & drop ─────────────────────────────────────────────────────────────
  const onDrop = useCallback(e => {
    e.preventDefault();
    setDragging(false);
    addFiles(e.dataTransfer.files);
  }, [addFiles]);

  const onFileChange = e => {
    addFiles(e.target.files);
    e.target.value = '';
  };

  const removeFile = (name) => {
    setFiles(prev => prev.filter(f => f.name !== name));
    setResult(null);
  };

  // ── Submit ──────────────────────────────────────────────────────────────────
  const handleSubmit = async e => {
    e.preventDefault();
    if (files.length === 0) { setError('Please upload at least one Excel attendance sheet.'); return; }

    setLoading(true);
    setError('');
    setResult(null);

    try {
      const fd = new FormData();
      files.forEach(f => fd.append('files', f));
      if (mappingFile) {
        fd.append('mappingFile', mappingFile);
      }
      Object.entries(form).forEach(([k, v]) => fd.append(k, v));

      const data = await generateReportBMED3(fd);
      setResult(data);
    } catch (err) {
      setError(err.response?.data?.error || 'Generation failed. Please check your files and try again.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-amber-50 to-slate-100">

      {/* ── Header ─────────────────────────────────────────────────────────── */}
      <header className="bg-white border-b border-slate-200 shadow-sm">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={onBack}
              className="p-2 -ml-2 rounded-lg hover:bg-slate-100 text-slate-500 hover:text-slate-800 transition-colors"
              title="Back to Branch Selection"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M10.5 19.5L3 12m0 0l7.5-7.5M3 12h18" />
              </svg>
            </button>
            <div className="w-10 h-10 rounded-lg bg-amber-600 flex items-center justify-center text-white font-bold text-lg shadow shrink-0">
              R
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-base sm:text-lg font-bold text-slate-800">RBU Detention Report Generator</h1>
                <span className="text-[10px] font-bold tracking-wide px-2 py-0.5 bg-amber-55 text-amber-700 border border-amber-200 rounded-full uppercase shrink-0">
                  Biomed · 3rd Year (VI Sem)
                </span>
              </div>
              <p className="text-xs text-slate-400">Shri Ramdeobaba College of Engineering and Management, Nagpur</p>
            </div>
          </div>

          <button
            type="button"
            onClick={onBack}
            className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-slate-200 hover:border-slate-350 hover:bg-slate-50 text-slate-600 hover:text-slate-800 rounded-lg transition"
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M9 15L3 9m0 0l6-6M3 9h12a6 6 0 010 12h-3" />
            </svg>
            Branch Selection
          </button>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-10 space-y-8">

        {/* ── Info Banner ─────────────────────────────────────────────────── */}
        <div className="bg-amber-50 border border-amber-200 rounded-xl px-5 py-4 text-sm text-amber-800 space-y-1">
          <p className="font-semibold">📋 Biomedical Engineering — VI Semester Detention Report</p>
          <p className="text-xs text-amber-700">
            Expected columns: <span className="font-mono">Roll No. | Unique Id | Seat No | Student Name | OverAll Attendance | BMICRO | BPPD | MLH | BIP | OOPS | RTOS | SSD | INC | BML LAB | BIP LAB | PR-II</span>
          </p>
        </div>

        <form onSubmit={handleSubmit} className="space-y-6">

          {/* ── Upload Zone ──────────────────────────────────────────────────── */}
          <section className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <h2 className="text-base font-semibold text-slate-700 mb-1">1. Upload Attendance Sheets</h2>
            <p className="text-xs text-slate-400 mb-4">
              Upload Biomedical Engineering 6th Semester Excel files (e.g. vi bme.xls). Multiple files will be parsed and merged.
            </p>

            <div
              onDragOver={e => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={onDrop}
              onClick={() => fileRef.current.click()}
              className={`border-2 border-dashed rounded-xl p-8 flex flex-col items-center gap-3 cursor-pointer transition mb-4
                ${dragging
                  ? 'border-amber-400 bg-amber-50 text-amber-600'
                  : 'border-slate-200 hover:border-amber-300 hover:bg-amber-50/40 text-slate-400'
                }`}
            >
              <UploadIcon />
              <p className="text-sm font-medium">
                {files.length === 0
                  ? <>Drop Excel files here or <span className="text-amber-600 underline">browse</span></>
                  : <><span className="text-amber-600 underline flex items-center gap-1 justify-center"><PlusIcon /> Add more files</span></>
                }
              </p>
              <p className="text-xs text-slate-300">Supports .xls and .xlsx — multiple files allowed</p>
              <input ref={fileRef} type="file" accept=".xls,.xlsx" multiple className="hidden" onChange={onFileChange} />
            </div>

            {/* File list */}
            {files.length > 0 && (
              <div className="space-y-2">
                <p className="text-xs font-medium text-slate-500 uppercase tracking-wide">
                  {files.length} file{files.length > 1 ? 's' : ''} selected
                </p>
                {files.map((f, i) => (
                  <div key={f.name} className="flex items-center gap-3 bg-amber-50/50 border border-amber-200 rounded-xl px-4 py-3">
                    <div className="text-amber-600 shrink-0"><FileIcon /></div>
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium text-slate-700 truncate">{f.name}</p>
                      <p className="text-xs text-slate-400">{(f.size / 1024).toFixed(1)} KB</p>
                    </div>
                    <div className="flex items-center gap-2 text-amber-700 text-xs font-medium shrink-0">
                      <CheckIcon /><span>Section {i + 1}</span>
                    </div>
                    <button
                      type="button"
                      onClick={() => removeFile(f.name)}
                      className="ml-1 text-slate-400 hover:text-red-400 transition shrink-0"
                      title="Remove this file"
                    >
                      <XIcon />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </section>

          {/* ── Mapping Upload Zone ──────────────────────────────────────────── */}
          <section className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <h2 className="text-base font-semibold text-slate-700 mb-1">1.5. Dynamic Course Mapping (Optional)</h2>
            <p className="text-xs text-slate-400 mb-4">
              If your curriculum has new subjects or different abbreviations, upload your course mapping Word document (.docx). If left blank, the system uses the default built-in mapping.
            </p>

            <div className="flex items-center gap-4">
              <input
                type="file"
                accept=".docx"
                id="mapping-file-input"
                className="hidden"
                onChange={e => {
                  const file = e.target.files[0];
                  if (file) {
                    if (file.name.endsWith('.docx')) {
                      setMappingFile(file);
                    } else {
                      setError('Only .docx files are allowed for course mapping.');
                    }
                  }
                }}
              />
              {!mappingFile ? (
                <button
                  type="button"
                  onClick={() => document.getElementById('mapping-file-input').click()}
                  className="flex items-center gap-2 px-4 py-2.5 bg-slate-50 border border-slate-200 text-slate-755 rounded-xl hover:bg-slate-100 font-medium text-sm transition shrink-0"
                >
                  <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
                  </svg>
                  Choose .docx File
                </button>
              ) : (
                <div className="flex items-center gap-3 bg-amber-50/50 border border-amber-200 rounded-xl px-4 py-3 w-full">
                  <div className="text-amber-500 shrink-0"><FileIcon /></div>
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium text-slate-700 truncate">{mappingFile.name}</p>
                    <p className="text-xs text-slate-400">{(mappingFile.size / 1024).toFixed(1)} KB</p>
                  </div>
                  <button
                    type="button"
                    onClick={() => setMappingFile(null)}
                    className="ml-1 text-slate-400 hover:text-red-400 transition shrink-0"
                    title="Remove mapping file"
                  >
                    <XIcon />
                  </button>
                </div>
              )}
            </div>
          </section>

          {/* ── Metadata ─────────────────────────────────────────────────────── */}
          <section className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <h2 className="text-base font-semibold text-slate-700 mb-5">2. Report Details</h2>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <InputField label="Exam Name"    name="examName"   value={form.examName}   onChange={handleForm} required />
              <InputField label="School Name"  name="schoolName" value={form.schoolName} onChange={handleForm} required />
              <InputField label="Programme"    name="programme"  value={form.programme}  onChange={handleForm} required />
              <InputField label="Semester"     name="semester"   value={form.semester}   onChange={handleForm} required />
              <InputField label="Date"         name="date"       value={form.date}       onChange={handleForm} type="date" required />
              <div className="flex flex-col gap-1.5">
                <label className="text-sm font-medium text-slate-600">
                  Condonation Seat Numbers
                  <span className="ml-1 text-xs text-slate-400 font-normal">(optional, comma-separated)</span>
                </label>
                <input
                  type="text"
                  name="condonation"
                  value={form.condonation}
                  onChange={handleForm}
                  placeholder="e.g. BMU26RS6027, BMU26RS6002"
                  className="px-3.5 py-2.5 rounded-lg border border-slate-200 bg-white text-slate-800 text-sm
                             focus:outline-none focus:ring-2 focus:ring-amber-500 focus:border-transparent
                             placeholder:text-slate-350 transition"
                />
              </div>
            </div>
          </section>

          {/* ── Error ────────────────────────────────────────────────────────── */}
          {error && (
            <div className="bg-red-50 border border-red-200 text-red-700 text-sm rounded-xl px-5 py-3 flex items-center gap-2">
              <XIcon /><span>{error}</span>
            </div>
          )}

          {/* ── Submit button ─────────────────────────────────────────────────── */}
          <button
            type="submit"
            disabled={loading}
            className="w-full py-3.5 rounded-xl bg-amber-600 hover:bg-amber-700 disabled:bg-amber-300
                       text-white font-semibold text-sm tracking-wide transition shadow-md shadow-amber-100
                       flex items-center justify-center gap-2"
          >
            {loading ? (
              <>
                <span className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                Generating Biomedical Report…
              </>
            ) : (
              `⚡  Generate Biomedical 3rd Year Report${files.length > 1 ? ` (${files.length} sheets)` : ''}`
            )}
          </button>
        </form>

        {/* ── Result ─────────────────────────────────────────────────────────── */}
        {result && (
          <section className="bg-white rounded-2xl shadow-sm border border-green-200 p-6 space-y-6">
            <div className="flex items-center gap-2 text-green-600">
              <CheckIcon />
              <h2 className="text-base font-semibold">Biomedical 3rd Year Detention Report Generated Successfully</h2>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <StatCard label="Table I (Overall < 75%)"     value={result.stats.tableICount}     color="red"    />
              <StatCard label="Table II (Condonation)"      value={result.stats.tableIICount}     color="amber"  />
              <StatCard label="Detained Students"           value={result.stats.detainedStudents} color="blue"   />
              <StatCard label="Detained Subject Entries"    value={result.stats.detainedCourses}  color="purple" />
            </div>

            <a
              href={getDownloadUrl(result.filename)}
              download
              className="flex items-center justify-center gap-2 w-full py-3.5 rounded-xl
                         bg-emerald-600 hover:bg-emerald-700 text-white font-semibold text-sm
                         tracking-wide transition shadow-md shadow-emerald-100"
            >
              <DownloadIcon />
              Download Biomedical Word Document (.docx)
            </a>

            <p className="text-xs text-slate-400 text-center">
              Generated: <span className="font-mono">{result.filename}</span>
            </p>
          </section>
        )}

      </main>

      {/* ── Footer ─────────────────────────────────────────────────────────── */}
      <footer className="text-center py-8 text-xs text-slate-400">
        RBU Detention Report Generator · Academic Administration System
      </footer>
    </div>
  );
}
