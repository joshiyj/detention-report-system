import React from 'react';

// ── Icons ─────────────────────────────────────────────────────────────────────
const SparklesIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5 text-indigo-400 animate-pulse" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M5 3v4M3 5h4M6 17v4m-2-2h4m5-16l2.286 6.857L21 12l-5.714 2.143L13 21l-2.286-6.857L5 12l5.714-2.143L13 3z" />
  </svg>
);

const LockIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M16.5 10.5V6.75a4.5 4.5 0 10-9 0v3.75m-.75 11.25h10.5a2.25 2.25 0 002.25-2.25v-6.75a2.25 2.25 0 00-2.25-2.25H6.75a2.25 2.25 0 00-2.25 2.25v6.75a2.25 2.25 0 002.25 2.25z" />
  </svg>
);

const ArrowRightIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5 transition-transform group-hover:translate-x-1" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M13.5 4.5L21 12m0 0l-7.5 7.5M21 12H3" />
  </svg>
);

const BackIcon = () => (
  <svg xmlns="http://www.w3.org/2000/svg" className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
    <path strokeLinecap="round" strokeLinejoin="round" d="M10.5 19.5L3 12m0 0l7.5-7.5M3 12h18" />
  </svg>
);

const EcsIcon = () => (
  <div className="w-12 h-12 rounded-xl bg-gradient-to-tr from-blue-500 to-indigo-500 flex items-center justify-center text-white shadow-lg shadow-blue-500/20">
    <svg xmlns="http://www.w3.org/2000/svg" className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M9.75 17L9 20l-1 1h8l-1-1-.75-3M3 13h18M5 17h14a2 2 0 002-2V5a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
    </svg>
  </div>
);

const BiomedIcon = ({ active }) => (
  <div className={`w-12 h-12 rounded-xl flex items-center justify-center text-white shadow-lg ${
    active
      ? 'bg-gradient-to-tr from-amber-500 to-orange-500 shadow-amber-500/20'
      : 'bg-slate-800/80 border border-slate-700/50 text-slate-500'
  }`}>
    <svg xmlns="http://www.w3.org/2000/svg" className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M4.318 6.318a4.5 4.5 0 000 6.364L12 20.364l7.682-7.682a4.5 4.5 0 00-6.364-6.364L12 7.636l-1.318-1.318a4.5 4.5 0 00-6.364 0z" />
    </svg>
  </div>
);

export default function BranchSelectionPage({ selectedYear, onSelectBranch, onBack }) {
  const getYearName = (yr) => {
    switch (yr) {
      case '1': return '1st Year (I & II Sem)';
      case '2': return '2nd Year (III & IV Sem)';
      case '3': return '3rd Year (V & VI Sem)';
      case '4': return '4th Year (VII & VIII Sem)';
      default: return `${yr} Year`;
    }
  };

  const isBiomedActive = selectedYear === '2' || selectedYear === '3';

  const branches = [
    {
      id: 'ECS',
      title: 'Electronics & Computer Science',
      subtitle: 'ECS Department Core',
      description: 'Fully active curriculum-mapped detention builder supporting merged batch/section spreadsheets, overall attendance filtering, and vertical-merged Table III summaries.',
      active: true,
      features: [
        'Curriculum-mapped core subjects',
        'Automatic Theory + Practical avg handler',
        'Alphabetical Exam Seat No. sorting (Y1-Y3)',
        'Fidelity insertion merge order (Y4)'
      ],
      icon: <EcsIcon />,
      color: 'from-blue-600 to-indigo-600',
      badge: 'Fully Integrated'
    },
    {
      id: 'BIOMED',
      title: 'Biomedical Engineering',
      subtitle: 'Biomed Department Core',
      description: isBiomedActive
        ? (selectedYear === '3'
            ? 'Fully integrated Biomedical Engineering VI Semester detention report builder with BMICRO, BPPD, MLH, BIP, OOPS, RTOS, SSD, BML LAB, BIP LAB, PR-II mapped.'
            : 'Fully integrated Biomedical Engineering IV Semester detention report builder with HAPE-II, FAIML, SPA, Biomechanics, MIOT, DF3D and all sem-4 courses mapped.')
        : 'Detention audit sheet analyzer designed for Biomedical course schemas, checking core physiological systems, lab modules, and clinical training attendance.',
      active: isBiomedActive,
      features: selectedYear === '3'
        ? [
            'Biomed Sem-6 syllabus mapping & codes',
            'BMICRO, BPPD, MLH, BIP, OOPS, RTOS, SSD',
            'Theory + Practical (MLH/BIP) average pairing',
            'Word (.docx) Detention Tables compilation'
          ]
        : [
            'Biomed Sem-4 syllabus mapping & codes',
            'HAPE II, FAIML, SPA, BMECH, MIOT, DF3D',
            'Section-wise spreadsheet merge builder',
            'Word (.docx) Detention Tables compilation'
          ],
      icon: <BiomedIcon active={isBiomedActive} />,
      color: 'from-amber-600 to-orange-600',
      badge: isBiomedActive ? 'Active Module' : 'Active Soon'
    }
  ];

  return (
    <div className="min-h-screen bg-[#0f172a] text-slate-100 flex flex-col relative overflow-hidden font-sans selection:bg-indigo-500/30">
      
      {/* ── Ambient Glow Background Elements ────────────────────────────────────── */}
      <div className="absolute top-[-10%] left-[-10%] w-[50%] h-[50%] rounded-full bg-blue-500/10 blur-[120px] pointer-events-none" />
      <div className="absolute bottom-[-10%] right-[-10%] w-[50%] h-[50%] rounded-full bg-indigo-500/10 blur-[120px] pointer-events-none" />

      {/* ── Header ─────────────────────────────────────────────────────────── */}
      <header className="relative z-10 border-b border-slate-800/80 bg-slate-900/40 backdrop-blur-md">
        <div className="max-w-7xl mx-auto px-6 py-5 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <button
              onClick={onBack}
              className="p-2 -ml-2 rounded-lg hover:bg-slate-800/60 text-slate-400 hover:text-slate-200 transition-colors"
              title="Back to Year Selection"
            >
              <BackIcon />
            </button>
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 to-indigo-600 flex items-center justify-center text-white font-bold text-lg shadow-lg shadow-blue-500/20">
              R
            </div>
            <div>
              <h1 className="text-md font-bold tracking-tight text-white flex items-center gap-2">
                RBU Detention System
                <span className="text-[10px] font-bold px-2 py-0.5 bg-blue-500/10 text-blue-400 border border-blue-500/20 rounded-full">
                  {getYearName(selectedYear)}
                </span>
              </h1>
              <p className="text-[11px] text-slate-400 font-medium">Shri Ramdeobaba University, Nagpur</p>
            </div>
          </div>

          <button
            onClick={onBack}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-slate-800/60 border border-slate-700/40 hover:border-slate-600 text-slate-300 hover:text-slate-100 rounded-lg transition"
          >
            <BackIcon /> Year Selection
          </button>
        </div>
      </header>

      {/* ── Main Content ───────────────────────────────────────────────────── */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-6 py-12 lg:py-20 relative z-10 flex flex-col justify-center items-center">
        
        {/* Headline */}
        <div className="text-center max-w-3xl mb-16 space-y-4">
          <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-indigo-500/10 border border-indigo-500/20 text-xs font-semibold text-indigo-300">
            <SparklesIcon /> Academic Department Portal
          </div>
          <h2 className="text-4xl sm:text-5xl font-extrabold tracking-tight text-white leading-tight">
            Select Academic <span className="bg-gradient-to-r from-blue-400 via-indigo-400 to-teal-400 bg-clip-text text-transparent">Branch / Department</span>
          </h2>
          <p className="text-base text-slate-400 max-w-2xl mx-auto leading-relaxed">
            Choose your branch to configure program defaults, course code matrices, and custom detention templates for <span className="text-indigo-400 font-semibold">{getYearName(selectedYear)}</span>.
          </p>
        </div>

        {/* Branch Cards Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 lg:gap-8 w-full max-w-4xl">
          {branches.map((branch) => (
            <div
              key={branch.id}
              onClick={() => branch.active && onSelectBranch(branch.id)}
              className={`group relative rounded-2xl p-6 lg:p-8 border backdrop-blur-xl transition-all duration-300 flex flex-col justify-between
                ${branch.active
                  ? 'bg-slate-900/50 hover:bg-slate-900/80 border-slate-800 hover:border-blue-500/50 cursor-pointer hover:-translate-y-1.5 hover:shadow-[0_20px_40px_-15px_rgba(59,130,246,0.15)] shadow-md'
                  : 'bg-slate-950/20 border-slate-900/80 opacity-70 select-none'
                }`}
            >
              {branch.active && (
                <div className="absolute inset-0 rounded-2xl bg-gradient-to-tr from-blue-500/5 to-indigo-500/5 opacity-0 group-hover:opacity-100 transition-opacity duration-300 pointer-events-none" />
              )}

              <div className="space-y-6 relative z-10">
                {/* Top: Icon & Badge */}
                <div className="flex items-center justify-between">
                  {branch.icon}
                  
                  {branch.active ? (
                    <span className="text-xs font-semibold px-2.5 py-1 rounded-full bg-blue-500/10 border border-blue-500/20 text-blue-400 shadow-sm">
                      {branch.badge}
                    </span>
                  ) : (
                    <span className="inline-flex items-center gap-1.5 text-xs font-semibold px-2.5 py-1 rounded-full bg-slate-800/80 border border-slate-700/60 text-slate-400">
                      <LockIcon /> {branch.badge}
                    </span>
                  )}
                </div>

                {/* Branch Info */}
                <div className="space-y-1.5">
                  <h3 className="text-xl font-bold text-white group-hover:text-blue-400 transition-colors">
                    {branch.title}
                  </h3>
                  <p className="text-xs text-blue-400/80 font-medium">
                    {branch.subtitle}
                  </p>
                  <p className="text-sm text-slate-400 leading-relaxed pt-2">
                    {branch.description}
                  </p>
                </div>

                {/* Features checklist */}
                <ul className="space-y-2 pt-2 border-t border-slate-800/40">
                  {branch.features.map((feature, index) => (
                    <li key={index} className="flex items-start gap-2.5 text-xs text-slate-400">
                      <span className={`text-[10px] translate-y-0.5 ${branch.active ? 'text-blue-400' : 'text-slate-600'}`}>✦</span>
                      <span>{feature}</span>
                    </li>
                  ))}
                </ul>
              </div>

              {/* Action Button */}
              <div className="pt-6 mt-6 border-t border-slate-800/30 flex items-center justify-end relative z-10">
                {branch.active ? (
                  <button className="flex items-center gap-2 text-sm font-bold bg-gradient-to-r from-blue-500 to-indigo-500 hover:from-blue-600 hover:to-indigo-600 text-white px-4 py-2 rounded-xl transition shadow-lg shadow-blue-500/15">
                    Launch Department <ArrowRightIcon />
                  </button>
                ) : (
                  <span className="text-xs font-medium text-slate-600 flex items-center gap-1">
                    Pending Department Config
                  </span>
                )}
              </div>
            </div>
          ))}
        </div>

      </main>

      {/* ── Footer ─────────────────────────────────────────────────────────── */}
      <footer className="relative z-10 text-center py-8 border-t border-slate-800/40 text-xs text-slate-500 bg-slate-950/40">
        RBU Detention Management Portal · Shri Ramdeobaba University, Nagpur <br />
        <span className="opacity-75 mt-1 block">Designed & Configured for Academic Administration © 2026</span>
      </footer>

    </div>
  );
}
