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

const YearOneIcon = () => (
  <div className="w-12 h-12 rounded-xl bg-gradient-to-tr from-blue-500 to-indigo-500 flex items-center justify-center text-white shadow-lg shadow-blue-500/20">
    <svg xmlns="http://www.w3.org/2000/svg" className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253" />
    </svg>
  </div>
);

const YearTwoIcon = () => (
  <div className="w-12 h-12 rounded-xl bg-gradient-to-tr from-purple-500 to-pink-500 flex items-center justify-center text-white shadow-lg shadow-purple-500/20">
    <svg xmlns="http://www.w3.org/2000/svg" className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M9.75 17L9 20l-1 1h8l-1-1-.75-3M3 13h18M5 17h14a2 2 0 002-2V5a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
    </svg>
  </div>
);


const YearThreeIcon = () => (
  <div className="w-12 h-12 rounded-xl bg-gradient-to-tr from-teal-500 to-emerald-500 flex items-center justify-center text-white shadow-lg shadow-teal-500/20">
    <svg xmlns="http://www.w3.org/2000/svg" className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M19.428 15.428a2 2 0 00-1.022-.547l-2.387-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z" />
    </svg>
  </div>
);

const YearFourIcon = () => (
  <div className="w-12 h-12 rounded-xl bg-gradient-to-tr from-rose-500 to-red-500 flex items-center justify-center text-white shadow-lg shadow-rose-500/20">
    <svg xmlns="http://www.w3.org/2000/svg" className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 14l9-5-9-5-9 5 9 5zm0 0l6.16-3.422a12.083 12.083 0 01.665 6.479A11.952 11.952 0 0012 20.055a11.952 11.952 0 00-6.824-2.998 12.078 12.078 0 01.665-6.479L12 14zm-4 6v-7.5l4-2.222" />
    </svg>
  </div>
);

export default function LandingPage({ onSelectYear }) {
  const years = [
    {
      id: '1',
      title: 'First Year',
      subtitle: 'B.Tech / B.E. Core Courses',
      description: 'Fully automated report builder featuring attendance aggregation, Theory & Practical average processing, and condonation generation.',
      active: true,
      features: [
        'Merged batch-wise report creation',
        'Automatic Theory + Practical avg checker',
        'Condonation seat number processing',
        'Word (.docx) Detention Tables builder'
      ],
      icon: <YearOneIcon />,
      color: 'from-blue-600 to-indigo-600',
      badge: 'Active Module'
    },
    {
      id: '2',
      title: 'Second Year',
      subtitle: 'Core Departments & Electives',
      description: 'Fully automated 4th Semester detention report builder supporting Excel processing with advanced component-wise average checking.',
      active: true,
      features: [
        'Merged division/section report creation',
        'Theory & Practical average logic filter',
        'Condonation roll list processor',
        'Refined Table III(A) format'
      ],
      icon: <YearTwoIcon />,
      color: 'from-purple-600 to-pink-600',
      badge: 'Active Module'
    },

    {
      id: '3',
      title: 'Third Year',
      subtitle: 'Electronics & Computer Science, VI Sem',
      description: 'Fully automated 6th Semester ECS detention report builder with student-wise Table III, 19-course Semester VI mapping and per-section merging.',
      active: true,
      features: [
        'Merged section-wise report (VI A, VI B ...)',
        'Student-wise detention Table III',
        'Deep Learning, VLSI, NLP, DHV course support',
        'Word (.docx) Detention Tables builder'
      ],
      icon: <YearThreeIcon />,
      color: 'from-teal-500 to-emerald-500',
      badge: 'Active Module'
    },
    {
      id: '4',
      title: 'Fourth Year',
      subtitle: 'Final Year & Electives, VIII Sem',
      description: 'Fully automated 8th Semester ECS detention report builder supporting Excel uploads, smart elective handling, and official 5-column Table III document layout.',
      active: true,
      features: [
        'Merged division/section report creation',
        'Smart elective enrollment logic',
        'Official 5-column Table III layout',
        'Word (.docx) Detention Tables builder'
      ],
      icon: <YearFourIcon />,
      color: 'from-rose-600 to-red-600',
      badge: 'Active Module'

    }
  ];

  return (
    <div className="min-h-screen bg-[#0f172a] text-slate-100 flex flex-col relative overflow-hidden font-sans selection:bg-indigo-500/30">
      
      {/* ── Ambient Glow Background Elements ────────────────────────────────────── */}
      <div className="absolute top-[-10%] left-[-10%] w-[50%] h-[50%] rounded-full bg-blue-500/10 blur-[120px] pointer-events-none" />
      <div className="absolute bottom-[-10%] right-[-10%] w-[50%] h-[50%] rounded-full bg-purple-500/10 blur-[120px] pointer-events-none" />
      <div className="absolute top-[30%] right-[15%] w-[300px] h-[300px] rounded-full bg-indigo-500/5 blur-[90px] pointer-events-none" />

      {/* ── Header ─────────────────────────────────────────────────────────── */}
      <header className="relative z-10 border-b border-slate-800/80 bg-slate-900/40 backdrop-blur-md">
        <div className="max-w-7xl mx-auto px-6 py-5 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 to-indigo-600 flex items-center justify-center text-white font-bold text-lg shadow-lg shadow-blue-500/20">
              R
            </div>
            <div>
              <h1 className="text-md font-bold tracking-tight text-white">RBU Detention System</h1>
              <p className="text-[11px] text-slate-400 font-medium">Shri Ramdeobaba University, Nagpur</p>
            </div>
          </div>
          
          <div className="flex items-center gap-2 px-3 py-1.5 rounded-full bg-slate-800/60 border border-slate-700/40 text-xs font-semibold text-slate-300">
            <span className="w-2 h-2 rounded-full bg-emerald-500 animate-ping" />
            <span className="text-slate-200">Session 2025-26</span>
          </div>
        </div>
      </header>

      {/* ── Main Content ───────────────────────────────────────────────────── */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-6 py-12 lg:py-20 relative z-10 flex flex-col justify-center items-center">
        
        {/* Hero Headline */}
        <div className="text-center max-w-3xl mb-16 space-y-4">
          <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-indigo-500/10 border border-indigo-500/20 text-xs font-semibold text-indigo-300">
            <SparklesIcon /> Academic Administration Portal
          </div>
          <h2 className="text-4xl sm:text-5xl font-extrabold tracking-tight text-white leading-tight">
            Academic Detention <br className="hidden sm:inline" />
            <span className="bg-gradient-to-r from-blue-400 via-indigo-400 to-purple-400 bg-clip-text text-transparent">
              Report Management
            </span>
          </h2>
          <p className="text-base text-slate-400 max-w-2xl mx-auto leading-relaxed">
            Select the target academic year to proceed with detention calculation, attendance file aggregation, and official Word document generation.
          </p>
        </div>

        {/* Years Card Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 lg:gap-8 w-full max-w-5xl">
          {years.map((year) => (
            <div
              key={year.id}
              onClick={() => year.active && onSelectYear(year.id)}
              className={`group relative rounded-2xl p-6 lg:p-8 border backdrop-blur-xl transition-all duration-300 flex flex-col justify-between
                ${year.active
                  ? 'bg-slate-900/50 hover:bg-slate-900/80 border-slate-800 hover:border-indigo-500/50 cursor-pointer hover:-translate-y-1.5 hover:shadow-[0_20px_40px_-15px_rgba(59,130,246,0.15)] shadow-md'
                  : 'bg-slate-950/20 border-slate-900/80 opacity-70 select-none'
                }`}
            >
              {/* Outer Card Glow Border for Hover (only for active card) */}
              {year.active && (
                <div className="absolute inset-0 rounded-2xl bg-gradient-to-tr from-blue-500/5 to-indigo-500/5 opacity-0 group-hover:opacity-100 transition-opacity duration-300 pointer-events-none" />
              )}

              <div className="space-y-6 relative z-10">
                {/* Card Top: Icon & Badge */}
                <div className="flex items-center justify-between">
                  {year.icon}
                  
                  {year.active ? (
                    <span className="text-xs font-semibold px-2.5 py-1 rounded-full bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 shadow-sm">
                      {year.badge}
                    </span>
                  ) : (
                    <span className="inline-flex items-center gap-1.5 text-xs font-semibold px-2.5 py-1 rounded-full bg-slate-800/80 border border-slate-700/60 text-slate-400">
                      <LockIcon /> {year.badge}
                    </span>
                  )}
                </div>

                {/* Card Title & Info */}
                <div className="space-y-1.5">
                  <h3 className="text-xl font-bold text-white group-hover:text-indigo-400 transition-colors">
                    {year.title}
                  </h3>
                  <p className="text-xs text-indigo-400/80 font-medium">
                    {year.subtitle}
                  </p>
                  <p className="text-sm text-slate-400 leading-relaxed pt-2">
                    {year.description}
                  </p>
                </div>

                {/* Features Checklist */}
                <ul className="space-y-2 pt-2 border-t border-slate-800/40">
                  {year.features.map((feature, index) => (
                    <li key={index} className="flex items-start gap-2.5 text-xs text-slate-400">
                      <span className={`text-[10px] translate-y-0.5 ${year.active ? 'text-indigo-400' : 'text-slate-600'}`}>✦</span>
                      <span>{feature}</span>
                    </li>
                  ))}
                </ul>
              </div>

              {/* Action Button at bottom */}
              <div className="pt-6 mt-6 border-t border-slate-800/30 flex items-center justify-end relative z-10">
                {year.active ? (
                  <button className="flex items-center gap-2 text-sm font-bold bg-gradient-to-r from-blue-500 to-indigo-500 hover:from-blue-600 hover:to-indigo-600 text-white px-4 py-2 rounded-xl transition shadow-lg shadow-blue-500/15">
                    Launch Module <ArrowRightIcon />
                  </button>
                ) : (
                  <span className="text-xs font-medium text-slate-600 flex items-center gap-1">
                    Pending Backend Configuration
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
