'use client';

interface StatusBarProps {
  rosterInfo: any;
}

export default function StatusBar({ rosterInfo }: StatusBarProps) {
  return (
    <div className="status-bar">
      <div className="flex items-center gap-4">
        <h1 className="text-2xl font-bold text-white drop-shadow-md">📊 Attendance Tracker</h1>
        <div className="flex items-center gap-2 bg-white/20 backdrop-blur-sm px-3 py-1.5 rounded-full">
          <div className={`w-3 h-3 rounded-full animate-pulse ${rosterInfo?.loaded ? 'bg-green-400 shadow-lg shadow-green-400/50' : 'bg-yellow-400 shadow-lg shadow-yellow-400/50'}`} />
          <span className="text-sm font-medium text-white">
            {rosterInfo?.loaded ? 'Roster Loaded' : 'No Roster'}
          </span>
        </div>
      </div>
      <div className="text-sm text-white/90 bg-white/10 backdrop-blur-sm px-3 py-1.5 rounded-full">
        Backend: http://localhost:5001
      </div>
    </div>
  );
}

