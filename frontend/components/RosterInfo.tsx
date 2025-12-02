'use client';

interface RosterInfoProps {
  rosterInfo: any;
  onRefresh: () => void;
}

export default function RosterInfo({ rosterInfo, onRefresh }: RosterInfoProps) {
  if (!rosterInfo?.loaded) {
    return (
      <div className="card bg-gradient-to-br from-gray-50 to-gray-100 border-2 border-dashed border-gray-300">
        <div className="flex items-center gap-2 mb-2">
          <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-400 to-indigo-500 flex items-center justify-center">
            <span className="text-white text-lg">📋</span>
          </div>
          <h2 className="text-lg font-bold text-gray-700">Roster Information</h2>
        </div>
        <p className="text-sm text-gray-500 italic">No roster loaded</p>
      </div>
    );
  }

  return (
    <div className="card bg-gradient-to-br from-blue-50 via-indigo-50 to-purple-50 border-2 border-blue-200">
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-500 to-indigo-600 flex items-center justify-center shadow-md">
            <span className="text-white text-lg">📋</span>
          </div>
          <h2 className="text-lg font-bold text-gray-800">Roster Information</h2>
        </div>
        <button
          onClick={onRefresh}
          className="text-sm font-medium text-blue-600 hover:text-blue-700 bg-blue-100 hover:bg-blue-200 px-3 py-1 rounded-lg transition-colors"
        >
          🔄 Refresh
        </button>
      </div>
      <div className="space-y-3 text-sm">
        <div className="flex items-center justify-between p-2 bg-white/60 rounded-lg">
          <span className="font-semibold text-gray-700">Students:</span>
          <span className="text-blue-600 font-bold text-base">{rosterInfo.student_count}</span>
        </div>
        <div className="flex items-center justify-between p-2 bg-white/60 rounded-lg">
          <span className="font-semibold text-gray-700">Date Columns:</span>
          <span className="text-indigo-600 font-bold text-base">{rosterInfo.date_columns?.length || 0}</span>
        </div>
        <div className="p-2 bg-white/60 rounded-lg">
          <span className="font-semibold text-gray-700 block mb-1">File:</span>
          <span className="text-gray-600 text-xs break-all">{rosterInfo.roster_file}</span>
        </div>
      </div>
    </div>
  );
}

