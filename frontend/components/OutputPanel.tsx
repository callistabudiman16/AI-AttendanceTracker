'use client';

interface OutputPanelProps {
  output: string;
  onClear: () => void;
  isLoading: boolean;
}

export default function OutputPanel({ output, onClear, isLoading }: OutputPanelProps) {
  return (
    <div className="card h-full flex flex-col bg-gradient-to-br from-gray-50 to-gray-100 border-2 border-gray-200">
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-indigo-500 to-purple-600 flex items-center justify-center shadow-md">
            <span className="text-white text-lg">💻</span>
          </div>
          <h2 className="text-lg font-bold text-gray-800">Output</h2>
        </div>
        <div className="flex items-center gap-3">
          {isLoading && (
            <div className="flex items-center gap-2 text-sm font-medium text-indigo-600 bg-indigo-100 px-3 py-1.5 rounded-full">
              <div className="w-4 h-4 border-2 border-indigo-600 border-t-transparent rounded-full animate-spin" />
              Processing...
            </div>
          )}
          <button
            onClick={onClear}
            className="btn btn-secondary text-sm bg-white hover:bg-gray-100 border-2 border-gray-300"
          >
            🗑️ Clear
          </button>
        </div>
      </div>
      <div className="flex-1 overflow-auto bg-gradient-to-br from-gray-900 via-gray-800 to-gray-900 text-green-400 p-6 rounded-xl font-mono text-sm whitespace-pre-wrap shadow-inner border-2 border-gray-700">
        {output || (
          <div className="text-gray-500 italic text-center py-8">
            <div className="text-4xl mb-2">✨</div>
            <div>No output yet. Use the actions on the left to get started.</div>
          </div>
        )}
      </div>
    </div>
  );
}

