'use client';

import { useState, useEffect } from 'react';
import StatusBar from '@/components/StatusBar';
import RosterInfo from '@/components/RosterInfo';
import ActionPanel from '@/components/ActionPanel';
import OutputPanel from '@/components/OutputPanel';

export default function Home() {
  const [rosterInfo, setRosterInfo] = useState<any>(null);
  const [output, setOutput] = useState<string>('');
  const [isLoading, setIsLoading] = useState(false);

  useEffect(() => {
    loadRosterInfo();
  }, []);

  const loadRosterInfo = async () => {
    try {
      const response = await fetch('http://localhost:5001/api/roster/info');
      const data = await response.json();
      if (data.success) {
        setRosterInfo(data);
      }
    } catch (error) {
      console.error('Error loading roster info:', error);
    }
  };

  const addOutput = (message: string) => {
    setOutput((prev) => prev + '\n' + message);
  };

  const clearOutput = () => {
    setOutput('');
  };

  return (
    <div className="min-h-screen flex flex-col">
      <StatusBar rosterInfo={rosterInfo} />
      
      <div className="flex-1 flex gap-4 p-4">
        <div className="w-80 space-y-4">
          <RosterInfo rosterInfo={rosterInfo} onRefresh={loadRosterInfo} />
          <ActionPanel
            onOutput={addOutput}
            onLoading={setIsLoading}
            onRosterUpdate={loadRosterInfo}
          />
        </div>
        
        <div className="flex-1">
          <OutputPanel output={output} onClear={clearOutput} isLoading={isLoading} />
        </div>
      </div>
    </div>
  );
}

