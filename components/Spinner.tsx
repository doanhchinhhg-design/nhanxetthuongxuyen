
import React from 'react';
import { SparklesIcon } from './icons';

export const Spinner: React.FC = () => {
  return (
    <div className="text-center p-8">
      <div className="flex justify-center items-center mb-4">
          <SparklesIcon className="w-10 h-10 text-cyan-500 animate-pulse" />
      </div>
      <h3 className="text-xl font-semibold text-slate-700">AI đang sáng tạo...</h3>
      <p className="text-slate-500 mt-1">Vui lòng chờ trong giây lát.</p>
    </div>
  );
};
