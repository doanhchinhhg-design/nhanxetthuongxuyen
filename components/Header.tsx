
import React from 'react';
import { BrainCircuitIcon } from './icons';

export const Header: React.FC = () => {
  return (
    <header className="text-center">
      <div className="inline-block bg-gradient-to-r from-teal-400 to-cyan-500 p-3 rounded-2xl mb-4 shadow-lg shadow-cyan-500/20">
        <BrainCircuitIcon className="h-10 w-10 text-white" />
      </div>
      <h1 className="text-4xl md:text-5xl font-extrabold text-slate-900 tracking-tight">
        Trợ lý AI <span className="bg-gradient-to-r from-teal-400 to-cyan-500 bg-clip-text text-transparent">Nhận Xét Học Sinh</span>
      </h1>
      <p className="mt-4 max-w-2xl mx-auto text-lg text-slate-600">
        Tạo 50 lời nhận xét chuyên nghiệp, phù hợp với từng học sinh theo chuẩn Bộ GD&ĐT chỉ trong vài giây.
      </p>
    </header>
  );
};
