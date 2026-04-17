
import React, { useCallback } from 'react';
import { SUBJECTS, GRADES, MONTHS } from '../constants';
import { BookOpenIcon, CalendarDaysIcon, GraduationCapIcon, UploadCloudIcon, PaperclipIcon, XCircleIcon, SparklesIcon } from './icons';

interface ControlPanelProps {
  subject: string;
  setSubject: (value: string) => void;
  grade: string;
  setGrade: (value: string) => void;
  month: string;
  setMonth: (value: string) => void;
  uploadedFile: File | null;
  setUploadedFile: (file: File | null) => void;
  onGenerate: () => void;
  isLoading: boolean;
}

const SelectInput: React.FC<{
  label: string;
  value: string;
  onChange: (e: React.ChangeEvent<HTMLSelectElement>) => void;
  options: { name: string; value: string }[];
  icon: React.ReactNode;
}> = ({ label, value, onChange, options, icon }) => (
  <div className="relative">
    <label className="block text-sm font-medium text-slate-600 mb-1">{label}</label>
    <div className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none pt-6">
      {icon}
    </div>
    <select
      value={value}
      onChange={onChange}
      className="w-full pl-10 pr-4 py-2.5 bg-slate-100 border border-slate-200 rounded-lg focus:ring-2 focus:ring-cyan-400 focus:border-cyan-400 transition"
    >
      {options.map((opt) => (
        <option key={opt.value} value={opt.value}>
          {opt.name}
        </option>
      ))}
    </select>
  </div>
);

const FileUpload: React.FC<{
  uploadedFile: File | null;
  setUploadedFile: (file: File | null) => void;
}> = ({ uploadedFile, setUploadedFile }) => {
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setUploadedFile(e.target.files[0]);
    }
  };

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      setUploadedFile(e.dataTransfer.files[0]);
    }
  };

  if (uploadedFile) {
    return (
      <div className="mt-1 flex justify-center items-center px-6 pt-5 pb-6 border-2 border-slate-300 border-dashed rounded-lg bg-slate-50">
        <div className="text-center">
          <PaperclipIcon className="mx-auto h-10 w-10 text-slate-500" />
          <p className="mt-1 text-sm text-slate-600 font-medium">{uploadedFile.name}</p>
          <button
            onClick={() => setUploadedFile(null)}
            className="mt-2 inline-flex items-center px-2.5 py-1.5 border border-transparent text-xs font-medium rounded text-red-700 bg-red-100 hover:bg-red-200 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-red-500"
          >
            <XCircleIcon className="w-4 h-4 mr-1" />
            Xóa tệp
          </button>
        </div>
      </div>
    );
  }

  return (
    <div onDragOver={handleDragOver} onDrop={handleDrop}>
        <label className="block text-sm font-medium text-slate-600 mb-1">Tài liệu đính kèm (tùy chọn)</label>
        <div className="mt-1 flex justify-center px-6 pt-5 pb-6 border-2 border-slate-300 border-dashed rounded-lg hover:border-cyan-400 transition bg-white">
            <div className="space-y-1 text-center">
                <UploadCloudIcon className="mx-auto h-12 w-12 text-slate-400" />
                <div className="flex text-sm text-slate-600">
                <label
                    htmlFor="file-upload"
                    className="relative cursor-pointer bg-white rounded-md font-medium text-cyan-600 hover:text-cyan-500 focus-within:outline-none focus-within:ring-2 focus-within:ring-offset-2 focus-within:ring-cyan-500"
                >
                    <span>Tải tệp lên</span>
                    <input id="file-upload" name="file-upload" type="file" className="sr-only" onChange={handleFileChange} accept="image/*,.pdf,.xlsx,.xls" />
                </label>
                <p className="pl-1">hoặc kéo và thả</p>
                </div>
                <p className="text-xs text-slate-500">Ảnh, PDF hoặc Excel (.xlsx, .xls)</p>
            </div>
        </div>
    </div>
  );
};


export const ControlPanel: React.FC<ControlPanelProps> = ({
  subject,
  setSubject,
  grade,
  setGrade,
  month,
  setMonth,
  uploadedFile,
  setUploadedFile,
  onGenerate,
  isLoading,
}) => {
  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <SelectInput 
            label="Môn học" 
            value={subject} 
            onChange={(e) => setSubject(e.target.value)} 
            options={SUBJECTS.map(s => ({name: s.name, value: s.name}))}
            icon={<BookOpenIcon className="w-5 h-5 text-slate-400" />}
        />
        <SelectInput 
            label="Lớp" 
            value={grade} 
            onChange={(e) => setGrade(e.target.value)} 
            options={GRADES.map(g => ({name: `Lớp ${g}`, value: g}))}
            icon={<GraduationCapIcon className="w-5 h-5 text-slate-400" />}
        />
        <SelectInput 
            label="Tháng" 
            value={month} 
            onChange={(e) => setMonth(e.target.value)} 
            options={MONTHS.map(m => ({ name: m.name, value: m.value }))}
            icon={<CalendarDaysIcon className="w-5 h-5 text-slate-400" />}
        />
      </div>
      <FileUpload uploadedFile={uploadedFile} setUploadedFile={setUploadedFile} />
      <div className="pt-4">
        <button
          onClick={onGenerate}
          disabled={isLoading}
          className="w-full flex items-center justify-center text-lg font-semibold text-white bg-gradient-to-r from-teal-400 to-cyan-500 hover:from-teal-500 hover:to-cyan-600 disabled:opacity-50 disabled:cursor-not-allowed rounded-xl px-8 py-4 shadow-lg shadow-cyan-500/20 hover:shadow-xl hover:shadow-cyan-500/30 transform hover:-translate-y-0.5 transition-all duration-300"
        >
          {isLoading ? (
            <>
              <svg className="animate-spin -ml-1 mr-3 h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
              </svg>
              Đang xử lý...
            </>
          ) : (
            <>
                <SparklesIcon className="w-6 h-6 mr-2"/>
                Tạo nhận xét
            </>
          )}
        </button>
      </div>
    </div>
  );
};
