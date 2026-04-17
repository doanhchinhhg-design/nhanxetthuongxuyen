
import React from 'react';
import type { Feedback } from '../types';
import { exportToExcel } from '../utils/excelUtils';
import { Spinner } from './Spinner';
// FIX: Removed unused DownloadIcon import that was causing a build error.
import { FileSpreadsheetIcon } from './icons';

interface ResultsTableProps {
  feedbackList: Feedback[];
  isLoading: boolean;
  error: string | null;
  fileNameDetails: {
    grade: string;
    subject: string;
    month: string;
  };
}

export const ResultsTable: React.FC<ResultsTableProps> = ({ feedbackList, isLoading, error, fileNameDetails }) => {
  const handleExport = () => {
    const { grade, subject, month } = fileNameDetails;
    const monthName = `Thang_${month}`;
    const subjectName = subject.replace(/\s/g, '_');
    const fileName = `Nhan_xet_Lop_${grade}_${subjectName}_${monthName}.xlsx`;
    exportToExcel(feedbackList, fileName);
  };
  
  const { grade, subject, month } = fileNameDetails;

  const headerTitle = `Kết quả cho Lớp ${grade}, môn ${subject}, Tháng ${month}`;

  if (isLoading) {
    return <Spinner />;
  }

  if (error) {
    return (
      <div className="text-center p-8 bg-red-50 text-red-700 rounded-lg">
        <h3 className="text-lg font-semibold">Lỗi</h3>
        <p>{error}</p>
      </div>
    );
  }

  if (feedbackList.length === 0) {
    return null;
  }

  return (
    <div className="bg-white p-6 md:p-8 rounded-2xl shadow-lg shadow-slate-200/50">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6">
        <h2 className="text-2xl font-bold text-slate-800 mb-3 sm:mb-0">{headerTitle}</h2>
        <button
          onClick={handleExport}
          className="inline-flex items-center justify-center px-4 py-2 border border-transparent text-sm font-medium rounded-md shadow-sm text-white bg-emerald-600 hover:bg-emerald-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-emerald-500 transition-colors"
        >
          <FileSpreadsheetIcon className="w-5 h-5 mr-2" />
          Xuất file Excel
        </button>
      </div>
      <div className="overflow-x-auto">
        <table className="min-w-full divide-y divide-slate-200">
          <thead className="bg-slate-50">
            <tr>
              <th scope="col" className="px-6 py-3 text-left text-xs font-bold text-slate-600 uppercase tracking-wider w-1/4 md:w-1/5">
                Mã nhận xét
              </th>
              <th scope="col" className="px-6 py-3 text-left text-xs font-bold text-slate-600 uppercase tracking-wider">
                Nội dung nhận xét
              </th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-slate-200">
            {feedbackList.map((feedback) => (
              <tr key={feedback.id} className="hover:bg-slate-50 transition-colors">
                <td className="px-6 py-4 whitespace-nowrap text-sm font-mono text-cyan-600">{feedback.id}</td>
                <td className="px-6 py-4 text-sm text-slate-700">{feedback.text}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};
