
import React, { useState, useCallback } from 'react';
import { ControlPanel } from './components/ControlPanel';
import { Header } from './components/Header';
import { ResultsTable } from './components/ResultsTable';
import type { Feedback } from './types';
import { generateFeedback } from './services/geminiService';
import { SUBJECTS } from './constants';

const App: React.FC = () => {
  const [subject, setSubject] = useState<string>('Tiếng Việt');
  const [grade, setGrade] = useState<string>('1');
  const [month, setMonth] = useState<string>('9');
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [feedbackList, setFeedbackList] = useState<Feedback[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  const handleGenerate = useCallback(async () => {
    setIsLoading(true);
    setError(null);
    setFeedbackList([]);

    try {
      const comments = await generateFeedback(subject, grade, month, uploadedFile);
      const subjectCode = SUBJECTS.find(s => s.name === subject)?.code || 'NA';
      
      const formattedFeedback = comments.map((comment, index) => ({
        id: `${subjectCode}${grade}.${String(index + 1).padStart(2, '0')}`,
        text: comment,
      }));

      setFeedbackList(formattedFeedback);
    } catch (err) {
      if (err instanceof Error) {
        setError(`Đã xảy ra lỗi: ${err.message}. Vui lòng thử lại.`);
      } else {
        setError('Đã xảy ra một lỗi không xác định. Vui lòng thử lại.');
      }
    } finally {
      setIsLoading(false);
    }
  }, [subject, grade, month, uploadedFile]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800">
      <main className="container mx-auto px-4 py-8 md:py-12">
        <Header />
        <div className="mt-8 bg-white p-6 md:p-8 rounded-2xl shadow-lg shadow-slate-200/50">
          <ControlPanel
            subject={subject}
            setSubject={setSubject}
            grade={grade}
            setGrade={setGrade}
            month={month}
            setMonth={setMonth}
            uploadedFile={uploadedFile}
            setUploadedFile={setUploadedFile}
            onGenerate={handleGenerate}
            isLoading={isLoading}
          />
        </div>
        
        {(isLoading || feedbackList.length > 0 || error) && (
           <div className="mt-10">
              <ResultsTable
                feedbackList={feedbackList}
                isLoading={isLoading}
                error={error}
                fileNameDetails={{ grade, subject, month }}
              />
           </div>
        )}
      </main>
      <footer className="text-center py-4 text-sm text-slate-500">
        <p>&copy; {new Date().getFullYear()} AI Tạo Nhận Xét. Phát triển bởi chuyên gia Frontend & Gemini.</p>
      </footer>
    </div>
  );
};

export default App;
