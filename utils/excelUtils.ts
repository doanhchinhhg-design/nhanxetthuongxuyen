
import * as XLSX from 'xlsx';
import type { Feedback } from '../types';

export const readExcelFile = (file: File): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        let fullText = '';
        
        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const text = XLSX.utils.sheet_to_txt(worksheet);
          fullText += `--- Trang: ${sheetName} ---\n${text}\n\n`;
        });
        
        resolve(fullText);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
};

export const exportToExcel = (data: Feedback[], fileName: string): void => {
  const worksheetData = data.map(item => ({
    "Mã nhận xét": item.id,
    "Nội dung nhận xét": item.text,
  }));

  const worksheet = XLSX.utils.json_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "NhanXet");

  // Set column widths
  worksheet['!cols'] = [{ wch: 20 }, { wch: 80 }];

  XLSX.writeFile(workbook, fileName);
};
