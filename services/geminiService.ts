
import { GoogleGenAI, Type } from "@google/genai";
import { fileToBase64 } from "../utils/fileUtils";
import { readExcelFile } from "../utils/excelUtils";

const API_KEY = process.env.API_KEY;

if (!API_KEY) {
  throw new Error("API_KEY environment variable is not set.");
}

const ai = new GoogleGenAI({ apiKey: API_KEY });

const buildPrompt = (subject: string, grade: string, month: string, hasFile: boolean, excelContent?: string) => {
  let fileContext = '';
  if (hasFile) {
    if (excelContent) {
      fileContext = `- Ngữ cảnh bổ sung (Dữ liệu từ Excel): \n${excelContent}\n\nHãy phân tích dữ liệu trên để đưa ra các nhận xét cụ thể cho từng học sinh hoặc nhóm học sinh dựa trên bối cảnh học tập.`;
    } else {
      fileContext = `- Ngữ cảnh bổ sung: Phân tích tệp đính kèm (bài làm hoặc bảng điểm của học sinh) để đưa ra nhận xét cá nhân hóa và phù hợp.`;
    }
  }

  return `
    Bạn là một trợ lý AI chuyên tạo lời nhận xét cho học sinh tiểu học Việt Nam, tuân thủ nghiêm ngặt Thông tư 27/2020 và 32/2018.

    Yêu cầu cụ thể:
    - Môn học: ${subject}
    - Lớp: ${grade}
    - Bối cảnh thời gian: Tháng ${month} của năm học.
    - Sách giáo khoa tham khảo: "Kết nối tri thức với cuộc sống".
    ${fileContext}

    Nhiệm vụ: Tạo ra chính xác 50 lời nhận xét ĐỘC ĐÁO.

    Các quy tắc cho mỗi nhận xét:
    1. Ngắn gọn, súc tích (tối đa 1-2 câu).
    2. Sử dụng ngôn ngữ sư phạm, chuẩn mực.
    3. TUYỆT ĐỐI KHÔNG cho điểm, KHÔNG xếp loại, KHÔNG so sánh với học sinh khác.
    4. Tập trung vào sự tiến bộ, nỗ lực và mang tính động viên, khích lệ.
    5. Nội dung phải phù hợp với chương trình học của Lớp ${grade}, môn ${subject} trong Tháng ${month} theo sách "Kết nối tri thức với cuộc sống".
    6. Các nhận xét không được trùng lặp ý tưởng.
    7. TRÁNH sử dụng các cấu trúc câu rập khuôn, máy móc theo barem đánh giá như "hoàn thành nội dung", "đạt yêu cầu cần đạt", "vận dụng được kiến thức". Hãy diễn đạt một cách tự nhiên và tập trung vào hành vi, thái độ cụ thể của học sinh.
    8. TUYỆT ĐỐI KHÔNG dùng các từ xưng hô như "em" hoặc "con". Hãy nhận xét trực tiếp về hành động, sản phẩm học tập hoặc sự tiến bộ của học sinh.

    Định dạng đầu ra: Trả về một đối tượng JSON duy nhất có khóa là "comments" chứa một mảng gồm chính xác 50 chuỗi ký tự (string), mỗi chuỗi là một lời nhận xét.
    Ví dụ JSON: { "comments": ["Nhận xét 1...", "Nhận xét 2...", ..., "Nhận xét 50..."] }
  `;
};

export const generateFeedback = async (
  subject: string,
  grade: string,
  month: string,
  file: File | null
): Promise<string[]> => {
  let excelContent = '';
  const isExcel = file && (
    file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
    file.type === 'application/vnd.ms-excel' ||
    file.name.endsWith('.xlsx') ||
    file.name.endsWith('.xls')
  );

  if (file && isExcel) {
    excelContent = await readExcelFile(file);
  }

  const prompt = buildPrompt(subject, grade, month, !!file, excelContent);
  const model = 'gemini-3-flash-preview';
  
  const parts: ({ text: string } | { inlineData: { mimeType: string; data: string } })[] = [{ text: prompt }];
  
  if (file && !isExcel) {
    const base64Data = await fileToBase64(file);
    parts.push({
      inlineData: {
        mimeType: file.type || 'application/pdf', // Default to PDF if type is missing but isExcel is false
        data: base64Data,
      },
    });
  }

  try {
    const response = await ai.models.generateContent({
      model: model,
      contents: { parts: parts },
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            comments: {
              type: Type.ARRAY,
              items: {
                type: Type.STRING
              }
            }
          }
        },
        temperature: 0.8,
      }
    });

    if (response.text) {
        const jsonResponse = JSON.parse(response.text);
        if (jsonResponse.comments && Array.isArray(jsonResponse.comments)) {
            return jsonResponse.comments;
        }
    }
    
    throw new Error('Định dạng phản hồi từ AI không hợp lệ.');
  } catch (error) {
    console.error("Error calling Gemini API:", error);
    throw new Error('Không thể tạo nhận xét từ AI. Vui lòng kiểm tra lại cấu hình và thử lại.');
  }
};
