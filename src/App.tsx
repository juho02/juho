/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useRef } from "react";
import { GoogleGenAI, Type } from "@google/genai";
import { FileText, Upload, Loader2, CheckCircle2, AlertCircle, RefreshCw, Table, Download } from "lucide-react";
import Markdown from "react-markdown";
import ExcelJS from "exceljs";
import * as pdfjsLib from "pdfjs-dist";
// @ts-ignore
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

interface PatentInfo {
  id: string;
  title: string;
  category: string;
  rightType: string;
  filingDate: string;
  applicationNumber: string;
  registrationDate: string;
  registrationNumber: string;
  status: string;
  inventor: string;
  applicant: string;
  summary: string;
  note: string;
  drawingPages: number[];
}

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [targetExcel, setTargetExcel] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [loadingImages, setLoadingImages] = useState(false);
  const [loadingDetails, setLoadingDetails] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [result, setResult] = useState<PatentInfo | null>(null);
  const [extractedImages, setExtractedImages] = useState<{ data: string, pageNum: number, width: number, height: number }[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile && selectedFile.type === "application/pdf") {
      setFile(selectedFile);
      setError(null);
      setResult(null);
      
      const url = URL.createObjectURL(selectedFile);
      setPreviewUrl(url);
    } else if (selectedFile) {
      setError("PDF 파일만 업로드 가능합니다.");
      setFile(null);
      setPreviewUrl(null);
    }
  };

  const handleExcelChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile && (selectedFile.name.endsWith(".xlsx") || selectedFile.name.endsWith(".xls"))) {
      setTargetExcel(selectedFile);
      setError(null);
    } else if (selectedFile) {
      setError("엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.");
    }
  };

  const getPdfPages = async (file: File, pageNumbers?: number[]): Promise<{ data: string, pageNum: number, width: number, height: number }[]> => {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const pages: { data: string, pageNum: number, width: number, height: number }[] = [];
      
      const targetPages = pageNumbers && pageNumbers.length > 0 
        ? pageNumbers.filter(n => n >= 1 && n <= pdf.numPages)
        : Array.from({ length: pdf.numPages }, (_, i) => i + 1);

      for (const pageNum of targetPages) {
        const page = await pdf.getPage(pageNum);
        const viewport = page.getViewport({ scale: 1.5 });
        const canvas = document.createElement("canvas");
        const context = canvas.getContext("2d");
        if (!context) continue;

        canvas.height = viewport.height;
        canvas.width = viewport.width;

        // @ts-ignore - pdfjs-dist v5 types might be slightly different
        await page.render({ canvasContext: context, viewport }).promise;
        pages.push({ data: canvas.toDataURL("image/png"), pageNum, width: canvas.width, height: canvas.height });
      }
      return pages;
    } catch (err) {
      console.error("PDF pages extraction error:", err);
      return [];
    }
  };

  const exportToExcel = async () => {
    if (!result || !file) return;
    setExporting(true);
    setError(null);

    try {
      const workbook = new ExcelJS.Workbook();
      
      if (targetExcel) {
        const arrayBuffer = await targetExcel.arrayBuffer();
        await workbook.xlsx.load(arrayBuffer);
      }

      // 1. "특허 리스트" Sheet
      const listSheetName = "특허 리스트";
      let listSheet = workbook.getWorksheet(listSheetName);
      if (!listSheet) {
        listSheet = workbook.addWorksheet(listSheetName);
        const headerRow = listSheet.getRow(1);
        headerRow.getCell(4).value = "관리번호";
        headerRow.getCell(5).value = "특허 등록 명칭";
        headerRow.getCell(8).value = "출원일";
        headerRow.getCell(9).value = "출원번호";
        headerRow.getCell(13).value = "발명자";
        headerRow.getCell(14).value = "출원인";
        headerRow.commit();
      }

      let lastListRow = 1;
      listSheet.eachRow((row, rowNumber) => {
        if (rowNumber > lastListRow) lastListRow = rowNumber;
      });
      
      const nextListRowNumber = lastListRow + 1;
      const listRow = listSheet.getRow(nextListRowNumber);

      // 관리번호 업카운트 로직 (D열 = 4번째 열)
      let prevMgmtNum = "";
      if (lastListRow > 1) {
        const prevCell = listSheet.getRow(lastListRow).getCell(4).value;
        if (prevCell !== null && prevCell !== undefined) {
          if (typeof prevCell === 'object' && 'result' in prevCell) {
            prevMgmtNum = String(prevCell.result);
          } else {
            prevMgmtNum = String(prevCell);
          }
        }
      }

      let newMgmtNum = "1"; // 기본값
      if (prevMgmtNum) {
        const match = prevMgmtNum.match(/(\d+)$/);
        if (match) {
          const numStr = match[1];
          const num = parseInt(numStr, 10) + 1;
          const paddedNum = num.toString().padStart(numStr.length, '0');
          newMgmtNum = prevMgmtNum.substring(0, prevMgmtNum.length - numStr.length) + paddedNum;
        } else {
          newMgmtNum = prevMgmtNum + "-1";
        }
      }

      listRow.getCell(4).value = newMgmtNum;
      listRow.getCell(5).value = result.title;
      listRow.getCell(8).value = result.filingDate;
      listRow.getCell(9).value = result.applicationNumber;
      listRow.getCell(13).value = result.inventor;
      listRow.getCell(14).value = result.applicant;
      listRow.commit();

      // 2. "특허 내용" Sheet
      const contentSheetName = "특허 내용";
      let contentSheet = workbook.getWorksheet(contentSheetName);
      if (!contentSheet) {
        contentSheet = workbook.addWorksheet(contentSheetName);
      }

      // Find the next available block for "특허 내용"
      // Each block is roughly 4 rows + spacing
      let nextContentRow = 1;
      contentSheet.eachRow((row, rowNumber) => {
        if (rowNumber >= nextContentRow) nextContentRow = rowNumber + 2;
      });

      // Template Header Row
      const headers = ["ID", "특허 등록 명칭", "구분", "권리", "출원일", "출원번호", "등록일", "등록번호", "현재상태", "발명자", "출원인", "비고"];
      const headerRow = contentSheet.getRow(nextContentRow);
      headers.forEach((h, i) => {
        const cell = headerRow.getCell(i + 2);
        cell.value = h;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { bold: true, size: 9 };
      });
      headerRow.height = 24.5;

      // Data Row
      const dataRow = contentSheet.getRow(nextContentRow + 1);
      const dataValues = [result.id, result.title, result.category, result.rightType, result.filingDate, result.applicationNumber, result.registrationDate, result.registrationNumber, result.status, result.inventor, result.applicant, result.note];
      dataValues.forEach((v, i) => {
        const cell = dataRow.getCell(i + 2);
        cell.value = v;
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.font = { size: 9 };
      });
      dataRow.height = 24.5;

      // Summary Row
      const summaryRow = contentSheet.getRow(nextContentRow + 2);
      const summaryLabelCell = summaryRow.getCell(2);
      summaryLabelCell.value = "요약";
      summaryLabelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
      summaryLabelCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      summaryLabelCell.alignment = { horizontal: 'center', vertical: 'middle' };
      summaryLabelCell.font = { bold: true, size: 9 };

      const summaryTextCell = summaryRow.getCell(3);
      summaryTextCell.value = result.summary;
      contentSheet.mergeCells(nextContentRow + 2, 3, nextContentRow + 2, 13);
      for (let i = 3; i <= 13; i++) {
        summaryRow.getCell(i).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      }
      summaryTextCell.alignment = { vertical: 'middle', wrapText: true };
      summaryTextCell.font = { size: 9 };
      summaryRow.height = 24.5;

      // Image Gallery Template
      const pages = extractedImages.length > 0 
        ? extractedImages 
        : await getPdfPages(file, result.drawingPages);
      let currentRow = nextContentRow + 3;
      
      if (pages.length > 0) {
        // Drawing Content (Image Area)
        const imageContentRow = contentSheet.getRow(currentRow);
        const imageLabelCell = imageContentRow.getCell(2);
        imageLabelCell.value = "이미지";
        imageLabelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
        imageLabelCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        imageLabelCell.alignment = { horizontal: 'center', vertical: 'middle' };
        imageLabelCell.font = { bold: true, size: 9 };
        
        contentSheet.mergeCells(currentRow, 3, currentRow, 13);
        for (let j = 3; j <= 13; j++) {
          imageContentRow.getCell(j).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        }
        imageContentRow.height = 340;

        for (let i = 0; i < pages.length; i++) {
          const pageInfo = pages[i];
          const imageId = workbook.addImage({
            base64: pageInfo.data.split(',')[1],
            extension: 'png',
          });
          
          // Force drawing initialization if it's a loaded workbook
          // @ts-ignore
          if (contentSheet.model && !contentSheet.model.drawings) {
            // @ts-ignore
            contentSheet.getImages();
          }

          try {
            const offset = i * 0.1; // 약간씩 겹치게 배치
            contentSheet.addImage(imageId, {
              tl: { col: 2 + offset, row: currentRow - 1 + offset },
              ext: { width: pageInfo.width, height: pageInfo.height },
              editAs: 'oneCell'
            });
          } catch (e) {
            console.warn("Retrying addImage with string range positioning...");
            // Fallback to string range if object positioning fails
            contentSheet.addImage(imageId, `C${currentRow}:L${currentRow + 1}`);
          }
        }
        
        currentRow++; // Move to next block
        contentSheet.getRow(currentRow).height = 10; // Spacing
        currentRow++;
      } else {
        // Fallback if no pages
        const imageRow = contentSheet.getRow(currentRow);
        const imageLabelCell = imageRow.getCell(2);
        imageLabelCell.value = "도면 없음";
        imageLabelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
        imageLabelCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        imageLabelCell.alignment = { horizontal: 'center', vertical: 'middle' };
        imageLabelCell.font = { bold: true, size: 9 };

        contentSheet.mergeCells(currentRow, 3, currentRow, 13);
        for (let i = 3; i <= 13; i++) {
          imageRow.getCell(i).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        }
        imageRow.height = 340;
      }

      // Generate buffer and download
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = targetExcel ? targetExcel.name : "특허_정리_결과.xlsx";
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Excel export error:", err);
      setError("엑셀 파일 저장 중 오류가 발생했습니다. 파일이 열려있는지 확인해주세요.");
    } finally {
      setExporting(false);
    }
  };

  const extractInfo = async () => {
    if (!file) return;
    setLoading(true);
    setLoadingImages(true);
    setLoadingDetails(false);
    setError(null);
    setResult(null);
    setExtractedImages([]);

    try {
      const reader = new FileReader();
      const base64Promise = new Promise<string>((resolve) => {
        reader.onload = () => {
          const base64 = (reader.result as string).split(",")[1];
          resolve(base64);
        };
        reader.readAsDataURL(file);
      });

      const base64Data = await base64Promise;
      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

      // Stage 1: Fast Drawing Detection using Gemini Flash
      const drawingDetectionResponse = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [
          {
            parts: [
              {
                inlineData: {
                  mimeType: "application/pdf",
                  data: base64Data,
                },
              },
              {
                text: "이 특허 문서에서 2D 도면, 기술 도식, 다이어그램, 그림 등이 포함된 모든 페이지 번호를 찾아주세요. 텍스트만 있는 페이지는 제외하세요. 결과는 오직 JSON 배열 형식으로만 응답하세요. 예: [5, 6, 7]",
              },
            ],
          },
        ],
        config: {
          responseMimeType: "application/json",
        }
      });

      let drawingPages: number[] = [];
      try {
        drawingPages = JSON.parse(drawingDetectionResponse.text || "[]");
      } catch (e) {
        console.error("Failed to parse drawing pages", e);
      }

      if (drawingPages.length > 0) {
        const images = await getPdfPages(file, drawingPages);
        setExtractedImages(images);
      }
      setLoadingImages(false);

      // Stage 2: Full Patent Extraction using Gemini Flash
      setLoadingDetails(true);
      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [
          {
            parts: [
              {
                inlineData: {
                  mimeType: "application/pdf",
                  data: base64Data,
                },
              },
              {
                text: `첨부된 특허 문서를 분석하여 다음 정보를 한국어로 추출해주세요. 
                결과는 반드시 JSON 형식으로 응답해야 합니다.
                
                추출 항목:
                1. ID (id)
                2. 발명의 명칭 (title)
                3. 구분 (category)
                4. 권리 (rightType)
                5. 출원 일자 (filingDate)
                6. 출원 번호 (applicationNumber)
                7. 등록 일자 (registrationDate)
                8. 등록 번호 (registrationNumber)
                9. 현재 상태 (status)
                10. 발명자 성명 (inventor)
                11. 출원인 명칭 (applicant)
                12. 비고 (note)
                13. 문서 요약 (summary)
                14. 도면이 포함된 모든 페이지 번호들 (drawingPages - 숫자 배열, 1부터 시작).`,
              },
            ],
          },
        ],
        config: {
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              id: { type: Type.STRING },
              title: { type: Type.STRING },
              category: { type: Type.STRING },
              rightType: { type: Type.STRING },
              filingDate: { type: Type.STRING },
              applicationNumber: { type: Type.STRING },
              registrationDate: { type: Type.STRING },
              registrationNumber: { type: Type.STRING },
              status: { type: Type.STRING },
              inventor: { type: Type.STRING },
              applicant: { type: Type.STRING },
              note: { type: Type.STRING },
              summary: { type: Type.STRING },
              drawingPages: { type: Type.ARRAY, items: { type: Type.INTEGER } },
            },
            required: ["id", "title", "filingDate", "applicationNumber", "inventor", "applicant", "summary", "drawingPages"],
          },
        },
      });

      const data = JSON.parse(response.text || "{}") as PatentInfo;
      setResult(data);

      // If Pro found more/different drawing pages, update them
      if (data.drawingPages && JSON.stringify(data.drawingPages) !== JSON.stringify(drawingPages)) {
        const moreImages = await getPdfPages(file, data.drawingPages);
        setExtractedImages(moreImages);
      }
    } catch (err) {
      console.error("Extraction error:", err);
      setError("정보를 추출하는 중 오류가 발생했습니다. 다시 시도해주세요.");
    } finally {
      setLoading(false);
      setLoadingDetails(false);
      setLoadingImages(false);
    }
  };

  const reset = () => {
    setFile(null);
    setResult(null);
    setExtractedImages([]);
    setLoadingImages(false);
    setLoadingDetails(false);
    setError(null);
    setPreviewUrl(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  return (
    <div className="min-h-screen bg-[#F5F5F0] text-[#141414] font-sans p-6 md:p-12">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <header className="mb-12 border-b border-[#141414]/10 pb-8">
          <h1 className="text-4xl md:text-5xl font-serif italic tracking-tight mb-2">
            Patent Info Extractor
          </h1>
          <p className="text-[#141414]/60 uppercase text-[11px] tracking-widest font-medium">
            AI-Powered Document Analysis Tool
          </p>
        </header>

        <main className="grid grid-cols-1 lg:grid-cols-2 gap-12">
          {/* Left Side: Upload & Preview */}
          <section className="space-y-8">
            <div className="bg-white rounded-3xl p-8 shadow-sm border border-[#141414]/5">
              <h2 className="text-xs uppercase tracking-widest font-bold mb-6 opacity-50">01. Upload Document</h2>
              
              {!file ? (
                <div 
                  onClick={() => fileInputRef.current?.click()}
                  className="border-2 border-dashed border-[#141414]/10 rounded-2xl p-12 flex flex-col items-center justify-center cursor-pointer hover:bg-[#141414]/5 transition-colors group"
                >
                  <div className="w-16 h-16 bg-[#141414]/5 rounded-full flex items-center justify-center mb-4 group-hover:scale-110 transition-transform">
                    <Upload className="w-8 h-8 opacity-40" />
                  </div>
                  <p className="text-sm font-medium">PDF 파일을 드래그하거나 클릭하여 선택하세요</p>
                  <p className="text-xs opacity-40 mt-2">특허 문서 (PDF) 전용</p>
                  <input 
                    type="file" 
                    ref={fileInputRef} 
                    onChange={handleFileChange} 
                    accept="application/pdf" 
                    className="hidden" 
                  />
                </div>
              ) : (
                <div className="space-y-4">
                  <div className="flex items-center justify-between bg-[#141414]/5 p-4 rounded-xl">
                    <div className="flex items-center gap-3">
                      <FileText className="w-6 h-6 opacity-60" />
                      <div>
                        <p className="text-sm font-medium truncate max-w-[200px]">{file.name}</p>
                        <p className="text-[10px] opacity-40">{(file.size / 1024 / 1024).toFixed(2)} MB</p>
                      </div>
                    </div>
                    <button 
                      onClick={reset}
                      className="text-xs font-bold uppercase tracking-tighter hover:underline"
                    >
                      Change
                    </button>
                  </div>

                  {!result && !loading && (
                    <button 
                      onClick={extractInfo}
                      className="w-full bg-[#141414] text-white py-4 rounded-xl font-bold uppercase tracking-widest text-xs hover:opacity-90 transition-opacity flex items-center justify-center gap-2"
                    >
                      Extract Information
                    </button>
                  )}
                </div>
              )}

              {error && (
                <div className="mt-4 p-4 bg-red-50 text-red-600 rounded-xl flex items-center gap-3 text-sm">
                  <AlertCircle className="w-5 h-5 flex-shrink-0" />
                  {error}
                </div>
              )}
            </div>

            {/* Excel Target Selection */}
            <div className="bg-white rounded-3xl p-8 shadow-sm border border-[#141414]/5">
              <h2 className="text-xs uppercase tracking-widest font-bold mb-6 opacity-50">02. Designate Excel File (Optional)</h2>
              <div 
                onClick={() => excelInputRef.current?.click()}
                className={`border-2 border-dashed ${targetExcel ? 'border-emerald-500/20 bg-emerald-50/30' : 'border-[#141414]/10'} rounded-2xl p-6 flex items-center gap-4 cursor-pointer hover:bg-[#141414]/5 transition-colors`}
              >
                <div className={`w-12 h-12 ${targetExcel ? 'bg-emerald-500/10 text-emerald-600' : 'bg-[#141414]/5 text-[#141414]/40'} rounded-full flex items-center justify-center`}>
                  <Table className="w-6 h-6" />
                </div>
                <div className="flex-grow">
                  <p className="text-sm font-medium">{targetExcel ? targetExcel.name : "정리할 엑셀 파일을 선택하세요"}</p>
                  <p className="text-[10px] opacity-40">미선택 시 새 파일이 생성됩니다</p>
                </div>
                {targetExcel && (
                  <button 
                    onClick={(e) => { e.stopPropagation(); setTargetExcel(null); }}
                    className="text-[10px] font-bold uppercase tracking-tighter opacity-40 hover:opacity-100"
                  >
                    Clear
                  </button>
                )}
                <input 
                  type="file" 
                  ref={excelInputRef} 
                  onChange={handleExcelChange} 
                  accept=".xlsx, .xls" 
                  className="hidden" 
                />
              </div>
            </div>

            {previewUrl && (
              <div className="bg-white rounded-3xl p-8 shadow-sm border border-[#141414]/5 h-[500px] flex flex-col">
                <h2 className="text-xs uppercase tracking-widest font-bold mb-6 opacity-50">03. Document Preview</h2>
                <iframe 
                  src={previewUrl} 
                  className="w-full flex-grow rounded-xl border border-[#141414]/10"
                  title="PDF Preview"
                />
              </div>
            )}
          </section>

          {/* Right Side: Results */}
          <section className="space-y-8">
            <div className="bg-white rounded-3xl p-8 shadow-sm border border-[#141414]/5 min-h-[400px] flex flex-col">
              <h2 className="text-xs uppercase tracking-widest font-bold mb-6 opacity-50">04. Extraction Results</h2>
              
              {loading && !extractedImages.length ? (
                <div className="flex-grow flex flex-col items-center justify-center space-y-4">
                  <Loader2 className="w-12 h-12 animate-spin opacity-20" />
                  <div className="text-center">
                    <p className="text-sm font-medium animate-pulse">
                      {loadingImages ? "도면을 먼저 찾고 있습니다..." : "상세 정보를 분석하고 있습니다..."}
                    </p>
                    <p className="text-[10px] uppercase tracking-widest font-bold opacity-30 mt-2">
                      {loadingImages ? "Stage 1: Drawing Detection" : "Stage 2: Patent Analysis"}
                    </p>
                  </div>
                </div>
              ) : result || extractedImages.length > 0 || loadingDetails ? (
                <div className="flex-grow flex flex-col">
                  <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-700">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2 text-emerald-600">
                        {result ? (
                          <>
                            <CheckCircle2 className="w-5 h-5" />
                            <span className="text-xs font-bold uppercase tracking-widest">Analysis Complete</span>
                          </>
                        ) : (
                          <>
                            <Loader2 className="w-5 h-5 animate-spin text-blue-500" />
                            <span className="text-xs font-bold uppercase tracking-widest text-blue-500">Stage 2: Extracting Details...</span>
                          </>
                        )}
                      </div>
                      {result && (
                        <button 
                          onClick={exportToExcel}
                          disabled={exporting}
                          className="bg-emerald-600 text-white px-4 py-2 rounded-lg text-[10px] font-bold uppercase tracking-widest flex items-center gap-2 hover:bg-emerald-700 transition-colors disabled:opacity-50"
                        >
                          {exporting ? <Loader2 className="w-3 h-3 animate-spin" /> : <Download className="w-3 h-3" />}
                          Export to Excel
                        </button>
                      )}
                    </div>

                    {/* Drawing Preview Section - Always show if we have images */}
                    {(extractedImages.length > 0 || loadingImages) && (
                      <div className="p-6 bg-[#141414]/5 rounded-3xl border border-[#141414]/5">
                        <div className="flex items-center justify-between mb-6">
                          <h2 className="text-[10px] uppercase tracking-widest font-bold opacity-40">Extracted Drawings</h2>
                          {loadingImages && (
                            <div className="flex items-center gap-2 text-[10px] uppercase tracking-widest font-bold text-blue-500 animate-pulse">
                              <RefreshCw className="w-3 h-3 animate-spin" />
                              Scanning PDF...
                            </div>
                          )}
                        </div>
                        
                        {extractedImages.length > 0 ? (
                          <div className="grid grid-cols-1 gap-6">
                            {extractedImages.map((img, idx) => (
                              <div key={idx} className="bg-white rounded-2xl p-4 flex flex-col items-center border border-[#141414]/5 shadow-sm">
                                <div className="w-full bg-white rounded-xl overflow-hidden mb-3 border border-[#141414]/5">
                                  <img 
                                    src={img.data} 
                                    alt={`Drawing ${idx + 1}`} 
                                    className="w-full h-auto object-contain max-h-[600px]"
                                    referrerPolicy="no-referrer"
                                  />
                                </div>
                                <p className="text-[10px] font-bold uppercase tracking-tighter opacity-40">
                                  Drawing {idx + 1} (Page {img.pageNum})
                                </p>
                              </div>
                            ))}
                          </div>
                        ) : (
                          <div className="h-40 flex items-center justify-center opacity-20">
                            <p className="text-sm font-medium">도면을 찾고 있습니다...</p>
                          </div>
                        )}
                      </div>
                    )}

                    {result && (
                      <div className="space-y-6">
                        <div className="border-b border-[#141414]/5 pb-4">
                          <label className="text-[10px] uppercase tracking-widest font-bold opacity-40 block mb-1">발명의 명칭</label>
                          <p className="text-xl font-serif italic">{result.title}</p>
                        </div>

                        <div className="grid grid-cols-2 gap-6">
                          <div className="border-b border-[#141414]/5 pb-4">
                            <label className="text-[10px] uppercase tracking-widest font-bold opacity-40 block mb-1">출원 일자</label>
                            <p className="text-sm font-mono">{result.filingDate}</p>
                          </div>
                          <div className="border-b border-[#141414]/5 pb-4">
                            <label className="text-[10px] uppercase tracking-widest font-bold opacity-40 block mb-1">출원 번호</label>
                            <p className="text-sm font-mono">{result.applicationNumber}</p>
                          </div>
                        </div>

                        <div className="grid grid-cols-2 gap-6">
                          <div className="border-b border-[#141414]/5 pb-4">
                            <label className="text-[10px] uppercase tracking-widest font-bold opacity-40 block mb-1">발명자 성명</label>
                            <p className="text-sm font-medium">{result.inventor}</p>
                          </div>
                          <div className="border-b border-[#141414]/5 pb-4">
                            <label className="text-[10px] uppercase tracking-widest font-bold opacity-40 block mb-1">출원인 명칭</label>
                            <p className="text-sm font-medium">{result.applicant}</p>
                          </div>
                        </div>

                        {result.summary && (
                          <div className="bg-[#141414]/5 p-6 rounded-2xl">
                            <label className="text-[10px] uppercase tracking-widest font-bold opacity-40 block mb-2">핵심 요약</label>
                            <div className="text-sm leading-relaxed opacity-80">
                              <Markdown>{result.summary}</Markdown>
                            </div>
                          </div>
                        )}
                      </div>
                    )}
                  </div>

                  <button 
                    onClick={reset}
                    className="flex items-center gap-2 text-xs font-bold uppercase tracking-widest opacity-40 hover:opacity-100 transition-opacity mt-8 pt-4 border-t border-[#141414]/5 w-full"
                  >
                    <RefreshCw className="w-4 h-4" />
                    Analyze Another Document
                  </button>
                </div>
              ) : (
                <div className="flex-grow flex flex-col items-center justify-center opacity-20">
                  <FileText className="w-20 h-20 mb-4" />
                  <p className="text-sm font-medium">문서를 업로드하면 결과가 여기에 표시됩니다</p>
                </div>
              )}
            </div>
          </section>
        </main>

        <footer className="mt-24 border-t border-[#141414]/10 pt-8 text-center">
          <p className="text-[10px] uppercase tracking-widest font-bold opacity-30">
            Powered by Gemini 3 Flash & Google AI Studio
          </p>
        </footer>
      </div>
    </div>
  );
}

