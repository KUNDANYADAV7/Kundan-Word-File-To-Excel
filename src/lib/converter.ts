"use client";

import mammoth from 'mammoth';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

type Question = {
  questionText: string;
  options: string[];
  images: { data: string; in: 'question' | string }[];
};

const parseHtmlToQuestions = (html: string): Question[] => {
  const questions: Question[] = [];
  if (typeof window === 'undefined') return questions;

  const container = document.createElement('div');
  container.innerHTML = html;

  const children = Array.from(container.children);
  let i = 0;
  while (i < children.length) {
    const el = children[i] as HTMLElement;
    const text = el.innerText.trim();
    const questionStartRegex = /^(?:Q|Question)?\s*\d+[.)]\s*/;
    
    if (el.tagName === 'P' && questionStartRegex.test(text)) {
      const questionData: Question = {
        questionText: text.replace(questionStartRegex, ''),
        options: [],
        images: [],
      };

      const questionImg = el.querySelector('img');
      if (questionImg?.src) {
        questionData.images.push({ data: questionImg.src, in: 'question' });
      }

      let j = i + 1;
      while (j < children.length) {
        const nextEl = children[j] as HTMLElement;
        const nextText = nextEl.innerText.trim();
        const optionRegex = /^\s*\([A-D]\)\s*/i;

        if (nextEl.tagName !== 'P' && !nextEl.querySelector('img')) { 
            j++;
            continue;
        }

        if (questionStartRegex.test(nextText)) {
          break;
        } else if (optionRegex.test(nextText)) {
          questionData.options.push(nextText);
          const optionImg = nextEl.querySelector('img');
          if (optionImg?.src) {
            questionData.images.push({ data: optionImg.src, in: `option${questionData.options.length}` });
          }
        } else if (nextText || nextEl.querySelector('img')) {
          if (questionData.options.length === 0) {
            if (nextText) questionData.questionText += '\n' + nextText;
            const img = nextEl.querySelector('img');
            if (img?.src) {
              questionData.images.push({ data: img.src, in: 'question' });
            }
          }
        }
        j++;
      }
      
      if (questionData.questionText && questionData.options.length > 0) {
        questions.push(questionData);
      }
      
      i = j;
    } else {
      i++;
    }
  }

  return questions;
};

export const convertDocxToExcel = async (file: File) => {
  const arrayBuffer = await file.arrayBuffer();

  const { value: html } = await mammoth.convertToHtml({ arrayBuffer });

  const questions = parseHtmlToQuestions(html);

  if (questions.length === 0) {
    throw new Error("No questions found. Check document format. Questions should be numbered (e.g., '1.') and options labeled (e.g., '(A)').");
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Questions');

  worksheet.columns = [
    { header: 'Sr. No', key: 'sr', width: 10 },
    { header: 'Question content', key: 'question', width: 70 },
    { header: 'Alternative1', key: 'opt1', width: 30 },
    { header: 'Alternative2', key: 'opt2', width: 30 },
    { header: 'Alternative3', key: 'opt3', width: 30 },
    { header: 'Alternative4', key: 'opt4', width: 30 },
  ];

  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4F81BD' }, // A nice blue
  };

  let currentRowNum = 2;
  for (const [index, q] of questions.entries()) {
    const questionTextWithImages = q.questionText + (q.images.length > 0 ? "\n(See image below)" : "");

    const row = worksheet.addRow({
      sr: index + 1,
      question: questionTextWithImages,
      opt1: q.options[0] || '',
      opt2: q.options[1] || '',
      opt3: q.options[2] || '',
      opt4: q.options[3] || '',
    });

    row.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    
    let textLines = questionTextWithImages.split('\n').length;
    let rowHeight = textLines * 15 + 10;
    
    const questionImages = q.images.filter(img => img.in === 'question');
    if (questionImages.length > 0) {
      const img = questionImages[0];
      if (img.data) {
        const base64string = img.data;
        const extension = base64string.startsWith('data:image/jpeg') ? 'jpeg' : 'png';
        const base64Data = base64string.substring(base64string.indexOf(',') + 1);

        const imageId = workbook.addImage({ base64: base64Data, extension });
        
        // Place image in the question column
        worksheet.addImage(imageId, {
          tl: { col: 1, row: currentRowNum - 1 },
          ext: { width: 420, height: 200 }
        });
        rowHeight = Math.max(rowHeight, 210); // Set row height to accommodate image
      }
    }
    
    row.height = rowHeight;
    currentRowNum++;
  }
  
  const totalRows = worksheet.rowCount;
  for (let i = 1; i <= totalRows; i++) {
    const row = worksheet.getRow(i);
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  saveAs(blob, `${file.name.replace(/\.docx$/, '')}.xlsx`);
};
