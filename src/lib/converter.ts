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

        if (questionStartRegex.test(nextText)) {
          break; // Next question found
        }

        if (nextEl.tagName === 'P') {
            if (optionRegex.test(nextText)) {
              // Split options that are on the same line
              const sameLineOptions = nextText.split(/\s*(?=\([B-D]\))/i);
              for(const opt of sameLineOptions) {
                if(optionRegex.test(opt)) {
                  questionData.options.push(opt);
                   const optionImg = nextEl.querySelector('img');
                  if (optionImg?.src) {
                    // This is imperfect for multiple images in one P tag, but will do for now
                    questionData.images.push({ data: optionImg.src, in: `option${questionData.options.length}` });
                  }
                }
              }
            } else if (nextText) { // continuation of previous line
                if(questionData.options.length > 0) {
                    const lastOptionIndex = questionData.options.length - 1;
                    questionData.options[lastOptionIndex] += '\n' + nextText;
                } else {
                    questionData.questionText += '\n' + nextText;
                }
            }
        }
        
        const nextElImgs = nextEl.querySelectorAll('img');
        nextElImgs.forEach(img => {
            if (img.src && !questionData.images.some(existingImg => existingImg.data === img.src)) {
                 if (questionData.options.length > 0) {
                    questionData.images.push({ data: img.src, in: `option${questionData.options.length}` });
                } else {
                    questionData.images.push({ data: img.src, in: 'question' });
                }
            }
        });

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
    { header: 'Sr. No', key: 'sr', width: 8 },
    { header: 'Question content', key: 'question', width: 50 },
    { header: 'Image', key: 'image', width: 40 },
    { header: 'Alternative1', key: 'alt1', width: 30 },
    { header: 'Alternative2', key: 'alt2', width: 30 },
    { header: 'Alternative3', key: 'alt3', width: 30 },
    { header: 'Alternative4', key: 'alt4', width: 30 },
  ];

  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4F81BD' },
  };

  let currentRowNum = 2;
  for (const [index, q] of questions.entries()) {
    
    const rowData: any = {
      sr: index + 1,
      question: q.questionText,
    };

    const cleanOption = (text: string) => text.replace(/^\s*\([A-D]\)\s*/i, '').trim();

    q.options.forEach((opt, i) => {
        rowData[`alt${i+1}`] = cleanOption(opt);
    });

    const row = worksheet.addRow(rowData);
    row.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
    
    let maxLines = q.questionText.split('\n').length;
    q.options.forEach(opt => {
        maxLines = Math.max(maxLines, opt.split('\n').length);
    });

    let rowHeight = maxLines * 15 + 10;
    
    const questionImages = q.images.filter(img => img.in === 'question');
    let imageAdded = false;

    if (questionImages.length > 0) {
      const img = questionImages[0];
      if (img.data) {
        const base64string = img.data;
        try {
            const extension = base64string.startsWith('data:image/jpeg') ? 'jpeg' : 'png';
            const base64Data = base64string.substring(base64string.indexOf(',') + 1);
            const imageId = workbook.addImage({ base64: base64Data, extension });
            
            worksheet.addImage(imageId, {
              tl: { col: 2, row: currentRowNum - 1 }, // Column C for Image
              ext: { width: 300, height: 225 }
            });
            imageAdded = true;
            rowHeight = Math.max(rowHeight, 235); // 225 for image + 10 padding
        } catch (e) {
            console.error("Could not add image", e);
        }
      }
    }
    
    // Add images for options in their respective columns if any
    for(let i=0; i<4; i++){
      const optionImages = q.images.filter(img => img.in === `option${i+1}`);
      if(optionImages.length > 0){
        const img = optionImages[0];
        if (img.data) {
           const base64string = img.data;
            try {
              const extension = base64string.startsWith('data:image/jpeg') ? 'jpeg' : 'png';
              const base64Data = base64string.substring(base64string.indexOf(',') + 1);
              const imageId = workbook.addImage({ base64: base64Data, extension });

              worksheet.addImage(imageId, {
                tl: { col: 3 + i, row: currentRowNum - 1 }, // Columns D, E, F, G
                ext: { width: 150, height: 112.5 }
              });
              rowHeight = Math.max(rowHeight, 122.5); // 112.5 for image + 10 padding
            } catch (e) {
                console.error(`Could not add image for option ${i+1}`, e);
            }
        }
      }
    }
    
    row.height = rowHeight;
    currentRowNum = worksheet.rowCount + 1;
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
    