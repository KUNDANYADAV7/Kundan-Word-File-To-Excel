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
          break; 
        }

        if (nextEl.tagName === 'P') {
            if (optionRegex.test(nextText)) {
              const sameLineOptions = nextText.split(/\s*(?=\([B-D]\))/i);
              for(const opt of sameLineOptions) {
                if(optionRegex.test(opt)) {
                  questionData.options.push(opt);
                }
              }
            } else if (nextText) { 
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
                    const lastOption = questionData.options[questionData.options.length - 1];
                    const optionLabelMatch = lastOption.match(/^\s*\(([A-D])\)/i);
                    if(optionLabelMatch){
                        const optionLetter = optionLabelMatch[1].toUpperCase();
                        questionData.images.push({ data: img.src, in: `option${optionLetter}` });
                    }
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

const PIXELS_TO_EMUS = 9525;
const DEFAULT_ROW_HEIGHT = 15; // Corresponds to font size 12

const getBase64Image = (imgSrc: string): { extension: 'png' | 'jpeg', data: string } => {
    const extension = imgSrc.startsWith('data:image/jpeg') ? 'jpeg' : 'png';
    const data = imgSrc.substring(imgSrc.indexOf(',') + 1);
    return { extension, data };
}

const getImageDimensions = (imgSrc: string): Promise<{ width: number; height: number }> => {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve({ width: img.width, height: img.height });
        img.onerror = reject;
        img.src = imgSrc;
    });
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
    { header: 'Question content', key: 'question', width: 60 },
    { header: 'Alternative1', key: 'alt1', width: 35 },
    { header: 'Alternative2', key: 'alt2', width: 35 },
    { header: 'Alternative3', key: 'alt3', width: 35 },
    { header: 'Alternative4', key: 'alt4', width: 35 },
  ];

  const headerRow = worksheet.getRow(1);
  headerRow.font = { name: 'Calibri', bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4F81BD' },
  };
  headerRow.height = 20;

  for (const [index, q] of questions.entries()) {
    const cleanOption = (text: string) => text.replace(/^\s*\([A-D]\)\s*/i, '').trim();

    const optionsMap: {[key: string]: string} = {};
    q.options.forEach(opt => {
        const match = opt.match(/^\s*\(([A-D])\)/i);
        if(match){
            const letter = match[1].toUpperCase();
            optionsMap[letter] = cleanOption(opt);
        }
    });

    const row = worksheet.addRow({
      sr: index + 1,
      question: q.questionText,
      alt1: optionsMap['A'] || '',
      alt2: optionsMap['B'] || '',
      alt3: optionsMap['C'] || '',
      alt4: optionsMap['D'] || '',
    });
    
    // Default alignment and font for all cells in the row
    row.eachCell({ includeEmpty: true }, cell => {
        cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        cell.font = { name: 'Calibri', size: 11 };
    });

    let maxRowHeight = 0;
    
    // --- Calculate height for Question Cell ---
    let questionCellHeight = 0;
    const questionTextLines = q.questionText.split('\n').length;
    questionCellHeight += questionTextLines * DEFAULT_ROW_HEIGHT;

    const questionImages = q.images.filter(img => img.in === 'question');
    if (questionImages.length > 0) {
      const imgData = questionImages[0].data;
      if (imgData) {
        try {
            const { extension, data } = getBase64Image(imgData);
            const imageId = workbook.addImage({ base64: data, extension });
            const imageDims = await getImageDimensions(imgData);
            
            const imageWidth = 300; // Fixed width for question image
            const imageHeight = (imageDims.height / imageDims.width) * imageWidth;

            // Add a small margin before the image
            const imageTopMargin = 5; 
            const textHeight = questionTextLines * DEFAULT_ROW_HEIGHT;

            worksheet.addImage(imageId, {
              tl: { col: 1, row: row.number - 1, rowOff: (textHeight + imageTopMargin) * PIXELS_TO_EMUS, colOff: 5 * PIXELS_TO_EMUS },
              ext: { width: imageWidth, height: imageHeight }
            });
            questionCellHeight += imageHeight + imageTopMargin + 5; // add bottom margin
        } catch (e) { console.error("Could not add question image", e); }
      }
    }
    maxRowHeight = Math.max(maxRowHeight, questionCellHeight);

    // --- Calculate height for Option Cells ---
    let maxOptionHeight = 0;
    const allOptionTexts = [optionsMap['A']||'', optionsMap['B']||'', optionsMap['C']||'', optionsMap['D']||''];
    const maxOptionTextLines = Math.max(...allOptionTexts.map(t => t.split('\n').length));
    maxOptionHeight += maxOptionTextLines * DEFAULT_ROW_HEIGHT;

    let maxOptionImageHeight = 0;
    for (const [i, letter] of ['A', 'B', 'C', 'D'].entries()) {
        const optionImages = q.images.filter(img => img.in === `option${letter}`);
        if(optionImages.length > 0) {
            const imgData = optionImages[0].data;
            if (imgData) {
                try {
                    const { extension, data } = getBase64Image(imgData);
                    const imageId = workbook.addImage({ base64: data, extension });
                    const imageDims = await getImageDimensions(imgData);

                    const imageWidth = 180; // Fixed width for option images
                    const imageHeight = (imageDims.height / imageDims.width) * imageWidth;
                    
                    const textHeight = (optionsMap[letter] || '').split('\n').length * DEFAULT_ROW_HEIGHT;
                    const imageTopMargin = 5;

                    worksheet.addImage(imageId, {
                        tl: { col: 2 + i, row: row.number - 1, rowOff: (textHeight + imageTopMargin) * PIXELS_TO_EMUS, colOff: 5 * PIXELS_TO_EMUS },
                        ext: { width: imageWidth, height: imageHeight }
                    });
                    maxOptionImageHeight = Math.max(maxOptionImageHeight, textHeight + imageHeight + imageTopMargin + 5);
                } catch (e) { console.error(`Could not add image for option ${letter}`, e); }
            }
        }
    }
    maxOptionHeight = Math.max(maxOptionHeight, maxOptionImageHeight);
    maxRowHeight = Math.max(maxRowHeight, maxOptionHeight);

    row.height = maxRowHeight > 0 ? maxRowHeight : DEFAULT_ROW_HEIGHT * Math.max(questionTextLines, maxOptionTextLines, 1);
  }
  
  // Apply borders to all cells
  worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    row.eachCell({ includeEmpty: true }, function(cell, colNumber) {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      };
    });
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  saveAs(blob, `${file.name.replace(/\.docx$/, '')}.xlsx`);
};
