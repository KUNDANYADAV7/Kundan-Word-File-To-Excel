"use client";

import mammoth from 'mammoth';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import * as pdfjsLib from 'pdfjs-dist';

// Set workerSrc to a CDN URL to avoid build issues with Next.js
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;


type Question = {
  questionText: string;
  options: { [key: string]: string };
  images: { data: string; in: 'question' | string }[];
};

const PIXELS_TO_EMUS = 9525;
const DEFAULT_ROW_HEIGHT_IN_POINTS = 21.75; 
const POINTS_TO_PIXELS = 4 / 3;
const IMAGE_MARGIN_PIXELS = 15;


const parseHtmlToQuestions = (html: string): Question[] => {
  const questions: Question[] = [];
  if (typeof window === 'undefined') return questions;

  const container = document.createElement('div');
  container.innerHTML = html;

  const processTextContent = (element: HTMLElement): string => {
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = element.innerHTML;
    tempDiv.querySelectorAll('sup').forEach(sup => {
      if (sup.textContent === '2') sup.textContent = '²';
    });
    tempDiv.querySelectorAll('img').forEach(img => img.remove());
    let text = tempDiv.textContent?.replace(/\s+/g, ' ').trim() || '';
    text = text.replace(/ deg/g, '°');
    return text;
  };
  
  const children = Array.from(container.children);
  let i = 0;
  while (i < children.length) {
    const el = children[i] as HTMLElement;
    const text = processTextContent(el);
    const questionStartRegex = /^(?:Q|Question)?\s*\d+[.)]\s*/;
    
    if (el.tagName === 'P' && questionStartRegex.test(text)) {
      const currentQuestion: Question = {
        questionText: text.replace(questionStartRegex, ''),
        options: {},
        images: [],
      };

      // Process question images that are in the same <p> as the question text
      el.querySelectorAll('img').forEach(img => {
        if(img.src && !currentQuestion.images.some(existing => existing.data === img.src)) {
            currentQuestion.images.push({ data: img.src, in: 'question' });
        }
      });

      let j = i + 1;
      let lastOptionLetter: string | null = null;
      let isOptionLine = false;

      while (j < children.length) {
        const nextEl = children[j] as HTMLElement;
        const nextText = processTextContent(nextEl);
        
        if (nextEl.tagName === 'P' && questionStartRegex.test(nextText)) {
          break; 
        }
        
        const optionRegex = /\s*\(([A-D])\)\s*/i;

        // Process images first, and associate them correctly
        const imagesInElement = Array.from(nextEl.querySelectorAll('img'));
        if (imagesInElement.length > 0) {
            const parentText = nextEl.textContent || '';
            const optionMatchInParent = parentText.match(optionRegex);
            
            imagesInElement.forEach(img => {
                if (img.src && !currentQuestion.images.some(existingImg => existingImg.data === img.src)) {
                    if (optionMatchInParent && optionMatchInParent[1]) {
                        const letter = optionMatchInParent[1].toUpperCase();
                        currentQuestion.images.push({ data: img.src, in: `option${letter}` });
                    } else if(lastOptionLetter) {
                        currentQuestion.images.push({ data: img.src, in: `option${lastOptionLetter}` });
                    } else {
                        currentQuestion.images.push({ data: img.src, in: 'question' });
                    }
                }
            });
        }

        // Process text content
        if (nextEl.tagName === 'P') {
          const parts = nextText.split(/\s*(?=\([B-D]\))/i);
          let containsOption = false;
          parts.forEach(part => {
              const optionMatch = part.match(optionRegex);
              if (optionMatch && optionMatch[1]) {
                  containsOption = true;
                  const letter = optionMatch[1].toUpperCase();
                  const optionText = part.replace(optionRegex, '').trim();
                  currentQuestion.options[letter] = (currentQuestion.options[letter] || '') + optionText;
                  lastOptionLetter = letter;
              }
          });
          
          if(!containsOption && nextText) { // If it's not an option line but has text
            if(lastOptionLetter) {
                 // Append to the last identified option if it's not a new option line
                currentQuestion.options[lastOptionLetter] += '\n' + nextText;
            } else {
                // Append to question text if it appears before any options
                currentQuestion.questionText += '\n' + nextText;
            }
          }
        }
        
        j++;
      }
      
      if (currentQuestion.questionText || Object.keys(currentQuestion.options).length > 0 || currentQuestion.images.length > 0) {
        questions.push(currentQuestion);
      }
      
      i = j;
    } else {
      i++;
    }
  }

  return questions;
};


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

const formatTextForExcel = (text: string): string => {
    return text;
};

const generateExcelFromQuestions = async (questions: Question[], fileName: string) => {
  if (questions.length === 0) {
    throw new Error("No questions found. Check document format. Questions should be numbered (e.g., '1.') and options labeled (e.g., '(A)').");
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Questions');

  worksheet.columns = [
    { header: 'Sr. No', key: 'sr', width: 5.43 },
    { header: 'Question content', key: 'question', width: 110.57 },
    { header: 'Alternative1', key: 'alt1', width: 35.71 },
    { header: 'Alternative2', key: 'alt2', width: 35.71 },
    { header: 'Alternative3', key: 'alt3', width: 35.71 },
    { header: 'Alternative4', key: 'alt4', width: 35.71 },
  ];

  const headerRow = worksheet.getRow(1);
  headerRow.font = { name: 'Calibri', bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4F81BD' },
  };
  headerRow.height = 43.5;

  for (const [index, q] of questions.entries()) {
    const row = worksheet.addRow({
      sr: index + 1,
      question: formatTextForExcel(q.questionText),
      alt1: formatTextForExcel(q.options['A'] || ''),
      alt2: formatTextForExcel(q.options['B'] || ''),
      alt3: formatTextForExcel(q.options['C'] || ''),
      alt4: formatTextForExcel(q.options['D'] || ''),
    });
    
    row.eachCell({ includeEmpty: true }, cell => {
        cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true };
        cell.font = { name: 'Calibri', size: 11 };
    });
    row.height = DEFAULT_ROW_HEIGHT_IN_POINTS;

    let maxRowHeightInPoints = 0;
    
    const calculateCellHeightAndPlaceImages = async (cell: ExcelJS.Cell, text: string, images: {data: string, in: string}[]) => {
        let textHeightInPixels = 0;
        if (text) {
          const lines = text.split('\n');
          const canvas = document.createElement("canvas");
          const context = canvas.getContext("2d");
          if(!context) return { totalHeight: 0 };
          context.font = "11pt Calibri";
          const columnWidth = worksheet.getColumn(cell.col).width;
          // approximate character width, this is a weak point.
          const avgCharWidth = 7; 
          const availableWidth = columnWidth ? columnWidth * avgCharWidth : 100 * avgCharWidth;

          lines.forEach(line => {
             let currentLine = '';
             const words = line.split(' ');
             for(const word of words) {
                if(context.measureText(currentLine + ' ' + word).width > availableWidth) {
                  textHeightInPixels += 20; // Approximate line height
                  currentLine = word;
                } else {
                  currentLine += (currentLine ? ' ' : '') + word;
                }
             }
             textHeightInPixels += 20;
          });
        }
        
        let cumulativeImageHeight = 0;

        if (images.length > 0) {
           for (const imgData of images) {
              try {
                  const { extension, data } = getBase64Image(imgData.data);
                  const imageId = workbook.addImage({ base64: data, extension });
                  const imageDims = await getImageDimensions(imgData.data);
                  
                  const imageWidthInPixels = 120;
                  const imageHeightInPixels = (imageDims.height / imageDims.width) * imageWidthInPixels;
                  
                  const rowOffsetInPixels = textHeightInPixels + cumulativeImageHeight + IMAGE_MARGIN_PIXELS;
                  cumulativeImageHeight += imageHeightInPixels + IMAGE_MARGIN_PIXELS;

                  const column = worksheet.getColumn(cell.col);
                  const cellWidthInPixels = column.width ? column.width * 7.5 : 100; // 7.5 is a better approximation
                  const colOffsetInPixels = Math.max(0, (cellWidthInPixels - imageWidthInPixels) / 2);
                  
                  worksheet.addImage(imageId, {
                    tl: { col: cell.col - 1, row: cell.row - 1 },
                    ext: { width: imageWidthInPixels, height: imageHeightInPixels }
                  });
                  
                  if ((worksheet as any).media && (worksheet as any).media.length > 0) {
                    const lastImage = (worksheet as any).media[(worksheet as any).media.length - 1];
                    if (lastImage && lastImage.range) {
                      lastImage.range.tl.rowOff = rowOffsetInPixels * PIXELS_TO_EMUS;
                      lastImage.range.tl.colOff = colOffsetInPixels * PIXELS_TO_EMUS;
                    }
                  }
              } catch (e) { console.error("Could not add image", e); }
           }
        }
        
        const totalCellHeightInPixels = textHeightInPixels + cumulativeImageHeight;
        return { totalHeight: totalCellHeightInPixels / POINTS_TO_PIXELS };
    };
    
    const questionImages = q.images.filter(img => img.in === 'question');
    const { totalHeight: questionCellHeight } = await calculateCellHeightAndPlaceImages(row.getCell('question'), q.questionText, questionImages);
    maxRowHeightInPoints = Math.max(maxRowHeightInPoints, questionCellHeight);

    let maxOptionHeight = 0;
    for (const [i, letter] of ['A', 'B', 'C', 'D'].entries()) {
        const optionText = q.options[letter] || '';
        const optionImages = q.images.filter(img => img.in === `option${letter}`);
        const cell = row.getCell(`alt${i+1}`);
        const { totalHeight: optionCellHeight } = await calculateCellHeightAndPlaceImages(cell, optionText, optionImages);
        maxOptionHeight = Math.max(maxOptionHeight, optionCellHeight);
    }
    maxRowHeightInPoints = Math.max(maxRowHeightInPoints, maxOptionHeight);

    row.height = maxRowHeightInPoints > DEFAULT_ROW_HEIGHT_IN_POINTS ? maxRowHeightInPoints : DEFAULT_ROW_HEIGHT_IN_POINTS;
  }
  
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
  saveAs(blob, `${fileName.replace(/\.(docx|pdf)$/, '')}.xlsx`);
};


export const convertDocxToExcel = async (file: File) => {
  const arrayBuffer = await file.arrayBuffer();

  const { value: rawHtml } = await mammoth.convertToHtml({ arrayBuffer }, {
    transformDocument: mammoth.transforms.paragraph(p => {
        p.children.forEach(run => {
            if (run.type === 'run') {
                if (run.isSuperscript) {
                     run.children.forEach(text => {
                        if (text.type === 'text' && text.value === '2') {
                           text.value = '²';
                        }
                    });
                }
                run.children.forEach(text => {
                    if (text.type === 'text') {
                        text.value = text.value.replace(/°/g, ' deg');
                    }
                });
            }
        });
        return p;
    })
  });
  
  const questions = parseHtmlToQuestions(rawHtml);
  await generateExcelFromQuestions(questions, file.name);
};


export const convertPdfToExcel = async (file: File) => {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
    const numPages = pdf.numPages;
    
    let fullText = '';
    for (let i = 1; i <= numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        fullText += textContent.items.map(item => ('str' in item ? item.str : '')).join(' ') + '\n';
    }

    const questions: Question[] = [];
    // A more robust regex to find questions, allowing for variations in spacing.
    const questionRegex = /^(?:Q|Question)?\s*(\d+)[.)\s]/i;
    // Split by what looks like a question marker, but keep the delimiter
    const potentialQuestions = fullText.split(new RegExp(`(${questionRegex.source})`, 'gi'));

    let questionBuffer: string[] = [];

    for (const part of potentialQuestions) {
        if (questionRegex.test(part)) {
            if (questionBuffer.length > 0) {
                // Process the previous question
                const questionBlock = questionBuffer.join('\n');
                const [questionText, ...optionsPart] = questionBlock.split(/\s*\([A-D]\)/i);

                if (questionText) {
                    const currentQuestion: Question = {
                        questionText: questionText.replace(questionRegex, '').trim(),
                        options: {},
                        images: [],
                    };
                    const optionLines = questionBlock.match(/\s*\([A-D]\).*?(?=\s*\([A-D]\)|$)/gis) || [];
                    optionLines.forEach(line => {
                        const optionMatch = line.match(/\s*\(([A-D])\)/i);
                        if(optionMatch && optionMatch[1]){
                            const letter = optionMatch[1].toUpperCase();
                            const optionText = line.replace(optionMatch[0], '').trim();
                            currentQuestion.options[letter] = optionText;
                        }
                    });

                    if(Object.keys(currentQuestion.options).length > 0) {
                        questions.push(currentQuestion);
                    }
                }
            }
            questionBuffer = [part];
        } else {
            questionBuffer.push(part);
        }
    }
    // Process the last question in the buffer
    if(questionBuffer.length > 0) {
        const questionBlock = questionBuffer.join('\n');
        const match = questionBlock.match(questionRegex);
        const textAfterMarker = match ? questionBlock.substring(match[0].length) : questionBlock;
        
        const optionLines = textAfterMarker.match(/\s*\([A-D]\).*?(?=\s*\([A-D]\)|$)/gis) || [];
        const firstOptionIndex = textAfterMarker.indexOf(optionLines[0] || '');

        if (match && optionLines.length > 0 && firstOptionIndex !== -1) {
            const questionText = textAfterMarker.substring(0, firstOptionIndex).trim();
            const currentQuestion: Question = {
                questionText: questionText,
                options: {},
                images: [],
            };
            
            optionLines.forEach(line => {
                const optionMatch = line.match(/\s*\(([A-D])\)/i);
                if(optionMatch && optionMatch[1]){
                    const letter = optionMatch[1].toUpperCase();
                    const optionText = line.replace(optionMatch[0], '').trim();
                    currentQuestion.options[letter] = optionText;
                }
            });

            if (Object.keys(currentQuestion.options).length > 0) {
                questions.push(currentQuestion);
            }
        }
    }

    await generateExcelFromQuestions(questions, file.name);
};
