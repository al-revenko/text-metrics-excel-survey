import fs from 'node:fs';
import { readDataFiles } from './helpers/readDataFiles';
import path from 'node:path';
import { BLANK_FOLDER } from './const/files';
import { Workbook } from 'exceljs';
import { LEGEND_COLUMNS, METRICS_CELLS } from './const/excel';

async function convertToBlanks() {
  const inputData = await readDataFiles('inputData')

  if (!fs.existsSync(BLANK_FOLDER)) fs.mkdirSync(BLANK_FOLDER);

  const blankMetrics = Array(10).fill('0.00')
  
  for (const { fileName, messagesData } of inputData) {
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet(fileName); 
    
    // Устанавливаем имена основных колонок
    worksheet.columns = LEGEND_COLUMNS;  
    
    // Устанавливаем заголовки
    worksheet.mergeCells('A1:M1');
    
    const titleRow = worksheet.getCell('A1');
    
    titleRow.value = fileName;
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    titleRow.font = { bold: true }; 
    
    worksheet.mergeCells('B2:C2');
    
    const sourceDataRow = worksheet.getCell('B2')
    
    sourceDataRow.value = 'ИСХОДНЫЕ ДАННЫЕ';
    sourceDataRow.alignment = { horizontal: 'center', vertical: 'middle' };
    sourceDataRow.font = { bold: true };  
    
    worksheet.mergeCells('D2:M2');
    
    const testDataRow = worksheet.getCell('D2')
    
    testDataRow.value = 'ОЦЕНОЧНЫЕ ДАННЫЕ';
    testDataRow.alignment = { horizontal: 'center', vertical: 'middle' };
    testDataRow.font = { bold: true };  
    
    // Добавляем заголовки второй строки
    const headerRow = worksheet.addRow([
      '№', 'Текст сообщения', 'id (ключ)', ...METRICS_CELLS,
    ]); 

    headerRow.font = { bold: true };  
    headerRow.eachCell((cell) => {
      cell.alignment = { horizontal: 'center', vertical: 'middle', indent: 10 }
    })

    let count = 1;
    
    messagesData.map((data) => {
      const cells = worksheet.addRow([count, data.content, data.messageId, ...blankMetrics]);

      cells.eachCell((cell, colNumber) => {
        const col = worksheet.getColumn(colNumber);
        
        if (col.key === 'text') return;

        cell.alignment = { horizontal: 'center', vertical: 'middle' }
      })

      count += 1
    })

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber <= 3) return;

      const textCell = row.getCell('text')

      textCell.alignment = { wrapText: true }
    })

    const filePath = path.join(BLANK_FOLDER, fileName + '.xlsx')

    // Сохраняем файл
    workbook.xlsx.writeFile(filePath);
  }
  console.log('Excel бланки успешно созданы!');
}

convertToBlanks()