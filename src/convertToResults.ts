import fs from 'node:fs';
import { readDataFiles } from './helpers/readDataFiles';
import { OUTPUT_FOLDER } from './const/files';
import { MessageData } from './types/message';
import { Workbook } from 'exceljs';
import { COLUMNS_KEYS, LEGEND_COLUMNS, METRICS_CELLS } from './const/excel';
import path from 'node:path';

async function convertToResults() {
  const inputData = await readDataFiles('inputData')
  const blankData = await readDataFiles('blankData')

  const messagesDataMap: Map<string, MessageData[]> = new Map(
    inputData.map((data) => [data.fileName, data.messagesData])
  )

  if (!fs.existsSync(OUTPUT_FOLDER)) fs.mkdirSync(OUTPUT_FOLDER);

  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet('Результаты');
  worksheet.columns = [
    LEGEND_COLUMNS[0],
    { header: 'Сотрудник', key: 'employee', width: 20 },
    ...LEGEND_COLUMNS.slice(1),
  ];
  worksheet.getCell('A1').value = null;
  
  worksheet.mergeCells('B1:D1');
  
  const sourceDataCell = worksheet.getCell('B1')
  
  sourceDataCell.value = 'ИСХОДНЫЕ ДАННЫЕ';
  sourceDataCell.alignment = { horizontal: 'center', vertical: 'middle' };
  sourceDataCell.font = { bold: true };  
  
  worksheet.mergeCells('E1:X1');
  
  const testDataCell = worksheet.getCell('E1')
  
  testDataCell.value = 'ОЦЕНОЧНЫЕ ДАННЫЕ';
  testDataCell.alignment = { horizontal: 'center', vertical: 'middle' };
  testDataCell.font = { bold: true };  
  const headerCells = [
    '№', 'Сотрудник', 'Текст сообщения', 'id (ключ)',
  ]
  // Добавляем заголовки второй строки
  const headerRow = worksheet.addRow(headerCells);
  METRICS_CELLS.map((metric, index) => {
    const cellNumber = !index ? headerCells.length + 1 + index : headerCells.length + 1 + index * 2
    
    const cell = headerRow.getCell(cellNumber)
    const nextCell = headerRow.getCell(cellNumber + 1)
    worksheet.mergeCells(`${cell.address}:${nextCell.address}`)
    cell.value = metric
  })
  headerRow.font = { bold: true };  
  
  headerRow.eachCell((cell) => {
    cell.alignment = { horizontal: 'center', vertical: 'middle', indent: 10 }
  })
  const metricsHint = Array.from({length: 20}, (_, index) => {
    return (index % 2) + 1
  })
  const neuroHintRow = worksheet.addRow([null, null, null, null, ...metricsHint])
  neuroHintRow.font = { bold: true }
  neuroHintRow.eachCell((cell) => {
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
  })
    
  let count = 1;

  for (const { fileName, workbook: blank } of blankData) {
    const messagesData = messagesDataMap.get(fileName)

    if (!messagesData) {
      console.log(`${fileName}.json - не найден. Пропуск`)
      continue;
    }

    const blankWorkSheet = blank.getWorksheet(1)

    if (!blankWorkSheet) {
      console.log(`worksheet бланка - не найден. Пропуск`)
      continue;
    }

    blankWorkSheet.eachRow((row, rowNumber) => {
      if (rowNumber < 4) return;

      const { value: id } = row.getCell(COLUMNS_KEYS.id);

      const message = messagesData.find((message) => message.messageId === id)

      if (!message) return;

      const neutralAnwser = Number.parseFloat(row.getCell(COLUMNS_KEYS.neutral).value?.toString() ?? '0.00').toFixed(2);
      const happyAnwser = Number.parseFloat(row.getCell(COLUMNS_KEYS.happy).value?.toString() ?? '0.00').toFixed(2);
      const sadAnwser = Number.parseFloat(row.getCell(COLUMNS_KEYS.sad).value?.toString() ?? '0.00').toFixed(2);
      const surpriseAnwser = Number.parseFloat(row.getCell(COLUMNS_KEYS.surprise).value?.toString() ?? '0.00').toFixed(2);
      const fearAnwser = Number.parseFloat(row.getCell(COLUMNS_KEYS.fear).value?.toString() ?? '0.00').toFixed(2);
      const angryAnwser = Number.parseFloat(row.getCell(COLUMNS_KEYS.angry).value?.toString() ?? '0.00').toFixed(2);
      const inappropriateAnwser = Number.parseFloat(row.getCell(COLUMNS_KEYS.inappropriate).value?.toString() ?? '0.00').toFixed(2);
      const negativeAnwser = Number.parseFloat(row.getCell(COLUMNS_KEYS.negative).value?.toString() ?? '0.00').toFixed(2);
      const toxicAnwser = Number.parseFloat(row.getCell(COLUMNS_KEYS.toxic).value?.toString() ?? '0.00').toFixed(2);
      const informalAnwser = Number.parseFloat(row.getCell(COLUMNS_KEYS.informal).value?.toString() ?? '0.00').toFixed(2);

      const rowNum = rowNumber - 3;
      
      const answerSum = [
        neutralAnwser,
        happyAnwser,
        sadAnwser,
        surpriseAnwser,
        fearAnwser,
        angryAnwser,
        inappropriateAnwser,
        negativeAnwser,
        toxicAnwser,
        informalAnwser,
      ].reduce((acc, curr, index) => {
        const colNum = index + 1;
        
        if (Number.isNaN(parseFloat(curr))) {
          console.log(`Обнаружен NaN в ${fileName} строка: ${rowNum} | колонка: ${colNum}`)
        }

        return acc + parseFloat(curr)
      }, 0)

      if (!answerSum) {
        console.log(`${fileName} строка ${rowNum} - пустая. Пропуск`)
        return;
      };

      const neutral = [neutralAnwser, message.neutral ? message.neutral.toFixed(2) : '0.00'];
      const happy = [happyAnwser, message.happy ? message.happy.toFixed(2) : '0.00'];
      const sad = [sadAnwser, message.sad ? message.sad.toFixed(2) : '0.00'];
      const surprise = [surpriseAnwser, message.surprise ? message.surprise.toFixed(2) : '0.00'];
      const fear = [fearAnwser, message.fear ? message.fear.toFixed(2) : '0.00'];
      const angry = [angryAnwser, message.angry ? message.angry.toFixed(2) : '0.00'];
      const inappropriate = [inappropriateAnwser, message.inappropriate ? message.inappropriate.toFixed(2) : '0.00'];
      const negative = [negativeAnwser, message.negative ? message.negative.toFixed(2) : '0.00'];
      const toxic = [toxicAnwser, message.toxic ? message.toxic.toFixed(2) : '0.00'];
      const informal = [informalAnwser, message.informal ? message.informal.toFixed(2) : '0.00'];

      const createdRow = worksheet.addRow([count, fileName, message.content, message.messageId, ...neutral, ...happy, ...sad, ...surprise, ...fear, ...angry, ...inappropriate, ...negative, ...toxic, ...informal]);
      
      createdRow.eachCell((cell, colNumber) => {
        const col = worksheet.getColumn(colNumber);
        
        if (col.number === COLUMNS_KEYS.text + 1) return;

        cell.alignment = { horizontal: 'center', vertical: 'middle' }
      })

      count += 1;
    })    

  }


  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= 3) return;

    const textCell = row.getCell(COLUMNS_KEYS.text + 1)

    textCell.alignment = { wrapText: true }
  })

  const filePath = path.join(OUTPUT_FOLDER, 'results.xlsx')

  workbook.xlsx.writeFile(filePath);

  console.log('Excel бланки и входящие данные успешно конвертированы в итоговый результат!');
}

convertToResults()