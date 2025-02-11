import fs from 'fs';
import path from 'path';
import { BLANK_FOLDER, INPUT_FOLDER } from '../const/files';
import { BlankFileData, InputFileData } from '../types/file';
import { Workbook } from 'exceljs';


export function readDataFiles(type: 'inputData'): Promise<InputFileData[]> 
export function readDataFiles(type: 'blankData'): Promise<BlankFileData[]> 
export function readDataFiles(type: string): Promise<InputFileData[] | BlankFileData[]> {
  return new Promise((resolve, reject) => {
    const folder = type === 'inputData' ? INPUT_FOLDER : BLANK_FOLDER

    fs.readdir(folder, async (err, files) => {
      if (err) throw reject(err);

      if (type === 'inputData') return resolve(makeInputFileData(files, folder))

      return resolve(await makeBlankFileData(files, folder))
    })
  })

}

function makeInputFileData(files: string[], folder: string): InputFileData[] {
  const inputData: InputFileData[] = [];
  
  for (const file of files) {
    if (path.extname(file) !== '.json') continue;

    const pathToFile = path.join(folder, file)
    const content = fs.readFileSync(pathToFile, 'utf-8')
    const fileName = path.parse(file).name

    inputData.push({ fileName, messagesData: JSON.parse(content) })
  }

  return inputData;
}

async function makeBlankFileData(files: string[], folder: string): Promise<BlankFileData[]> {
  const blankData: BlankFileData[] = [];
  
  for (const file of files) {
    if (path.extname(file) !== '.xlsx') continue;

    const pathToFile = path.join(folder, file)
    const workbook = new Workbook();
    
    await workbook.xlsx.readFile(pathToFile);

    const fileName = path.parse(file).name
    
    blankData.push({ fileName, workbook })
  }

  return blankData;
}