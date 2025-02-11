import { Column } from "exceljs";

export enum COLUMNS_KEYS  {
  count = 1,
  text,
  id,
  neutral,
  happy,
  sad,
  surprise,
  fear,
  angry,
  inappropriate,
  negative,
  toxic,
  informal,
}

export const LEGEND_COLUMNS: Partial<Column>[] = [
  { header: '№', key: 'number', width: 5, alignment: { horizontal: 'center' } },
  { header: 'Текст сообщения', key: 'text', width: 50 },
  { header: 'id (ключ)', key: 'id', width: 22 },
  { header: 'Neutral', key: 'neutral', width: 10 },
  { header: 'Happy', key: 'happy', width: 10 },
  { header: 'Sad', key: 'sad', width: 10 },
  { header: 'Surprise', key: 'surprise', width: 10 },
  { header: 'Fear', key: 'fear', width: 10 },
  { header: 'Angry', key: 'angry', width: 10 },
  { header: 'Inappropriate', key: 'inappropriate', width: 15 },
  { header: 'Negative', key: 'negative', width: 10 },
  { header: 'Toxic', key: 'toxic', width: 10 },
  { header: 'Informal', key: 'informal', width: 10 },
] as const;

export const METRICS_CELLS: string[] = [
  'Neutral', 
  'Happy', 
  'Sad', 
  'Surprise', 
  'Fear', 
  'Angry', 
  'Inappropriate', 
  'Negative', 
  'Toxic', 
  'Informal'
] as const; 

