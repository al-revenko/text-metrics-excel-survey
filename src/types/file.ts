import { Workbook } from "exceljs";
import { MessageData } from "./message";

export interface InputFileData {
  fileName: string,
  messagesData: MessageData[],
}

export interface BlankFileData {
  fileName: string,
  workbook: Workbook,
}
