import * as XLSX from 'xlsx';

export interface ExcelData {
  sheets: string[];
  selectedSheet?: string;
  columns?: string[];
  rows?: any[];
  error?: string;
}

export async function parseExcel(file: File): Promise<ExcelData> {
  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheetNames = workbook.SheetNames;
    
    return {
      sheets: sheetNames
    };
  } catch (e) {
    console.error('Error parsing Excel', e);
    return {
      sheets: [],
      error: 'Erro ao ler o arquivo Excel'
    };
  }
}

export function getSheetData(file: ArrayBuffer, sheetName: string): { columns: string[], rows: any[] } {
  const workbook = XLSX.read(file);
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) return { columns: [], rows: [] };
  
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
  if (jsonData.length === 0) return { columns: [], rows: [] };
  
  const columns = jsonData[0].map(c => String(c));
  const rows = jsonData.slice(1).map(row => {
    const obj: any = {};
    columns.forEach((col, index) => {
      obj[col] = row[index];
    });
    return obj;
  });
  
  return { columns, rows };
}
