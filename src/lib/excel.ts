import * as XLSX from "xlsx";

export type ExcelSutun = { baslik: string; genislik?: number };

export function excelIndir(
  dosyaAdi: string,
  sayfaAdi: string,
  sutunlar: ExcelSutun[],
  satirlar: (string | number)[][],
) {
  const veri = [sutunlar.map((s) => s.baslik), ...satirlar];
  const ws = XLSX.utils.aoa_to_sheet(veri);
  ws["!cols"] = sutunlar.map((s) => ({ wch: s.genislik ?? 18 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sayfaAdi.slice(0, 31));
  XLSX.writeFile(wb, `${dosyaAdi}.xlsx`);
}
