// GenerateExcelFile.js
import XLSX from "xlsx";
import { headers1, headers2, headers3 } from "./headers.js";
import {
  generateWorksheet1,
  generateWorksheet2,
  generateWorksheet3,
} from "./generateWorksheet.js";

export function generateExcelFile(data, data2) {
  const headers = headers1;
  const headers2 = headers2;
  const headers3 = headers3;

  const data1 = generateWorksheet1(data);
  const data2 = generateWorksheet2(data2);
  const uniqueGiangVienData = data2.reduce((acc, item) => {
    if (!acc[item.giang_vien]) {
      acc[item.giang_vien] = {
        tong_sinh_vien: 0,
        sinh_vien_chua_hoc: 0,
        sinh_vien_dang_hoc: 0,
        sinh_vien_du_dieu_kien_thi: 0,
        ti_le_chua_tham_du_hoc: 0,
        ti_le_chua_tham_du_hoc_count: 0,
        sinh_vien_cham_tien_do: 0,
        ti_le_cham_tien_do: 0,
        ti_le_cham_tien_do_count: 0,
      };
    }
    acc[item.giang_vien].tong_sinh_vien += parseInt(item.tong_sinh_vien, 10);
    acc[item.giang_vien].sinh_vien_chua_hoc += parseInt(
      item.sinh_vien_chua_hoc,
      10
    );
    acc[item.giang_vien].sinh_vien_dang_hoc += parseInt(
      item.sinh_vien_dang_hoc,
      10
    );
    acc[item.giang_vien].sinh_vien_du_dieu_kien_thi += parseInt(
      item.sinh_vien_du_dieu_kien_thi,
      10
    );
    acc[item.giang_vien].ti_le_chua_tham_du_hoc += parseFloat(
      item.ti_le_chua_tham_du_hoc
    );
    acc[item.giang_vien].ti_le_chua_tham_du_hoc_count++;
    acc[item.giang_vien].sinh_vien_cham_tien_do += parseInt(
      item.sinh_vien_cham_tien_do,
      10
    );
    acc[item.giang_vien].ti_le_cham_tien_do += parseFloat(
      item.ti_le_cham_tien_do
    );
    acc[item.giang_vien].ti_le_cham_tien_do_count++;
    return acc;
  }, {});

  const data3 = generateWorksheet3(uniqueGiangVienData);

  const ws1 = XLSX.utils.aoa_to_sheet([headers, ...data1]);
  const ws2 = XLSX.utils.aoa_to_sheet([headers2, ...data2]);
  const ws3 = XLSX.utils.aoa_to_sheet([headers3, ...data3]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "Sheet 1");
  XLSX.utils.book_append_sheet(wb, ws2, "Sheet 2");
  XLSX.utils.book_append_sheet(wb, ws3, "Sheet 3");

  XLSX.writeFile(wb, "combined_data.xlsx");
}
