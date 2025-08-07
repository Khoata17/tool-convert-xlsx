// GenerateWorksheet.js
import {
  commonStyle,
  dataStyle,
  headerStyle,
  specificHeaderStyle,
} from "./styles.js";

export function generateWorksheet1(data) {
  return data.map((item, index) => [
    { v: index + 1, s: dataStyle },
    { v: item.ma_sinh_vien, s: dataStyle },
    { v: item.Email, s: dataStyle },
    { v: item.ho_ten, s: dataStyle },
    { v: item.id_lop || "", s: dataStyle },
    { v: item.ma, s: dataStyle },
    {
      v: item.totalQuizzes,
      s: { ...dataStyle, alignment: { horizontal: "center" } },
    },
    {
      v: item.quizzesNotAttempted,
      s: { ...dataStyle, alignment: { horizontal: "center" } },
    },
    {
      v: item.quizDangLam,
      s: { ...dataStyle, alignment: { horizontal: "center" } },
    },
    {
      v: item.quizzesAttempted,
      s: {
        ...dataStyle,
        alignment: { horizontal: "center" },
        font: { color: { rgb: "FF0000" } },
      },
    },
    {
      v: item.noParticipation,
      s: { ...dataStyle, alignment: { horizontal: "center" } },
    },
    {
      v: item.examEligibility,
      s: { ...dataStyle, alignment: { horizontal: "center" } },
    },
    {
      v:
        item.chamTienDo.replace(
          /Quiz \d+: Quiz \d+/g,
          (match) => match.split(": ")[1]
        ) || "",
      s: { ...dataStyle, alignment: { horizontal: "center" } },
    },
  ]);
}

export function generateWorksheet2(data) {
  return data.map((item, index) => [
    { v: index + 1, s: dataStyle },
    { v: item.lop, s: dataStyle },
    { v: item.mon, s: dataStyle },
    { v: item.giang_vien, s: dataStyle },
    { v: item.ngay_bat_dau || "", s: dataStyle },
    { v: item.ngay_ket_thuc, s: dataStyle },
    { v: item.tong_sinh_vien, s: dataStyle },
    { v: item.sinh_vien_chua_hoc, s: dataStyle },
    { v: item.sinh_vien_dang_hoc, s: dataStyle },
    { v: item.sinh_vien_du_dieu_kien_thi, s: dataStyle },
    {
      v: item.ti_le_chua_tham_du_hoc,
      s: {
        ...dataStyle,
        fill: { fgColor: { rgb: "FFAAAA" } },
      },
    },
    { v: item.sinh_vien_cham_tien_do, s: dataStyle },
    { v: item.ti_le_cham_tien_do, s: dataStyle },
  ]);
}

export function generateWorksheet3(uniqueGiangVienData) {
  const data3 = Object.keys(uniqueGiangVienData).map((giang_vien) => {
    const giangVienData = uniqueGiangVienData[giang_vien];
    return [
      { v: giang_vien, s: dataStyle },
      { v: giangVienData.tong_sinh_vien, s: dataStyle },
      { v: giangVienData.sinh_vien_chua_hoc, s: dataStyle },
      { v: giangVienData.sinh_vien_dang_hoc, s: dataStyle },
      { v: giangVienData.sinh_vien_du_dieu_kien_thi, s: dataStyle },
      {
        v:
          (
            giangVienData.ti_le_chua_tham_du_hoc /
            giangVienData.ti_le_chua_tham_du_hoc_count
          ).toFixed(2) + "%",
        s: {
          ...dataStyle,
          fill: { fgColor: { rgb: "FFAAAA" } },
        },
      },
      { v: giangVienData.sinh_vien_cham_tien_do, s: dataStyle },
      {
        v:
          (
            giangVienData.ti_le_cham_tien_do /
            giangVienData.ti_le_cham_tien_do_count
          ).toFixed(2) + "%",
        s: {
          ...dataStyle,
          fill: { fgColor: { rgb: "FFAAAA" } },
        },
      },
    ];
  });

  return data3;
}
