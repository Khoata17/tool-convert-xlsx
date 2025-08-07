// Headers.js
let headers = [
  { v: "STT", s: headerStyle },
  { v: "Mã", s: headerStyle },
  { v: "ACCONT", s: headerStyle },
  { v: "HỌ VÀ TÊN", s: headerStyle },
  { v: "Tên lớp", s: headerStyle },
  { v: "Mã môn", s: headerStyle },
  { v: "∑ Quiz", s: headerStyle },
  { v: "Điểm asm", s: headerStyle },
  {
    v: "∑ Quiz\n(Chưa học)",
    s: {
      ...headerStyle,
      alignment: {
        wrapText: true,
        horizontal: "center",
        vertical: "center",
      },
    },
  },
  {
    v: "∑ Quiz\n(Đang học)",
    s: {
      ...headerStyle,
      alignment: {
        wrapText: true,
        horizontal: "center",
        vertical: "center",
      },
    },
  },
  {
    v: "∑ Quiz\n(Đã đạt)",
    s: {
      ...headerStyle,
      alignment: {
        wrapText: true,
        horizontal: "center",
        vertical: "center",
      },
      font: { ...headerStyle.font, color: { rgb: "FF0000" } },
    },
  },
  { v: "Chưa tham gia học", s: headerStyle },
  { v: "Điều kiện dự thi", s: headerStyle },
  { v: "Tiến độ học tập", s: headerStyle },
];
// Lưu trữ headers vào localStorage 



const headers2 = [
  { v: "STT", s: headerStyle },
  { v: "Lớp", s: headerStyle },
  { v: "Môn", s: headerStyle },
  { v: "Giảng viên", s: headerStyle },
  { v: "Ngày bắt đầu học", s: headerStyle },
  { v: "Ngày kết thúc", s: headerStyle },
  { v: "∑ Sinh viên", s: headerStyle },
  { v: "∑ Sinh viên\n(Chưa học)", s: headerStyle },
  { v: "∑ Sinh viên\n(Đang học)", s: headerStyle },
  { v: "∑ Sinh viên\n(Đủ điều \nkiện thi)", s: headerStyle },
  { v: "∑ Tỷ lệ\n(Chưa tham\ndự học)", s: headerStyle },
  { v: "∑ Sinh viên\n(Chậm tiến độ)", s: specificHeaderStyle },
  { v: "∑ Tỷ lệ\n(Chậm tiến độ)", s: specificHeaderStyle },
  { v: "Block ↓", s: specificHeaderStyle },

];

const headers3 = [
  { v: "GV", s: headerStyle },
  { v: "∑ lớp", s: headerStyle },
  { v: "∑ Sinh viên", s: headerStyle },
  { v: "∑ Sinh viên\n(Chưa học)", s: headerStyle },
  { v: "∑ Sinh viên\n(Đang học)", s: headerStyle },
  { v: "∑ Sinh viên\n(Đủ điều\nkiện thi)", s: headerStyle },
  { v: "Tỷ lệ\n(Chưa tham\ndự học)", s: headerStyle },
  { v: "∑ Sinh viên\n(Chậm tiến độ)", s: specificHeaderStyle },
  { v: "Tỷ lệ\n(Chậm tiến độ)", s: specificHeaderStyle },
];
