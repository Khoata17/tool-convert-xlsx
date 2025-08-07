// Styles.js
const commonStyle = {
  border: {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } },
  },
};

const headerStyle = {
  ...commonStyle,
  font: { bold: true },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
  fill: { fgColor: { rgb: "FFFF00" } },
};

const dataStyle = { ...commonStyle };

const passedQuizStyle = {
  ...commonStyle,
  font: { color: { rgb: "FF0000" } }, // Red text color for "Bài quiz (Đã đạt)"
};

const specificHeaderStyle = {
  ...headerStyle,
  fill: { fgColor: { rgb: "F4B084" } },
};
