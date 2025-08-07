const smallId = "#small_dropzone";
const smallDropzone = document.querySelector(smallId);

var smallPreviewNode = smallDropzone.querySelector(".dropzone-item");
smallPreviewNode.id = "";
var smallPreviewTemplate = smallPreviewNode.parentNode.innerHTML;
smallPreviewNode.parentNode.removeChild(smallPreviewNode);

let smallAlertDisplayed = false;

const downloadButton = document.querySelector(".btn-download");

let staticLopMon = null;
let staticSinhVien = null;
let staticLop = null;

downloadButton.style.display = "none";

const smallMyDropzone = new Dropzone(smallId, {
  url: "/",
  method: "get",
  parallelUploads: 20,
  maxFilesize: 10000,
  maxFiles: 100,
  acceptedFiles: ".xlsx, .csv",
  previewTemplate: smallPreviewTemplate,
  previewsContainer: smallId + " .dropzone-items",
  clickable: smallId + " .dropzone-select",
});

let uploadedFiles = [];
smallMyDropzone.on("addedfile", function (file) {
  uploadedFiles.push(file);
  const smallDropzoneItems = smallDropzone.querySelectorAll(".dropzone-item");
  smallDropzoneItems.forEach((dropzoneItem) => {
    dropzoneItem.style.display = "";
  });
  downloadButton.style.display = "block";
});

// Update the total progress bar
smallMyDropzone.on("totaluploadprogress", function (progress) {
  const smallProgressBars = smallDropzone.querySelectorAll(".progress-bar");
  smallProgressBars.forEach((progressBar) => {
    progressBar.style.width = progress + "%";
  });
});

smallMyDropzone.on("sending", function (file) {
  const smallProgressBars = smallDropzone.querySelectorAll(".progress-bar");
  smallProgressBars.forEach((progressBar) => {
    progressBar.style.opacity = "1";
  });
});

smallMyDropzone.on("complete", function (progress) {
  const smallProgressBars = smallDropzone.querySelectorAll(".dz-complete");

  setTimeout(function () {
    smallProgressBars.forEach((progressBar) => {
      progressBar.querySelector(".progress-bar").style.opacity = "0";
      progressBar.querySelector(".progress").style.opacity = "0";
    });
  }, 300);
});

smallMyDropzone.on("maxfilesexceeded", function (file) {
  smallMyDropzone.removeFile(file);
  if (!smallAlertDisplayed) {
    createToast(
      "error",
      "bi bi-exclamation-circle",
      "Error",
      "You can only upload a maximum of 3 Excel files."
    );
    smallAlertDisplayed = true;
  }
});

smallMyDropzone.on("removedfile", function (file) {
  uploadedFiles = uploadedFiles.filter((f) => f !== file); // Xóa file khỏi mảng
  if (uploadedFiles.length === 0) {
    downloadButton.style.display = "none"; // Ẩn nút download nếu không còn file nào
  }
  if (smallMyDropzone.files.length < smallMyDropzone.options.maxFiles) {
    smallAlertDisplayed = false;
  }
});

downloadButton.addEventListener("click", async function () {
  for (const file of uploadedFiles) {
    await processFile(file);
  }
});

function clearFiles() {
  smallMyDropzone.removeAllFiles(true);
}

document
  .querySelector(smallId + " .dropzone-remove-all")
  .addEventListener("click", function () {
    smallMyDropzone.removeAllFiles(true);
    smallAlertDisplayed = false;
    setTimeout(() => {
      resetSmallCheckboxes();
    }, 0);
  });

let dataHandler = [];
let dataSinhDSSV, dataDSLop, dataDKMon, dataTienDo;
let staticDataLopMon = null;
let staticTienDo = null;
let jsonData = null;
const readFile = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        resolve(jsonData);
      } catch (error) {
        reject(error);
      }
    };
    reader.readAsArrayBuffer(file);
  });
};

let name_file = null;
let globalMa = null;

async function processFile(files) {
  dataHandler = [];

  try {
    // name_file = files.name;
    // console.log(files,name_file);

    name_file = files.name; // Lấy tên file
    console.log("Tên file:", name_file);

    // Trích xuất mã từ tên file
    const match = name_file.match(/_(\w{6,7})_/);
    if (match) {
      globalMa = match[1]; // Gán mã vào biến global
      console.log("Mã tìm được:", globalMa);
    }

    const jsonData = await readFile(files);
    if (!checkMatch(jsonData[0], header_cms)) return;
    // console.log(jsonData);
    const db = await openDatabase();
    const transaction = db.transaction(
      ["DSSVStore", "DSLopStore", "tienDoStore", "DKMonStore"],
      "readonly"
    );
    ////--

    const objectStoreDSSV = transaction.objectStore("DSSVStore");
    const objectStoreDSLop = transaction.objectStore("DSLopStore");
    const objectStoreTienDo = transaction.objectStore("tienDoStore");
    const objectStoreDKMon = transaction.objectStore("DKMonStore");
    const dataSinhDSSV = await getDataFromObjectStore(objectStoreDSSV);
    const dataDSLop = await getDataFromObjectStore(objectStoreDSLop);
    const dataDKMon = await getDataFromObjectStore(objectStoreDKMon);
    const dataTienDo = await getDataFromObjectStore(objectStoreTienDo);

    staticDataLopMon = dataSinhDSSV;
    staticSinhVien = dataDSLop;
    staticLop = dataDKMon;
    staticTienDo = dataTienDo;
    processData(dataSinhDSSV, dataDSLop, dataDKMon, dataTienDo, jsonData);
    // dữ liệu cms
    console.log(jsonData);
  } catch (error) {
    console.error("Error:", error);
  }
}

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = window.indexedDB.open("fpt-tool", 1);
    request.onerror = reject;
    request.onsuccess = (event) => resolve(event.target.result);
  });
}

function getDataFromObjectStore(objectStore) {
  return new Promise((resolve, reject) => {
    const request = objectStore.getAll();
    request.onerror = reject;
    request.onsuccess = (event) => resolve(event.target.result);
  });
}

function splitClassName(className) {
  const match = className.match(/([a-zA-Z]+)([\d.]+)/);
  if (match) {
    return {
      prefix: match[1],
      number: parseFloat(match[2].replace(".", "")),
    };
  } else {
    return { prefix: className, number: 0 };
  }
}

const dataBlock1 = [];
const dataBlock2 = [];

let selectedBlock = null;
// sự kiện thay đổi khi người dùng chọn Block
// document
//   .getElementById("blockSelect")
//   .addEventListener("change", function (event) {
//     selectedBlock = event.target.value;
//     console.log("Block đã chọn:", selectedBlock);
//     console.log(jsonData);
//     processData(dataSinhDSSV, dataDSLop, dataDKMon, dataTienDo, jsonData);
//   });
document
  .getElementById("blockSelect")
  .addEventListener("change", function (event) {
    selectedBlock = event.target.value;
    console.log("Block đã chọn:", selectedBlock);
    processData(
      dataSinhDSSV,
      dataDSLop.filter((item) => item.block === selectedBlock), // Lọc theo block
      dataDKMon,
      dataTienDo,
      jsonData
    );
  });

function processData(dataSinhDSSV, dataDSLop, dataDKMon, dataTienDo, jsonData) {
  let headers = jsonData[0];

  console.log(dataDKMon);
  const data = jsonData.slice(1).map((row) => {
    let rowData = {};
    headers.forEach((header, index) => {
      rowData[header] = row[index];
    });
    return rowData;
  });

  // console.log(dataSinhDSSV)
  // console.log(globalMa)
  /// lấy sinh viên xử lý đang bug 1
  // dữ liệu cms
  data.forEach((item) => {
    const matchingSinhVien = dataSinhDSSV.find(
      (sinh_vien) =>
        item.Email?.toLowerCase() === sinh_vien?.email?.toLowerCase() &&
        sinh_vien.block == selectedBlock &&
        sinh_vien.ma_mon == globalMa
    );
    if (matchingSinhVien && matchingSinhVien !== undefined) {
      // if (matchingSinhVien && matchingSinhVien.block === selectedBlock) {

      // console.log(matchingSinhVien);

      const sinhVienData = {
        ...item,
        ma_sinh_vien: matchingSinhVien.ma_sinh_vien,
        ho_va_ten: matchingSinhVien.ho_va_ten,
        ma_mon: matchingSinhVien.ma_mon,
        id_lop: matchingSinhVien.id_lop,
      };
      // console.log(sinhVienData);
      dataHandler.push(sinhVienData);
    }
  });

  for (let i = dataHandler.length - 1; i >= 0; i--) {
    // đang lấy mã sv từ dssv
    const maSinhVien = dataHandler[i].ma_sinh_vien;
    // console.log(maSinhVien);
    // console.log(dataSinhDSSV);
    const matchingSV = dataSinhDSSV.find(
      (ma_dssv) => ma_dssv.ma_sinh_vien === maSinhVien
    );
    // console.log(matchingSV);
    if (!matchingSV) {
      dataHandler.splice(i, 1);
    }
  }

  /* TÌM LỚP DỰA THEO MÔN*/
  let lop = null;
  const uniqueTenLop = new Set(dataDSLop.map((item) => item.ma_mon));

  // console.log(uniqueTenLop);

  if (name_file) {
    // console.log(name_file);
    try {
      const foundClasses = findClassesInText([...uniqueTenLop], name_file);
      lop = foundClasses;
      // console.log(lop);
    } catch (error) {
      console.error(error.message);
    }
  }

  // mò bug1
  // console.log(dataHandler);
  // console.log(dataDSLop);

  const filteredDSLop = dataDSLop.filter(
    (dslop) => dslop.block === selectedBlock
  );
  // console.log("tên lớp nè:", lop);

  dataHandler.forEach((item, index) => {
    // đã lấy được điểm asm

    const matchingLopMon = filteredDSLop.find(
      (dslop) => item.id_lop === dslop.ten_lop && dslop.ma_mon === lop
    );
    if (matchingLopMon) {
      dataHandler[index] = {
        ...item,
        ...matchingLopMon,
      };
    }
  });

  function findClassesInText(classes, name_file) {
    if (typeof name_file !== "string") {
      throw new Error("Text must be a string");
    }

    // Lọc tất cả mã khớp
    const matchedClasses = classes.filter((className) => {
      if (typeof className !== "string") {
        console.warn(`Class name is not a string: ${className}`);
        return false;
      }
      return name_file.includes(className);
    });

    // Sắp xếp mã khớp theo độ dài giảm dần (ưu tiên mã dài hơn, ví dụ SOF3061 trước SOF306)
    matchedClasses.sort((a, b) => b.length - a.length);

    // Trả về mã đầu tiên (chính xác nhất) hoặc "No class found"
    return matchedClasses.length > 0 ? matchedClasses[0] : "No class found";
  }

  const matchingTienDo = dataTienDo.find((item) => item.ma === lop);

  /// tiến độ
  dataHandler.map((item) => {
    checkProgress(matchingTienDo, item);
  });
  // console.log(dataHandler);

  //TRANG THỨ 2
  const uniqueIdLopSet = new Set();
  let ma_mon;
  // dataHandler.forEach((item) => {
  //   console.log('ma mon dang bug')
  //   console.log(item.ma)
  //   ma_mon = item.ma;
  //   if (item.ten_lop !== undefined) {
  //     uniqueIdLopSet.add(item.ten_lop);
  //   }
  // });

  dataHandler.forEach((item) => {
    if (item.ma_mon && item.ma && item.ma_mon.startsWith(item.ma)) {
      console.log("ma mon hợp lệ:", item.ma_mon);
      ma_mon = item.ma;

      if (item.ten_lop !== undefined) {
        uniqueIdLopSet.add(item.ten_lop);
      }
    } else {
      console.log(
        "ma mon không khớp:",
        item.ma_mon,
        "không thuộc ma:",
        item.ma
      );
    }
  });

  // đúng 1
  // console.log(dataHandler);

  const uniqueIdLopArray = [...uniqueIdLopSet];
  // console.log("các lớp đã có sinh viên học:", uniqueIdLopSet);
  // console.log(uniqueIdLopArray);
  const uniqueIdLopArrayWithField = uniqueIdLopArray.map((item) => {
    return {
      lop: item,
      mon: ma_mon,
    };
  });
  // Lọc dataDSLop trước để chỉ lấy các lớp thuộc selectedBlock
  const filteredDataDSLop = dataDSLop.filter(
    (item) => item.block === selectedBlock
  );
  // bug 11
  // console.log(filteredDataDSLop);
  // console.log(uniqueIdLopArrayWithField);
  const updatedUniqueIdLopArrayWithField = uniqueIdLopArrayWithField
    .map((lop) => {
      // console.log(lop)
      // bug 2 lop
      const matchingLop = filteredDataDSLop.find(
        (item) => item.ten_lop === lop["lop"] && item.ma_mon === ma_mon
        // &&
        // item.ma_mon.startsWith(item.ma) &&
        // item.block === selectedBlock // Chỉ lấy dữ liệu theo block đã chọn
      );

      // console.log(filteredDataDSLop);
      // console.log(matchingLop);
      // Sử dụng toán tử ?. để lấy giá trị matchingLop khi tồn tại
      // Sử dụng toán tử ?., nếu matchingLop không tồn tại sẽ trả về undefined và bị loại bởi filter(Boolean)

      return (
        matchingLop && {
          ...lop,
          id_lop: matchingLop?.ten_lop,
          giang_vien: matchingLop?.giang_vien,
          ngay_bat_dau: matchingLop?.ngay_bat_dau,
          ngay_ket_thuc: matchingLop?.ngay_ket_thuc,
          so_luong_sinh_vien: matchingLop?.so_luong_sinh_vien,
          block: matchingLop.block,
        }
      );
    })
    .filter(Boolean); // Loại bỏ các phần tử là undefined hoặc null

  // console.log(updatedUniqueIdLopArrayWithField);
  const updatedStatistics = updatedUniqueIdLopArrayWithField.map((lop) => {
    const studentsInClass = dataHandler.filter(
      (item) => item.ten_lop === lop.id_lop && item.block === selectedBlock
    );
    // 1
    // console.log(dataHandler);
    // console.log(lop);
    // console.log(studentsInClass);

    const tong_sinh_vien = studentsInClass.length;
    // console.log(tong_sinh_vien);
    const sl_sv = lop.so_luong_sinh_vien;
    const sinh_vien_chua_hoc = studentsInClass.filter(
      (item) => item.noParticipation === "Chưa tham gia học lần nào"
    ).length;
    noParticipation = "Chưa tham gia học lần nào";

    // console.log("Students in Class:", sinh_vien_chua_hoc);
    // console.log("Number of 'sinh_vien_chua_hoc':", sinh_vien_chua_hoc);
    // item quiz111
    let sinh_vien_dang_hoc = 0;
    studentsInClass.forEach((item) => {
      if (
        item.quizzesAttempted >= 0 &&
        item.quizzesAttempted < item.totalQuizzes &&
        item.quizzesNotAttempted !== item.totalQuizzes
      ) {
        // console.log("Đã vào đang học nè !");
        sinh_vien_dang_hoc += 1;
      }
    });

    const sinh_vien_du_dieu_kien_thi = studentsInClass.filter(
      (item) => item.examEligibility === "Đủ điều kiện dự thi"
    ).length;
    const sinh_vien_cham_tien_do = studentsInClass.filter(
      (item) => item.chamTienDo != ""
    ).length;

    const ti_le_chua_tham_du_hoc =
      sl_sv > 0 ? (sinh_vien_chua_hoc / sl_sv) * 100 : 0;
    const ti_le_cham_tien_do =
      sl_sv > 0 ? (sinh_vien_cham_tien_do / sl_sv) * 100 : 0;

    return {
      ...lop,
      sl_sv,
      sinh_vien_chua_hoc,
      sinh_vien_dang_hoc,
      sinh_vien_du_dieu_kien_thi,
      ti_le_chua_tham_du_hoc: ti_le_chua_tham_du_hoc.toFixed(2) + "%",
      sinh_vien_cham_tien_do,
      ti_le_cham_tien_do: ti_le_cham_tien_do.toFixed(2) + "%",
    };
  });

  // lọc lại lấy kỹ hơn để lấy  lớp theo mã môn trong block đã select
  dataHandler = dataHandler.filter(
    (item) =>
      (item.ten_lop !== undefined &&
        item.block == selectedBlock &&
        item.ma_mon.startsWith(item.ma)) ||
      item.ma.startsWith(item.id_lop)
  );
  // console.log(selectedBlock);
  // console.log(dataHandler);

  //TRANG ĐẦU
  generateExcelFile(
    downloadButton,
    dataHandler,
    updatedStatistics,
    selectedBlock
  );
}
function getLargestNumber(str) {
  if (!str) {
    return undefined;
  }
  // Sử dụng RegExp để trích xuất các số từ chuỗi
  const numbers = str.match(/\d+/g);

  // Chuyển đổi các số thành số nguyên
  const numbersInt = numbers.map(Number);

  // Tìm số lớn nhất sử dụng Math.max()
  const largestNumber = Math.max(...numbersInt);

  return largestNumber;
}

function findPositionInTimeRange(
  formattedDate,
  tuNgayLamQuiz,
  deadlineHoanThanhQuiz
) {
  // Chuyển đổi định dạng ngày tháng "dd/mm/yyyy" sang Date object
  const [day, month, year] = formattedDate.split("/");
  const currentDate = new Date(year, month - 1, day);

  // Tìm vị trí bằng cách so sánh ngày tháng hiện tại với các mảng
  for (let i = 0; i < tuNgayLamQuiz.length; i++) {
    const [startDay, startMonth, startYear] = tuNgayLamQuiz[i].split("/");
    const startDate = new Date(startYear, startMonth - 1, startDay);

    const [endDay, endMonth, endYear] = deadlineHoanThanhQuiz[i].split("/");
    const endDate = new Date(endYear, endMonth - 1, endDay);

    if (currentDate >= startDate && currentDate <= endDate) {
      return i;
    }
  }

  // Nếu không nằm trong bất kỳ khoảng nào, trả về null
  return 0;
}

function findPositionInTimeRangeFail(
  formattedDate,
  tuNgayLamQuiz,
  deadlineHoanThanhQuiz
) {
  const [day, month, year] = formattedDate.split("/");
  const currentDate = new Date(year, month - 1, day);

  // Khởi tạo biến để lưu trữ vị trí và khoảng cách nhỏ nhất
  let closestIndex = -1;
  let closestDistance = Number.MAX_SAFE_INTEGER;

  // Tìm khoảng cách nhỏ nhất giữa ngày hiện tại và các ngày trong mảng
  for (let i = 0; i < tuNgayLamQuiz.length; i++) {
    const [startDay, startMonth, startYear] = tuNgayLamQuiz[i].split("/");
    const startDate = new Date(startYear, startMonth - 1, startDay);

    const [endDay, endMonth, endYear] = deadlineHoanThanhQuiz[i].split("/");
    const endDate = new Date(endYear, endMonth - 1, endDay);

    // Tính khoảng cách tuyệt đối giữa ngày hiện tại và khoảng thời gian
    const distanceToStart = Math.abs(currentDate - startDate);
    const distanceToEnd = Math.abs(currentDate - endDate);

    // Kiểm tra nếu khoảng cách đến ngày bắt đầu hoặc kết thúc gần nhất
    if (distanceToStart < closestDistance) {
      closestDistance = distanceToStart;
      closestIndex = i;
    }
    if (distanceToEnd < closestDistance) {
      closestDistance = distanceToEnd;
      closestIndex = i;
    }
  }

  // Trả về vị trí của khoảng thời gian gần nhất
  return closestIndex;
}

// # cách 1 của 6 tuần
function checkWeeklyProgress(progressTemplate, studentProgress, currentWeek) {
  // console.log("quiz nè");
  // console.log(progressTemplate);
  const quizzesPerWeek = {
    1: progressTemplate.tuan1.match(/\d+/g).map(Number),
    2: progressTemplate.tuan2.match(/\d+/g).map(Number),
    3: progressTemplate.tuan3.match(/\d+/g).map(Number),
    4: progressTemplate.tuan4.match(/\d+/g).map(Number),
    5: progressTemplate.tuan5.match(/\d+/g).map(Number),
    6: progressTemplate.tuan6
      ? progressTemplate.tuan6.match(/\d+/g).map(Number)
      : [],
  };
  // console.log(quizzesPerWeek[currentWeek]);
  // console.log(quizzesPerWeek);

  // Giới hạn tuần tối đa là 6
  if (currentWeek > 6) {
    currentWeek = 6;
  }

  const currentWeekQuizzes = quizzesPerWeek[currentWeek] || [];

  let isBehindSchedule = false;
  let quizDangLam = 0;
  let pendingQuiz = "";

  const keys = Object.keys(studentProgress);
  const quizArray = [];

  keys?.forEach((key) => {
    if (key.startsWith("Quiz")) {
      const quizObject = {
        quizKey: key,
        value: studentProgress[key],
      };
      quizArray.push(quizObject);
    }
  });

  let pivot = null;
  for (let i = 0; i < quizArray.length; i++) {
    if (quizArray[i].value === "Not Attempted") {
      pivot = quizArray[i].quizKey;
      break;
    }
  }

  for (const quizNum of currentWeekQuizzes) {
    const quizKey = `Quiz ${quizNum}`;

    if (!quizArray[quizKey] || quizArray[quizKey] === "Not Attempted") {
      isBehindSchedule = true;
      pendingQuiz = quizKey;
    } else {
      quizDangLam++;
    }
  }

  studentProgress.isBehindSchedule = isBehindSchedule;
  studentProgress.quizDangLam = quizDangLam;
  studentProgress.pendingQuiz = isBehindSchedule ? pendingQuiz : "N/A";
  studentProgress.thong_tin_quiz_cham = pivot;

  return studentProgress;
}

// // # cách 2 của 6 tuần
// function checkWeeklyProgress(progressTemplate, studentProgress, currentWeek) {
//   const quizzesPerWeek = {
//       1: progressTemplate.tuan1.match(/\d+/g).map(Number),
//       2: progressTemplate.tuan2.match(/\d+/g).map(Number),
//       3: progressTemplate.tuan3.match(/\d+/g).map(Number),
//       4: progressTemplate.tuan4.match(/\d+/g).map(Number),
//       5: progressTemplate.tuan5.match(/\d+/g).map(Number),
//       6: progressTemplate.tuan6 ? progressTemplate.tuan6.match(/\d+/g).map(Number) : [], // Thêm tuần 6
//   };

//   const allQuizzes = [
//       ...quizzesPerWeek[1],
//       ...quizzesPerWeek[2],
//       ...quizzesPerWeek[3],
//       ...quizzesPerWeek[4],
//       ...quizzesPerWeek[5],
//       ...quizzesPerWeek[6], // Bao gồm cả quiz 10
//   ];

//   let quizDangLam = 0;
//   let quizzesNotAttempted = 0;
//   let quizzesAttempted = 0;

//   allQuizzes.forEach((quiz) => {
//       const quizKey = `Quiz ${quiz}`;
//       if (studentProgress[quizKey] === "Not Attempted") {
//           quizzesNotAttempted++;
//       } else if (studentProgress[quizKey] === "In Progress") {
//           quizDangLam++;
//       } else if (studentProgress[quizKey] === "Completed") {
//           quizzesAttempted++;
//       }
//   });

//   studentProgress.quizzesNotAttempted = quizzesNotAttempted;
//   studentProgress.quizDangLam = quizDangLam;
//   studentProgress.quizzesAttempted = quizzesAttempted;

//   return studentProgress;
// }

async function generateExcelFile(
  downloadButton,
  combinedData,
  updatedUniqueIdLopArrayWithField,
  selectedBlock
) {
  const firstItemMaMon = combinedData[0].ma;
  // console.log(combinedData[0]);
  // @@
  const filteredStudents = staticDataLopMon.filter(
    (student) =>
      student.ma_mon === firstItemMaMon && student.block == selectedBlock
  );
  // console.log(filteredStudents);

  const filteredStudentObjects = [];
  filteredStudents.forEach((student) => {
    const found = combinedData.some(
      (item) => item.ma_sinh_vien === student.ma_sinh_vien
    );

    if (!found) {
      /* FIND EMAIL */
      const sinhVienInfo = staticSinhVien.find(
        (sv) => sv.ten_lop === student.id_lop
      );

      // console.log(student);
      // console.log(staticSinhVien);

      const email = student ? student.email.toLowerCase() : "N/A";
      const name = student ? student.ho_va_ten : "N/A";
      /* FIND LOP */
      // console.log(student);

      let tenLop = "N/A";
      if (sinhVienInfo) {
        // console.log(firstItemMaMon);
        // console.log(sinhVienInfo);
        // console.log(staticSinhVien);
        const lopInfo = staticSinhVien.find(
          (lop) =>
            lop.ma_mon === sinhVienInfo.ma_mon &&
            lop.ten_lop === sinhVienInfo.ten_lop
        );
        tenLop = lopInfo ? lopInfo.ten_lop : "N/A";
        // console.log(tenLop);
      }

      const newStudent = {
        ma_sinh_vien: student.ma_sinh_vien,
        Email: email.toLowerCase() || "N/A",
        ho_va_ten: name || "N/A",
        id_lop: tenLop || "N/A",
        ma: firstItemMaMon || "N/A",
        totalQuizzes: "N/A",
        quizzesNotAttempted: "N/A",
        chamTienDo: "N/A",
        quizDangLam: "N/A",
        quizzesAttempted: "N/A",
        noParticipation: "Sinh viên chưa enroll vào khóa học",
        examEligibility: "Không đủ điều kiện dự thi",
      };

      // console.log(newStudent);
      combinedData.push(newStudent);
      filteredStudentObjects.push(newStudent);
    }
  });

  const calculateStatistics = (lop, filteredStudents) => {
    // console.log(filteredStudents);
    filteredStudents.forEach((student) => {
      // console.log(student);
      // console.log(lop);
      // console.log(student.id_lop === lop.id_lop && lop.mon === student.ma);

      if (student.id_lop === lop.id_lop) {
        lop.tong_sinh_vien += 1;
        lop.sinh_vien_chua_hoc += 1;
        let ti_le_chua_tham_du_hoc =
          lop.sl_sv > 0 ? (lop.sinh_vien_chua_hoc / lop.sl_sv) * 100 : 0;
        lop.ti_le_chua_tham_du_hoc = ti_le_chua_tham_du_hoc.toFixed(2) + "%";
      }
    });
    return lop;
  };

  updatedUniqueIdLopArrayWithField.map((lop) => {
    const statistics = calculateStatistics(lop, filteredStudentObjects);
    return statistics;
  });

  function handleDataExcel() {
    // lấy giá trị điều kiện từ file điều kiện cms;
    combinedData.forEach((item) => {
      // console.log(item.totalQuizzes);
      const maMon = item.ma;

      // Tìm đối tượng trong dataDKMon có 'mon' khớp với 'ma' của combinedData
      const matchingAsm = staticLop.find(
        (dkMonItem) => dkMonItem.mon === maMon
      );

      console.log(matchingAsm);
      console.log(staticLop);
      // console.log(jsonData)
      console.log(item.Assignment);

      // tính điểm hệ số 10 của file cms điểm asm nếu điều kiện không có sẽ trả về "không"
      // Gán giá trị asm nếu tìm thấy, nếu không thì để trống hoặc N/A
      // item.asm = matchingAsm ? matchingAsm.asm : "Không";
      // dùng normalize("NFD") không phân biệt kiểu chữ trong unicode,
      // trim() xóa khoản trắng,
      // tolowerCase chuyệt tất cả các ký tự thành chuỗi để không phân biệt hoa thường
      if (
        matchingAsm.asm.trim().normalize("NFD").toLowerCase() !==
        "không".normalize("NFD")
      ) {
        // Kiểm tra item.Assignment có phải là một số hợp lệ không
        const assignmentValue = parseFloat(item.Assignment);
        item.asm = !isNaN(assignmentValue) ? assignmentValue * 10 : 0;
        if (isNaN(item.asm)) {
          item.asm = 0;
        }

        // xử lý nếu không đủ điểm
        if (item.asm < 5) {
          item.examEligibility = "Không đủ điều kiện dự thi";
          item.chamTienDo = item.chamTienDo
            ? `${item.chamTienDo}, không đủ điểm asm`
            : "Không đủ điểm asm";
          console.log("đã vào tính điểm.");
        }
      }
      // else if(matchingAsm.asm.toLowerCase() === "Không"){
      //   item.asm = "Không";
      // } // hàm này không phù hợp
      else if (
        matchingAsm.asm.trim().normalize("NFD").toLowerCase() ===
        "không".normalize("NFD")
      ) {
        item.asm = "Không DK";

        // console.log(header.v);
        // Xóa cột "Điểm asm" khỏi headers
        // headers = headers.filter(header => header.v.trim().normalize("NFD") !== "Điểm asm".trim().normalize("NFD"));
      }

      // xét điều kiện dự thi với tiến độ dựa trên điều kiện tính điểm của môn nếu bắt buột tính điểm asm mới được dự thi
      // if (
      //   matchingAsm.asm.trim().normalize("NFD").toLowerCase() !==
      //   "không".normalize("NFD")
      // ) {
      //   // item.examEligibility =
      //   // item.chamTienDo =

      //   // 111
      //   // console.log(item);
      // }
    });

    // combinedData.forEach((item) => {
    //   const maMon = item.ma;

    //   // Tìm đối tượng trong dataDKMon có 'mon' khớp với 'ma' của combinedData
    //   const matchingAsm = staticLop.find(
    //     (dkMonItem) => dkMonItem.mon === maMon
    //   );
    //   // Gán giá trị asm nếu tìm thấy, nếu không thì để trống hoặc N/A
    //   item.asm = matchingAsm ? matchingAsm.asm : "Không";

    //   const matchingTienDo = staticTienDo.find((tienDo) => tienDo.ma === maMon);

    //   if (matchingTienDo) {
    //     const quizzes = [
    //       ...matchingTienDo.tuan1.match(/\d+/g).map(Number),
    //       ...matchingTienDo.tuan2.match(/\d+/g).map(Number),
    //       ...matchingTienDo.tuan3.match(/\d+/g).map(Number),
    //       ...matchingTienDo.tuan4.match(/\d+/g).map(Number),
    //       ...matchingTienDo.tuan5.match(/\d+/g).map(Number),
    //       ...(matchingTienDo.tuan6
    //         ? matchingTienDo.tuan6.match(/\d+/g).map(Number)
    //         : []),
    //     ];
    //     item.totalQuizzes = quizzes.length;
    //     console.log(item.totalQuizzes);
    //   } else {
    //     item.totalQuizzes = 0;
    //   }
    // });

    // console.log(combinedData.totalQuizzes);
    // Sắp xếp combinedData dựa trên id_lop
    combinedData.sort((a, b) => {
      const classA = splitClassName(a.ten_lop || a.id_lop);
      const classB = splitClassName(b.ten_lop || a.id_lop);

      if (classA.prefix < classB.prefix) return -1;
      if (classA.prefix > classB.prefix) return 1;
      return classA.number - classB.number;
    });

    if (combinedData.length > 0) {
      const data = combinedData.map((item, index) => [
        { v: index + 1, s: dataStyle },
        { v: item.ma_sinh_vien, s: dataStyle },
        { v: item.Email.toLowerCase(), s: dataStyle },
        { v: item.ho_va_ten, s: dataStyle },
        { v: item.ten_lop || item.id_lop || "Không xác định", s: dataStyle },
        { v: item.ma, s: dataStyle },
        {
          v: item.totalQuizzes,
          s: { ...dataStyle, alignment: { horizontal: "center" } },
        },
        {
          v: item.asm,
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
      // console.log("danh sach ne", data);

      const worksheet = XLSX.utils.aoa_to_sheet([headers, ...data]);

      worksheet["!rows"] = chieu_cao_sheet_1;
      worksheet["!cols"] = chieu_rong_sheet_1;
      const workbook = XLSX.utils.book_new();

      //TRANG 2

      updatedUniqueIdLopArrayWithField.sort((a, b) => {
        const classA = splitClassName(a.lop);
        const classB = splitClassName(b.lop);

        if (classA.prefix < classB.prefix) return -1;
        if (classA.prefix > classB.prefix) return 1;
        return classA.number - classB.number;
      });
      // console.log(updatedUniqueIdLopArrayWithField);

      const data2 = updatedUniqueIdLopArrayWithField.map((item, index) => [
        { v: index + 1, s: dataStyle },
        { v: item.id_lop, s: dataStyle },
        { v: item.mon, s: dataStyle },
        { v: item.giang_vien, s: dataStyle },
        { v: item.ngay_bat_dau || "", s: dataStyle },
        { v: item.ngay_ket_thuc, s: dataStyle },
        { v: item.sl_sv, s: dataStyle },
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
        { v: item.block, s: dataStyle },
      ]);
      // console.log("dữ liệu thống kê lớp môn: ", data2);

      const titleRow2 = [
        {
          v: "THỐNG KÊ TÌNH HÌNH HỌC",
          s: {
            font: { bold: true },
            alignment: { horizontal: "center", vertical: "center" },
          },
        },
      ];

      const worksheetData2 = [titleRow2, headers2, ...data2];

      const worksheet2 = XLSX.utils.aoa_to_sheet(worksheetData2);

      worksheet2["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: headers2.length - 1 } },
      ];

      // Đặt chiều cao cho hàng tiêu đề và các hàng khác theo yêu cầu
      worksheet2["!rows"] = chieu_cao_sheet_2;

      worksheet2["!cols"] = chieu_rong_sheet_2;

      //TRANG 3

      const uniqueGiangVienData = {};
      updatedUniqueIdLopArrayWithField.forEach((item) => {
        // console.log(item);
        const lop_giang_vien = item.id_lop;
        // console.log(lop_giang_vien);
        const giang_vien = item.giang_vien;
        if (!uniqueGiangVienData[giang_vien]) {
          uniqueGiangVienData[giang_vien] = {
            id_lop: [],
            sl_sv: 0,
            sinh_vien_chua_hoc: 0,
            sinh_vien_dang_hoc: 0,
            sinh_vien_du_dieu_kien_thi: 0,
            ti_le_chua_tham_du_hoc: 0,
            sinh_vien_cham_tien_do: 0,
            ti_le_cham_tien_do: 0,
            ti_le_chua_tham_du_hoc_count: 0,
            ti_le_cham_tien_do_count: 0,
          };
        }
        // uniqueGiangVienData[giang_vien].id_lop = lop_giang_vien;

        // Thêm id_lop vào mảng, đảm bảo không trùng lặp
        if (!uniqueGiangVienData[giang_vien].id_lop.includes(lop_giang_vien)) {
          uniqueGiangVienData[giang_vien].id_lop.push(lop_giang_vien);
        }
        uniqueGiangVienData[giang_vien].sl_sv += item.sl_sv;
        uniqueGiangVienData[giang_vien].sinh_vien_chua_hoc +=
          item.sinh_vien_chua_hoc;
        uniqueGiangVienData[giang_vien].sinh_vien_dang_hoc +=
          item.sinh_vien_dang_hoc;
        uniqueGiangVienData[giang_vien].sinh_vien_du_dieu_kien_thi +=
          item.sinh_vien_du_dieu_kien_thi;
        uniqueGiangVienData[giang_vien].sinh_vien_cham_tien_do +=
          item.sinh_vien_cham_tien_do;

        uniqueGiangVienData[giang_vien].ti_le_chua_tham_du_hoc += parseFloat(
          item.ti_le_chua_tham_du_hoc.replace("%", "")
        );
        uniqueGiangVienData[giang_vien].ti_le_chua_tham_du_hoc_count += 1;

        uniqueGiangVienData[giang_vien].ti_le_cham_tien_do += parseFloat(
          item.ti_le_cham_tien_do.replace("%", "")
        );
        uniqueGiangVienData[giang_vien].ti_le_cham_tien_do_count += 1;
      });

      const data3 = Object.keys(uniqueGiangVienData).map((giang_vien) => {
        // console.log(uniqueGiangVienData[giang_vien]);
        const giangVienData = uniqueGiangVienData[giang_vien];

        // Nối các lớp thành chuỗi, cách nhau bởi dấu phẩy
        const idLopString = giangVienData.id_lop.join(", ");

        // console.log("Giảng viên:", giang_vien, "Danh sách lớp:", idLopString);
        return [
          { v: giang_vien, s: dataStyle },
          { v: idLopString, s: dataStyle },
          { v: giangVienData.sl_sv, s: dataStyle },
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
            s: dataStyle,
          },
        ];
      });

      // const totalRow = [
      //   { v: "Total", s: { ...dataStyle, font: { bold: true } } },
      //   { v: "Không", s: { ...dataStyle, font: { bold: true } } },
      //   {
      //     v: data3.reduce((sum, row) => sum + row[2].v, 0),
      //     s: { ...dataStyle, font: { bold: true } },
      //   },
      //   {
      //     v: data3.reduce((sum, row) => sum + row[3].v, 0),
      //     s: { ...dataStyle, font: { bold: true } },
      //   },
      //   {
      //     v: data3.reduce((sum, row) => sum + row[4].v, 0),
      //     s: { ...dataStyle, font: { bold: true } },
      //   },
      //   {
      //     v: data3.reduce((sum, row) => sum + row[5].v, 0),
      //     s: { ...dataStyle, font: { bold: true } },
      //   },
      //   {
      //     v:
      //       data3
      //         .reduce(
      //           (sum, row) => sum + parseFloat(row[6].v.replace("%", "")),
      //           0
      //         )
      //         .toFixed(2) + "%",
      //     s: { ...dataStyle, font: { bold: true } },
      //   },
      //   {
      //     v: data3.reduce((sum, row) => sum + row[7].v, 0),
      //     s: { ...dataStyle, font: { bold: true } },
      //   },
      //   {
      //     v:
      //       data3
      //         .reduce(
      //           (sum, row) => sum + parseFloat(row[8].v.replace("%", "")),
      //           0
      //         )
      //         .toFixed(2) + "%",
      //     s: { ...dataStyle, font: { bold: true } },
      //   },
      // ];
      // data3.push(totalRow);

      const titleRow = [
        {
          v: "THỐNG KÊ GIẢNG VIÊN",
          s: {
            font: { bold: true },
            alignment: { horizontal: "center", vertical: "center" },
          },
        },
      ];

      const worksheetData = [titleRow, ...[headers3], ...data3];

      const worksheet3 = XLSX.utils.aoa_to_sheet(worksheetData);

      worksheet3["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: headers3.length - 1 } },
      ];

      worksheet3["!rows"] = chieu_cao_sheet_3;
      worksheet3["!cols"] = chieu_rong_sheet_3;

      XLSX.utils.book_append_sheet(workbook, worksheet, "DSSV COM");
      XLSX.utils.book_append_sheet(workbook, worksheet2, "Tke lop mon");
      XLSX.utils.book_append_sheet(workbook, worksheet3, "Tke Gv");

      const id_lop = combinedData.find((item) => item.ma)?.ma;
      if (!id_lop) {
        throw new Error("Không tìm thấy id_lop trong combinedData");
      }
      const fileName = `Tiến độ sinh viên ${id_lop}.xlsx`;
      setTimeout(() => {
        XLSX.writeFile(workbook, fileName, {
          bookType: "xlsx",
          type: "array",
        });
        console.log("Excel file downloaded successfully!");
        combinedData = [];
        uniqueIdLopArray = [];
        updatedUniqueIdLopArrayWithField = [];
        fileName = "";
      }, 3200);

      dataHandler = [];
      setTimeout(() => {
        let btn = $(".dl-button"),
          label = btn.find(".label"),
          counter = label.find(".counter");

        setLabel(
          label,
          label.find(".state"),
          label.find(".default"),
          function () {
            counter.removeClass("hide");
            btn.removeClass("done");
          }
        );
      }, 5000);

      return false;
    } else {
      console.log(
        "No data available to download. Please ensure files are uploaded and processed correctly."
      );
    }
  }
  handleDataExcel();
  downloadButton.style.display = "block";
}
