function checkProgress(progressTemplate, studentProgress) {

  studentProgress.ma = progressTemplate.ma;
  const studentProgressKeys = Object.keys(studentProgress);
  const quizKeys = studentProgressKeys.filter(
    (key) => key.startsWith("Quiz") && key !== "Quiz (Avg)"
  );

  let quizzesAttempted = 0;

  quizKeys.forEach((quizKey) => {
    const value = studentProgress[quizKey];
    if (value !== "Not Attempted" && value >= 1) {
      quizzesAttempted++;
    }
  });

  let quizDangLamChuaDuDiem = 0;

  quizKeys.forEach((quizKey) => {
    const value = studentProgress[quizKey];
    if (value !== "Not Attempted" && value >= 0) {
      if (value < 1) {
        // console.log(value);
        quizDangLamChuaDuDiem++;
      }
    }
  });

  /* ---------------------------------------------------------------OPTION TẠM THỜI--------------------------------- */
  studentProgress.totalQuizzes = getLargestNumber(progressTemplate.tuan6);
  if (quizzesAttempted > studentProgress.totalQuizzes) {
    studentProgress.quizzesAttempted = studentProgress.totalQuizzes;
  } else {
    studentProgress.quizzesAttempted = quizzesAttempted;
  }
  if (quizDangLamChuaDuDiem != 0) {
    studentProgress.quizzesNotAttempted =
      studentProgress.totalQuizzes - quizDangLamChuaDuDiem;
  } else {
    studentProgress.quizzesNotAttempted =
      studentProgress.totalQuizzes - studentProgress.quizzesAttempted;
  }

  // console.log(studentProgress.totalQuizzes, quizDangLamChuaDuDiem);
  /* ---------------------------------------------------------------OPTION TẠM THỜI--------------------------------- */
  const today = new Date();
  const formattedDate = `${today.getDate()}/${
    today.getMonth() + 1
  }/${today.getFullYear()}`;

  const tuNgayLamQuiz = localStorage.getItem("tuNgayLamQuiz");
  const deadlineHoanThanhQuiz = localStorage.getItem("deadlineHoanThanhQuiz");

  let position = findPositionInTimeRange(
    formattedDate,
    JSON.parse(tuNgayLamQuiz),
    JSON.parse(deadlineHoanThanhQuiz)
  );

  if (position === 0) {
    position = findPositionInTimeRangeFail(
      formattedDate,
      JSON.parse(tuNgayLamQuiz),
      JSON.parse(deadlineHoanThanhQuiz)
    );
  }
  let progressField;
 
  if (position === 0) {
    progressField = progressTemplate.tuan1;
  } else if (position === 1) {
    progressField = progressTemplate.tuan2;
  } else if (position === 2) {
    progressField = progressTemplate.tuan3;
  } else if (position === 3) {
    progressField = progressTemplate.tuan4;
  } else if (position === 4) {
    progressField = progressTemplate.tuan5;
  }  else if (position === 5) {
    progressField = progressTemplate.tuan6;
    console.log(progressField);
  }

  if (progressField == undefined) {
    progressField = progressTemplate.tuan6;
  }
  if (progressField) {
    const updatedProgress = checkWeeklyProgress(
      progressTemplate,
      studentProgress,
      position
    );

    /* let chamTienDoDanhSach = [];

    for (let i = 1; i <= updatedProgress.totalQuizzes; i++) {
      let quizKey = `Quiz ${i}: Quiz ${i}`;
      if (updatedProgress[quizKey] == "Not Attempted") {
        chamTienDoDanhSach.push(`quiz ${i}`);
      }
    }

    let chamTienDoChuoi = chamTienDoDanhSach.join(", ");
    if (chamTienDoChuoi == "") {
      for (let i = 1; i <= updatedProgress.totalQuizzes; i++) {
        let quizKey = `Quiz ${i}: Quiz 0${i}`;
        console.log(`lần ${i}`, updatedProgress[quizKey])
        if (updatedProgress[quizKey] == "Not Attempted") {
          chamTienDoDanhSach.push(`quiz ${i}`);
        }
      }
      chamTienDoChuoi = chamTienDoDanhSach.join(", ");
    } */
      let chamTienDoDanhSach = [];

      // Hàm để lấy quiz key phù hợp
      function getQuizKey(updatedProgress, i) {
        let quizKeyWithoutLeadingZero = `Quiz ${i}: Quiz ${i}`;
        let quizKeyWithLeadingZero = `Quiz ${i}: Quiz 0${i}`;
        
        if (updatedProgress[quizKeyWithoutLeadingZero] == "Not Attempted") {
          return quizKeyWithoutLeadingZero;
        } else if (updatedProgress[quizKeyWithLeadingZero] == "Not Attempted") {
          return quizKeyWithLeadingZero;
        } else {
          return null;
        }
      }
      
      for (let i = 1; i <= updatedProgress.totalQuizzes; i++) {
        let quizKey = getQuizKey(updatedProgress, i);
        if (quizKey) {
          chamTienDoDanhSach.push(`quiz ${i}`);
        }
      }
      
      let chamTienDoChuoi = chamTienDoDanhSach.join(", ");
      // console.log(chamTienDoChuoi);
      

    if (
      /* updatedProgress.isBehindSchedule && */
      updatedProgress.thong_tin_quiz_cham != null
    ) {
      studentProgress.chamTienDo = `Chậm tiến độ với ${chamTienDoChuoi}`;
    } else {
      studentProgress.chamTienDo = "";
    }
  } else {
    studentProgress.chamTienDo = "N/A";
  }

  if (
    studentProgress.quizDangLam === 0 &&
    studentProgress.quizzesNotAttempted === studentProgress.totalQuizzes
  ) {
    studentProgress.noParticipation = "Chưa tham gia học lần nào";
  } else {
    studentProgress.noParticipation = "";
  }

  if (quizzesAttempted === studentProgress.totalQuizzes) {
    studentProgress.examEligibility = "Đủ điều kiện dự thi";
  } else {
    studentProgress.examEligibility = "Không đủ điều kiện dự thi";
  }

  return studentProgress;
}

