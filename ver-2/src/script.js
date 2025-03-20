// 儲存學生資料的陣列
let studentData = [];

// 當使用者上傳檔案時，處理文件內容
document.getElementById('fileInput').addEventListener('change', handleFile, false);

function handleFile(event) {
    let file = event.target.files[0];
    let reader = new FileReader();
    reader.onload = function(e) {
        let data = e.target.result;
        let workbook = XLSX.read(data, { type: 'binary' });
        let sheetName = workbook.SheetNames[0];
        let sheet = workbook.Sheets[sheetName];
        let jsonData = XLSX.utils.sheet_to_json(sheet);
        displayStudents(jsonData);
    };
    reader.readAsBinaryString(file);
}

// 顯示學生資料並提供輸入成績的欄位
function displayStudents(data) {
    studentData = data;
    let tbody = document.querySelector('#studentTable tbody');
    tbody.innerHTML = ''; // 清空表格內容

    // 計算全班的各個平均和及格、不及格學生數
    let totalSelectScore = 0;
    let totalAnswerScore = 0;
    let totalScore = 0;
    let totalAdjustedScore = 0;
    let validStudents = 0;
    let failedStudents = 0;
    let passedStudents = 0;

    data.forEach((student, index) => {
        let row = tbody.insertRow();
        row.insertCell(0).textContent = student['學號'];
        row.insertCell(1).textContent = student['姓名'];

        // 選擇題成績輸入框
        let selectScoreCell = row.insertCell(2);
        let selectScoreInput = document.createElement('input');
        selectScoreInput.type = 'number';
        selectScoreInput.value = student['選擇題成績'] || '';
        selectScoreInput.addEventListener('input', function() {
            student['選擇題成績'] = selectScoreInput.value;
            updateRow(row, student);  // 更新顯示的總成績與不及格標記
        });
        selectScoreCell.appendChild(selectScoreInput);

        // 簡答題成績輸入框
        let answerScoreCell = row.insertCell(3);
        let answerScoreInput = document.createElement('input');
        answerScoreInput.type = 'number';
        answerScoreInput.value = student['簡答題成績'] || '';
        answerScoreInput.addEventListener('input', function() {
            student['簡答題成績'] = answerScoreInput.value;
            updateRow(row, student);  // 更新顯示的總成績與不及格標記
        });
        answerScoreCell.appendChild(answerScoreInput);

        // 顯示計算後的原始總成績
        let totalScoreCell = row.insertCell(4);
        totalScoreCell.textContent = student['選擇題成績'] && student['簡答題成績']
            ? (parseFloat(student['選擇題成績']) + parseFloat(student['簡答題成績'])).toFixed(2)
            : '尚未輸入';

        // 選擇題調分倍率輸入框
        let selectMultiplierCell = row.insertCell(5);
        let selectMultiplierInput = document.createElement('input');
        selectMultiplierInput.type = 'number';
        selectMultiplierInput.value = student['選擇題調分倍率'] || 1; // 預設倍率為1
        selectMultiplierInput.addEventListener('input', function() {
            student['選擇題調分倍率'] = selectMultiplierInput.value;
            updateRow(row, student);  // 更新顯示的調分後總成績
        });
        selectMultiplierCell.appendChild(selectMultiplierInput);

        // 顯示調分後的總成績
        let adjustedTotalScoreCell = row.insertCell(6);
        adjustedTotalScoreCell.textContent = student['選擇題成績'] && student['簡答題成績']
            ? ((parseFloat(student['選擇題成績']) * parseFloat(student['選擇題調分倍率'])) + parseFloat(student['簡答題成績'])).toFixed(2)
            : '尚未輸入';

        // 標記不及格者
        if (parseFloat(student['選擇題成績']) + parseFloat(student['簡答題成績']) < 60) {
            row.style.backgroundColor = '#f8d7da'; // 不及格學生標紅
            failedStudents++;
        } else {
            row.style.backgroundColor = ''; // 恢復背景顏色
            passedStudents++;
        }

        // 計算總分和有效學生數量
        if (!isNaN(parseFloat(student['選擇題成績'])) && !isNaN(parseFloat(student['簡答題成績']))) {
            totalSelectScore += parseFloat(student['選擇題成績']);
            totalAnswerScore += parseFloat(student['簡答題成績']);
            totalScore += parseFloat(student['選擇題成績']) + parseFloat(student['簡答題成績']);
            totalAdjustedScore += (parseFloat(student['選擇題成績']) * parseFloat(student['選擇題調分倍率'])) + parseFloat(student['簡答題成績']);
            validStudents++;
        }
    });

    // 顯示各部分平均
    let selectAverage = validStudents > 0 ? (totalSelectScore / validStudents).toFixed(2) : 0;
    let answerAverage = validStudents > 0 ? (totalAnswerScore / validStudents).toFixed(2) : 0;
    let classAverage = validStudents > 0 ? (totalScore / validStudents).toFixed(2) : 0;
    let adjustedAverage = validStudents > 0 ? (totalAdjustedScore / validStudents).toFixed(2) : 0;

    // 更新選擇題、簡答題、總成績平均與調分後總成績平均
    document.getElementById('selectAvg').textContent = selectAverage;
    document.getElementById('answerAvg').textContent = answerAverage;
    document.getElementById('classAvg').textContent = classAverage;
    document.getElementById('adjustedAvg').textContent = adjustedAverage;

    // 更新及格與不及格學生數量
    document.getElementById('passedCount').textContent = passedStudents;
    document.getElementById('failedCount').textContent = failedStudents;
}

// 更新每一列的總成績與調分後總成績
function updateRow(row, student) {
    let totalScoreCell = row.cells[4];
    let selectScore = parseFloat(student['選擇題成績']) || 0;
    let answerScore = parseFloat(student['簡答題成績']) || 0;

    totalScoreCell.textContent = (selectScore + answerScore).toFixed(2);

    let adjustedTotalScoreCell = row.cells[6];
    let multiplier = parseFloat(student['選擇題調分倍率']) || 1;

    adjustedTotalScoreCell.textContent = ((selectScore * multiplier) + answerScore).toFixed(2);

    // 標記不及格者
    if ((selectScore + answerScore) < 60) {
        row.style.backgroundColor = '#f8d7da'; // 不及格學生標紅
    } else {
        row.style.backgroundColor = ''; // 恢復背景顏色
    }

    // 更新平均和及格、不及格統計
    updateClassAverage();
}

// 計算全班平均
function updateClassAverage() {
    let totalSelectScore = 0;
    let totalAnswerScore = 0;
    let totalScore = 0;
    let totalAdjustedScore = 0;
    let validStudents = 0;
    let failedStudents = 0;
    let passedStudents = 0;

    studentData.forEach(student => {
        let selectScore = parseFloat(student['選擇題成績']) || 0;
        let answerScore = parseFloat(student['簡答題成績']) || 0;
        let multiplier = parseFloat(student['選擇題調分倍率']) || 1;

        if (!isNaN(selectScore) && !isNaN(answerScore)) {
            totalSelectScore += selectScore;
            totalAnswerScore += answerScore;
            totalScore += (selectScore + answerScore);
            totalAdjustedScore += (selectScore * multiplier) + answerScore;
            validStudents++;
            if ((selectScore + answerScore) < 60) {
                failedStudents++;
            } else {
                passedStudents++;
            }
        }
    });

    // 顯示各部分平均
    let selectAverage = validStudents > 0 ? (totalSelectScore / validStudents).toFixed(2) : 0;
    let answerAverage = validStudents > 0 ? (totalAnswerScore / validStudents).toFixed(2) : 0;
    let classAverage = validStudents > 0 ? (totalScore / validStudents).toFixed(2) : 0;
    let adjustedAverage = validStudents > 0 ? (totalAdjustedScore / validStudents).toFixed(2) : 0;

    document.getElementById('selectAvg').textContent = selectAverage;
    document.getElementById('answerAvg').textContent = answerAverage;
    document.getElementById('classAvg').textContent = classAverage;
    document.getElementById('adjustedAvg').textContent = adjustedAverage;
    document.getElementById('passedCount').textContent = passedStudents;
    document.getElementById('failedCount').textContent = failedStudents;
}

// 生成Excel檔案，並下載
function generateExcel() {
    let wb = XLSX.utils.book_new();
    
    // 構建學生資料的欄位，包括學號、姓名、選擇題成績、簡答題成績、總成績、選擇題調分倍率、調分後總成績
    let dataForExcel = studentData.map(student => {
        let selectScore = parseFloat(student['選擇題成績']) || 0;
        let answerScore = parseFloat(student['簡答題成績']) || 0;
        let multiplier = parseFloat(student['選擇題調分倍率']) || 1;

        let totalScore = selectScore + answerScore;
        let adjustedTotalScore = (selectScore * multiplier) + answerScore;

        return {
            "學號": student['學號'],
            "姓名": student['姓名'],
            "選擇題成績": selectScore,
            "簡答題成績": answerScore,
            "總成績": totalScore.toFixed(2),
            "選擇題調分倍率": multiplier,
            "調分後總成績": adjustedTotalScore.toFixed(2)
        };
    });

    // 計算平均值與及格、不及格數量
    let totalSelectScore = 0, totalAnswerScore = 0, totalScore = 0, totalAdjustedScore = 0;
    let passedStudents = 0, failedStudents = 0;
    let validStudents = 0;

    studentData.forEach(student => {
        let selectScore = parseFloat(student['選擇題成績']) || 0;
        let answerScore = parseFloat(student['簡答題成績']) || 0;
        let multiplier = parseFloat(student['選擇題調分倍率']) || 1;

        let totalStudentScore = selectScore + answerScore;
        let adjustedTotalStudentScore = (selectScore * multiplier) + answerScore;

        if (!isNaN(selectScore) && !isNaN(answerScore)) {
            totalSelectScore += selectScore;
            totalAnswerScore += answerScore;
            totalScore += totalStudentScore;
            totalAdjustedScore += adjustedTotalStudentScore;
            validStudents++;

            // 計算及格、不及格學生數量
            if (totalStudentScore < 60) {
                failedStudents++;
            } else {
                passedStudents++;
            }
        }
    });

    // 計算各科目的平均
    let selectAverage = validStudents > 0 ? (totalSelectScore / validStudents).toFixed(2) : 0;
    let answerAverage = validStudents > 0 ? (totalAnswerScore / validStudents).toFixed(2) : 0;
    let classAverage = validStudents > 0 ? (totalScore / validStudents).toFixed(2) : 0;
    let adjustedAverage = validStudents > 0 ? (totalAdjustedScore / validStudents).toFixed(2) : 0;

    // 生成統計信息行
    let statisticsRow = {
        "學號": "統計",
        "姓名": "",
        "選擇題成績": selectAverage,
        "簡答題成績": answerAverage,
        "總成績": classAverage,
        "選擇題調分倍率": "",
        "調分後總成績": adjustedAverage
    };

    // 把學生資料和統計信息加入Excel表格
    let ws = XLSX.utils.json_to_sheet(dataForExcel);
    let statisticsRowArr = [statisticsRow];

    // 將統計信息追加到Excel
    XLSX.utils.sheet_add_json(ws, statisticsRowArr, {skipHeader: true, origin: -1});

    // 定義標題行
    const headers = ["學號", "姓名", "選擇題成績", "簡答題成績", "總成績", "選擇題調分倍率", "調分後總成績"];
    ws['!cols'] = [
        { width: 10 }, // 學號
        { width: 15 }, // 姓名
        { width: 15 }, // 選擇題成績
        { width: 15 }, // 簡答題成績
        { width: 15 }, // 總成績
        { width: 15 }, // 選擇題調分倍率
        { width: 15 }, // 調分後總成績
    ];

    // 追加及格、不及格學生數量
    let passFailRow = {
        "學號": "及格學生數",
        "姓名": passedStudents,
        "選擇題成績": "不及格學生數",
        "簡答題成績": failedStudents,
        "總成績": "",
        "選擇題調分倍率": "",
        "調分後總成績": ""
    };

    // 追加這行
    XLSX.utils.sheet_add_json(ws, [passFailRow], {skipHeader: true, origin: -1});

    // 把資料寫入 Excel 工作表
    XLSX.utils.book_append_sheet(wb, ws, '學生成績');
    
    // 觸發下載Excel檔案
    XLSX.writeFile(wb, '學生成績.xlsx');
}
