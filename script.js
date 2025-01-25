// 中文筆畫數對應表
const strokes = {
  一: 1, 丁: 2, 七: 2, 九: 2, // 示例：需要完整字典映射
  // ... 添加完整的字典
  龍: 16, 龜: 16,
};

// 計算單個字的筆畫數
function getStrokeCount(char) {
  if (/[\u4e00-\u9fa5]/.test(char)) { // 中文字
    return strokes[char] || 20; // 未知筆畫默認20
  } else {
    return char.charCodeAt(0); // 非中文字符按 ASCII 值排序
  }
}

// 計算全名的筆畫數，僅針對第一個字（姓氏）
function calculateStrokes(name) {
  return name ? getStrokeCount(name[0]) : 0; // 使用姓氏筆畫進行排序
}

// 搜尋功能
function filterNames() {
  const input = document.querySelector('.search-box').value.toLowerCase().trim();
  const rows = document.querySelectorAll('#nameTable tbody tr');
  rows.forEach(row => {
    const name = row.cells[0].textContent.toLowerCase().trim();
    row.style.display = name.includes(input) ? '' : 'none';
  });
}

// 點名功能
function markStatus(button, status) {
  const row = button.closest('tr');
  row.className = status;
  updateStats();
}

// 更新統計
function updateStats() {
  const rows = document.querySelectorAll('#nameTable tbody tr');
  let presentCount = 0, absentCount = 0, leaveCount = 0;
  const totalCount = rows.length; // 總人數包括所有行

  rows.forEach(row => {
    if (row.classList.contains('present')) presentCount++;
    else if (row.classList.contains('absent')) absentCount++;
    else if (row.classList.contains('leave')) leaveCount++;
  });

  document.getElementById('presentCount').textContent = presentCount;
  document.getElementById('absentCount').textContent = absentCount;
  document.getElementById('leaveCount').textContent = leaveCount;

  const attendanceRate = totalCount > 0 ? ((presentCount / totalCount) * 100).toFixed(2) : '0.00';
  document.getElementById('attendanceRate').textContent = `${attendanceRate}%`;
}

// 處理文件上傳
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const fileType = file.name.split('.').pop().toLowerCase();
  if (fileType === 'docx') {
    loadWord(file);
  } else if (fileType === 'xlsx') {
    loadExcel(file);
  } else {
    alert('請上傳有效的 Word (.docx) 或 Excel (.xlsx) 文件！');
  }
}

// 匯入 Word 文件
function loadWord(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    mammoth.extractRawText({ arrayBuffer: e.target.result })
      .then(res => {
        const rawText = res.value;
        const names = parseNames(rawText);
        populateTable(names);
      })
      .catch(() => alert('無法解析 Word 文件！'));
  };
  reader.readAsArrayBuffer(file);
}

// 匯入 Excel 文件
function loadExcel(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const workbook = XLSX.read(e.target.result, { type: 'binary' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }).flat();
    const names = parseNames(rows.join('\n'));
    populateTable(names);
  };
  reader.readAsBinaryString(file);
}

// 解析人名
function parseNames(rawText) {
  return rawText
    .replace(/[\u200B-\u200D\uFEFF]/g, '') // 移除隱藏字元
    .replace(/、|，|,|\t/g, '\n') // 替換分隔符為換行符
    .split(/\r?\n/) // 按換行符分割
    .map(name => name.trim()) // 去除前後空格
    .filter(name => /^[\u4e00-\u9fa5a-zA-Z\s]+$/.test(name)); // 過濾掉非有效人名
}

// 表格排序：按姓氏筆畫數排序
function sortTableByStroke() {
  const tableBody = document.querySelector('#nameTable tbody');
  const rows = Array.from(tableBody.querySelectorAll('tr'));

  rows.sort((rowA, rowB) => {
    const nameA = rowA.cells[0].textContent.trim();
    const nameB = rowB.cells[0].textContent.trim();
    const strokeA = calculateStrokes(nameA);
    const strokeB = calculateStrokes(nameB);

    return strokeA === strokeB ? nameA.localeCompare(nameB, 'zh-Hant') : strokeA - strokeB;
  });

  rows.forEach(row => tableBody.appendChild(row));
}

// 動態生成表格
function populateTable(names) {
  const tableBody = document.querySelector('#nameTable tbody');
  const existingNames = Array.from(tableBody.querySelectorAll('tr')).map(
    row => row.cells[0].textContent.trim()
  );

  names.forEach(name => {
    if (!existingNames.includes(name)) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${name}</td>
        <td class="actions">
          <button class="present-btn" onclick="markStatus(this, 'present')">出席</button>
          <button class="absent-btn" onclick="markStatus(this, 'absent')">請假</button>
          <button class="leave-btn" onclick="markStatus(this, 'leave')">未到</button>
        </td>
        <td><button onclick="deleteRow(this)">刪除</button></td>
      `;
      tableBody.appendChild(tr);
    }
  });
  sortTableByStroke(); // 排序表格
  updateStats();
}

// 新增人名
function addName() {
  const nameInput = document.getElementById('newName');
  const name = nameInput.value.trim();
  if (!name) return alert('請輸入名字！');

  const tableBody = document.querySelector('#nameTable tbody');
  const existingNames = Array.from(tableBody.querySelectorAll('tr')).map(
    row => row.cells[0].textContent.trim()
  );

  if (existingNames.includes(name)) {
    const confirmDuplicate = confirm(`名單中已存在 "${name}"，是否仍然新增？`);
    if (!confirmDuplicate) return;
  }

  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td>${name}</td>
    <td class="actions">
      <button class="present-btn" onclick="markStatus(this, 'present')">出席</button>
      <button class="absent-btn" onclick="markStatus(this, 'absent')">請假</button>
      <button class="leave-btn" onclick="markStatus(this, 'leave')">未到</button>
    </td>
    <td><button onclick="deleteRow(this)">刪除</button></td>
  `;
  tableBody.appendChild(tr);

  nameInput.value = '';
  sortTableByStroke();
  updateStats();
}

// 匯出 Excel 檔案
function exportToExcel() {
  const rows = document.querySelectorAll('#nameTable tbody tr');
  const data = [['未到', '請假', '出席']];
  const stats = { present: [], absent: [], leave: [] };

  rows.forEach(row => {
    const name = row.cells[0].textContent.trim();
    if (row.classList.contains('present')) stats.present.push(name);
    else if (row.classList.contains('absent')) stats.absent.push(name);
    else if (row.classList.contains('leave')) stats.leave.push(name);
  });

  const maxLength = Math.max(stats.present.length, stats.absent.length, stats.leave.length);

  for (let i = 0; i < maxLength; i++) {
    data.push([
      stats.leave[i] || '',
      stats.absent[i] || '',
      stats.present[i] || ''
    ]);
  }

  const totalCount = rows.length;
  const attendanceRate = totalCount > 0 ? ((stats.present.length / totalCount) * 100).toFixed(2) : '0.00';
  data.push([]);
  data.push([`總計: 出席: ${stats.present.length}, 請假: ${stats.absent.length}, 未到: ${stats.leave.length}, 出席率: ${attendanceRate}%`]);

  const worksheet = XLSX.utils.aoa_to_sheet(data);
  worksheet['!cols'] = [{ wch: 20 }, { wch: 20 }, { wch: 20 }];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '點名結果');
  XLSX.writeFile(workbook, '點名結果.xlsx');
}
// 刪除功能：刪除一行並顯示二次確認提示
function deleteRow(button) {
  if (confirm('確定要刪除此筆資料嗎？')) {
    const row = button.closest('tr');
    row.remove(); // 刪除該行
    updateStats(); // 更新統計
  }
}

// 清除所有名單
function clearAll() {
  if (confirm('確定要清除所有名單嗎？')) {
    document.querySelector('#nameTable tbody').innerHTML = '';
    document.getElementById('fileInput').value = '';
    updateStats();
  }
}
