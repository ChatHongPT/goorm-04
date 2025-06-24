// 전역 변수
let currentRow = 5;
let currentCol = 4;
const columns = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z",
];

// 초기화
document.addEventListener("DOMContentLoaded", function () {
  initializeEventListeners();
  addKeyboardNavigation();
  initializeButtonListeners();
});

// 버튼 이벤트 리스너 초기화
function initializeButtonListeners() {
  document.getElementById("exportBtn").addEventListener("click", exportExcel);
  document.getElementById("addRowBtn").addEventListener("click", addRow);
  document.getElementById("addColumnBtn").addEventListener("click", addColumn);
  document.getElementById("removeRowBtn").addEventListener("click", removeRow);
  document
    .getElementById("removeColumnBtn")
    .addEventListener("click", removeColumn);
  document.getElementById("clearAllBtn").addEventListener("click", clearAll);
}

// 이벤트 리스너 초기화
function initializeEventListeners() {
  const cellInputs = document.querySelectorAll(".cell-input");

  cellInputs.forEach((input) => {
    input.addEventListener("focus", handleCellFocus);
    input.addEventListener("blur", handleCellBlur);
    input.addEventListener("input", handleCellInput);
  });
}

// 셀 포커스 처리
function handleCellFocus(event) {
  clearHighlights();
  const input = event.target;
  const row = input.dataset.row;
  const col = input.dataset.col;

  // 헤더 하이라이트
  const colHeader = document.querySelector(
    `th.col-header:nth-child(${getColumnIndex(col) + 2})`
  );
  const currentRowHeader = input.closest("tr").querySelector(".row-header");

  if (colHeader) colHeader.classList.add("header-highlight");
  if (currentRowHeader) currentRowHeader.classList.add("header-highlight");

  // 현재 셀 하이라이트
  input.closest("td").classList.add("cell-highlight");
}

// 셀 블러 처리
function handleCellBlur(event) {
  setTimeout(() => {
    if (!document.activeElement.classList.contains("cell-input")) {
      clearHighlights();
    }
  }, 100);
}

// 셀 입력 처리
function handleCellInput(event) {
  const input = event.target;
  // 여기에 데이터 검증이나 포맷팅 로직 추가 가능
}

// 하이라이트 제거
function clearHighlights() {
  document
    .querySelectorAll(".header-highlight, .cell-highlight")
    .forEach((el) => {
      el.classList.remove("header-highlight", "cell-highlight");
    });
}

// 열 인덱스 가져오기
function getColumnIndex(col) {
  return columns.indexOf(col);
}

// 행 추가
function addRow() {
  currentRow++;
  const tbody = document.querySelector("#sheet tbody");
  const newRow = document.createElement("tr");

  newRow.innerHTML = `
    <th class="row-header">${currentRow}</th>
    ${columns
      .slice(0, currentCol)
      .map(
        (col) =>
          `<td><input type="text" class="cell-input" data-row="${currentRow}" data-col="${col}"></td>`
      )
      .join("")}
  `;

  tbody.appendChild(newRow);

  // 새로운 셀에 이벤트 리스너 추가
  newRow.querySelectorAll(".cell-input").forEach((input) => {
    input.addEventListener("focus", handleCellFocus);
    input.addEventListener("blur", handleCellBlur);
    input.addEventListener("input", handleCellInput);
  });

  // 애니메이션 효과
  newRow.style.opacity = "0";
  newRow.style.transform = "translateY(20px)";
  setTimeout(() => {
    newRow.style.transition = "all 0.3s ease";
    newRow.style.opacity = "1";
    newRow.style.transform = "translateY(0)";
  }, 10);
}

// 열 추가
function addColumn() {
  if (currentCol >= columns.length) return;

  const newColLetter = columns[currentCol];
  currentCol++;

  // 헤더에 새 열 추가
  const headerRow = document.querySelector("#sheet thead tr");
  const newHeader = document.createElement("th");
  newHeader.className = "col-header";
  newHeader.textContent = newColLetter;
  headerRow.appendChild(newHeader);

  // 각 행에 새 셀 추가
  const bodyRows = document.querySelectorAll("#sheet tbody tr");
  bodyRows.forEach((row) => {
    const rowNum = row.querySelector(".row-header").textContent;
    const newCell = document.createElement("td");
    newCell.innerHTML = `<input type="text" class="cell-input" data-row="${rowNum}" data-col="${newColLetter}">`;
    row.appendChild(newCell);

    // 이벤트 리스너 추가
    const input = newCell.querySelector(".cell-input");
    input.addEventListener("focus", handleCellFocus);
    input.addEventListener("blur", handleCellBlur);
    input.addEventListener("input", handleCellInput);
  });

  // 애니메이션 효과
  newHeader.style.opacity = "0";
  newHeader.style.transform = "translateX(-20px)";
  setTimeout(() => {
    newHeader.style.transition = "all 0.3s ease";
    newHeader.style.opacity = "1";
    newHeader.style.transform = "translateX(0)";
  }, 10);
}

// 행 제거
function removeRow() {
  if (currentRow <= 1) {
    alert("최소 1개의 행은 유지되어야 합니다.");
    return;
  }

  if (confirm("마지막 행을 삭제하시겠습니까?")) {
    const tbody = document.querySelector("#sheet tbody");
    const lastRow = tbody.lastElementChild;

    if (lastRow) {
      // 삭제 애니메이션
      lastRow.style.transition = "all 0.3s ease";
      lastRow.style.opacity = "0";
      lastRow.style.transform = "translateY(-20px)";

      setTimeout(() => {
        lastRow.remove();
        currentRow--;
      }, 300);
    }
  }
}

// 열 제거
function removeColumn() {
  if (currentCol <= 1) {
    alert("최소 1개의 열은 유지되어야 합니다.");
    return;
  }

  if (confirm("마지막 열을 삭제하시겠습니까?")) {
    const lastColLetter = columns[currentCol - 1];

    // 헤더에서 마지막 열 제거
    const headerRow = document.querySelector("#sheet thead tr");
    const lastHeader = headerRow.lastElementChild;
    if (lastHeader && lastHeader.textContent === lastColLetter) {
      // 삭제 애니메이션
      lastHeader.style.transition = "all 0.3s ease";
      lastHeader.style.opacity = "0";
      lastHeader.style.transform = "translateX(20px)";

      setTimeout(() => {
        lastHeader.remove();
      }, 300);
    }

    // 각 행에서 마지막 셀 제거
    const bodyRows = document.querySelectorAll("#sheet tbody tr");
    bodyRows.forEach((row, index) => {
      const lastCell = row.lastElementChild;
      if (lastCell && lastCell.tagName === "TD") {
        // 삭제 애니메이션 (약간의 지연으로 순차적 효과)
        setTimeout(() => {
          lastCell.style.transition = "all 0.3s ease";
          lastCell.style.opacity = "0";
          lastCell.style.transform = "translateX(20px)";

          setTimeout(() => {
            lastCell.remove();
          }, 300);
        }, index * 50);
      }
    });

    currentCol--;
  }
}

// 전체 삭제
function clearAll() {
  if (confirm("모든 데이터를 삭제하시겠습니까?")) {
    document.querySelectorAll(".cell-input").forEach((input) => {
      input.value = "";
    });

    // 시각적 피드백
    const container = document.querySelector(".container");
    container.style.transform = "scale(0.98)";
    setTimeout(() => {
      container.style.transform = "scale(1)";
    }, 150);
  }
}

// 키보드 내비게이션
function addKeyboardNavigation() {
  document.addEventListener("keydown", function (e) {
    const activeElement = document.activeElement;
    if (!activeElement.classList.contains("cell-input")) return;

    const currentTd = activeElement.closest("td");
    const currentTr = activeElement.closest("tr");
    let targetInput = null;

    switch (e.key) {
      case "ArrowUp":
        e.preventDefault();
        const prevRow = currentTr.previousElementSibling;
        if (prevRow) {
          const cellIndex = Array.from(currentTr.children).indexOf(currentTd);
          targetInput =
            prevRow.children[cellIndex]?.querySelector(".cell-input");
        }
        break;

      case "ArrowDown":
        e.preventDefault();
        const nextRow = currentTr.nextElementSibling;
        if (nextRow) {
          const cellIndex = Array.from(currentTr.children).indexOf(currentTd);
          targetInput =
            nextRow.children[cellIndex]?.querySelector(".cell-input");
        }
        break;

      case "ArrowLeft":
        e.preventDefault();
        const prevCell = currentTd.previousElementSibling;
        if (prevCell && !prevCell.classList.contains("row-header")) {
          targetInput = prevCell.querySelector(".cell-input");
        }
        break;

      case "ArrowRight":
      case "Tab":
        e.preventDefault();
        const nextCell = currentTd.nextElementSibling;
        if (nextCell) {
          targetInput = nextCell.querySelector(".cell-input");
        }
        break;

      case "Enter":
        e.preventDefault();
        const nextRowEnter = currentTr.nextElementSibling;
        if (nextRowEnter) {
          const cellIndex = Array.from(currentTr.children).indexOf(currentTd);
          targetInput =
            nextRowEnter.children[cellIndex]?.querySelector(".cell-input");
        }
        break;
    }

    if (targetInput) {
      targetInput.focus();
      targetInput.select();
    }
  });
}

// Excel 내보내기
function exportExcel() {
  const table = document.getElementById("sheet");

  // 테이블 데이터를 배열로 변환
  const data = [];
  const headers = [];

  // 헤더 추출
  const headerCells = table.querySelectorAll("thead th");
  headerCells.forEach((cell, index) => {
    if (index > 0) {
      // 첫 번째 빈 셀 제외
      headers.push(cell.textContent);
    }
  });
  data.push(headers);

  // 데이터 추출
  const rows = table.querySelectorAll("tbody tr");
  rows.forEach((row) => {
    const rowData = [];
    const inputs = row.querySelectorAll(".cell-input");
    inputs.forEach((input) => {
      rowData.push(input.value || "");
    });
    data.push(rowData);
  });

  // Excel 파일 생성
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "스프레드시트");

  // 파일 다운로드
  const fileName = `스프레드시트_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, fileName);

  // 성공 피드백
  const btn = document.getElementById("exportBtn");
  const originalText = btn.innerHTML;
  btn.innerHTML = "✅ 내보내기 완료!";
  btn.style.background = "linear-gradient(135deg, #51cf66, #40c057)";

  setTimeout(() => {
    btn.innerHTML = originalText;
    btn.style.background = "";
  }, 2000);
}
