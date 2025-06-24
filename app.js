// 셀 클릭 시 헤더 하이라이트
document.querySelectorAll("#sheet td").forEach((td) => {
  td.addEventListener("focus", function () {
    // 기존 하이라이트 제거
    document
      .querySelectorAll(".header-highlight")
      .forEach((el) => el.classList.remove("header-highlight"));
    // 현재 셀 위치 파악
    const cell = td;
    const row = cell.parentElement;
    const rowIdx = Array.from(row.parentElement.children).indexOf(row);
    const colIdx = Array.from(row.children).indexOf(cell);
    // 위쪽 헤더, 왼쪽 헤더 하이라이트
    document
      .querySelector(`#sheet thead tr th:nth-child(${colIdx + 1})`)
      .classList.add("header-highlight");
    document
      .querySelector(`#sheet tbody tr:nth-child(${rowIdx + 1}) th`)
      .classList.add("header-highlight");
  });
});

// 엑셀로 내보내기
function exportExcel() {
  const table = document.getElementById("sheet");
  const wb = XLSX.utils.table_to_book(table, { sheet: "Sheet1" });
  XLSX.writeFile(wb, "spreadsheet.xlsx");
}
