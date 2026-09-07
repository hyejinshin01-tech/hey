/**
 * App Main Entry Point
 */
document.addEventListener('DOMContentLoaded', () => {
  // Check if SheetJS is available
  if (typeof XLSX === 'undefined') {
    console.error('SheetJS(xlsx) 라이브러리를 불러오지 못했습니다.');
  }

  // Initialize Core and UI
  const excelCore = new ExcelCore();
  const uiController = new UIController(excelCore);

  console.log('✨ ExcelMerge Pro 웹 애플리케이션이 성공적으로 초기화되었습니다.');
});
