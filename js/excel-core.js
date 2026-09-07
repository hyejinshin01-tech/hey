/**
 * ExcelCore - 엑셀 파일 파싱, 규격(스마트 해시) 분석, 데이터 병합 및 내보내기 엔진
 */
class ExcelCore {
  constructor() {
    this.files = []; // { id, file, name, size, sheets: [{ name, rows, headers, signature, ... }] }
    this.groups = []; // { id, name, signature, headers, items: [{ fileId, fileName, sheetName, rowCount, data }], mergedData: [] }
    this.options = {
      headerRowIndex: 1, // 1-based index for header row
      trimWhitespace: true,
      caseSensitive: false,
      sheetProcessingMode: 'first', // 'first', 'all', or 'by_name'
      targetSheetName: '',
      addSourceColumn: true,
      sourceColumnName: '출처_파일명',
      addSheetColumn: false,
      sheetColumnName: '출처_시트명',
      removeDuplicates: false,
      skipEmptyRows: true,
    };
  }

  /**
   * 옵션 업데이트
   */
  updateOptions(newOptions) {
    this.options = { ...this.options, ...newOptions };
    // 옵션 변경 시 재분석
    if (this.files.length > 0) {
      this.analyzeAndGroup();
    }
  }

  /**
   * 파일 추가 및 파싱
   * @param {File[]} fileList 
   * @param {Function} onProgress 
   */
  async addFiles(fileList, onProgress = () => {}) {
    const newFiles = [];
    const total = fileList.length;

    for (let i = 0; i < total; i++) {
      const file = fileList[i];
      // 중복 파일명 방지 (이름과 크기가 동일하면 스킵)
      const isDuplicate = this.files.some(f => f.name === file.name && f.size === file.size);
      if (isDuplicate) continue;

      try {
        const parsedFile = await this.parseSingleFile(file);
        this.files.push(parsedFile);
        newFiles.push(parsedFile);
      } catch (err) {
        console.error(`파일 파싱 실패: ${file.name}`, err);
        throw new Error(`'${file.name}' 파일을 읽는 중 오류가 발생했습니다: ${err.message}`);
      }
      onProgress(i + 1, total, file.name);
    }

    this.analyzeAndGroup();
    return newFiles;
  }

  /**
   * 단일 파일 파싱 (ArrayBuffer -> SheetJS Workbook)
   */
  parseSingleFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, {
            type: 'array',
            cellDates: true,
            cellNF: false,
            cellText: false,
            raw: false
          });

          const sheets = [];
          workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            // 2차원 배열로 추출 (빈 셀도 null/undefined 처리)
            const rawMatrix = XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
              defval: '',
              blankrows: false,
              raw: false
            });

            sheets.push({
              name: sheetName,
              rawMatrix: rawMatrix,
              totalRawRows: rawMatrix.length
            });
          });

          resolve({
            id: 'file_' + Math.random().toString(36).substr(2, 9),
            file: file,
            name: file.name,
            size: file.size,
            sizeFormatted: this.formatFileSize(file.size),
            sheets: sheets
          });
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * 파일 제거
   */
  removeFile(fileId) {
    this.files = this.files.filter(f => f.id !== fileId);
    this.analyzeAndGroup();
  }

  /**
   * 전체 파일 초기화
   */
  clearAll() {
    this.files = [];
    this.groups = [];
  }

  /**
   * 헤더 정규화 및 해시 시그니처 생성
   */
  normalizeHeader(header) {
    if (header === null || header === undefined) return '';
    let str = String(header);
    if (this.options.trimWhitespace) {
      str = str.trim().replace(/\s+/g, ' ');
    }
    if (!this.options.caseSensitive) {
      str = str.toLowerCase();
    }
    return str;
  }

  /**
   * 헤더 배열로부터 고유 시그니처 생성
   */
  generateSignature(headers) {
    return headers
      .map(h => this.normalizeHeader(h))
      .filter(h => h.length > 0)
      .join('__||__');
  }

  /**
   * 파일 분석 및 동일 규격별 자동 그룹화
   */
  analyzeAndGroup() {
    const groupMap = new Map(); // signature -> Group Object
    const headerRowIdx = Math.max(0, this.options.headerRowIndex - 1);

    this.files.forEach(file => {
      let sheetsToProcess = [];
      if (this.options.sheetProcessingMode === 'first') {
        if (file.sheets.length > 0) sheetsToProcess.push(file.sheets[0]);
      } else if (this.options.sheetProcessingMode === 'all') {
        sheetsToProcess = file.sheets;
      } else if (this.options.sheetProcessingMode === 'by_name' && this.options.targetSheetName) {
        const found = file.sheets.find(s => s.name === this.options.targetSheetName);
        if (found) sheetsToProcess.push(found);
      } else {
        if (file.sheets.length > 0) sheetsToProcess.push(file.sheets[0]);
      }

      sheetsToProcess.forEach(sheet => {
        const rawMatrix = sheet.rawMatrix;
        if (!rawMatrix || rawMatrix.length <= headerRowIdx) return;

        // 원본 헤더 행 추출
        const rawHeaderRow = rawMatrix[headerRowIdx] || [];
        // 공백이 아닌 마지막 컬럼 인덱스 찾기
        let lastColIdx = -1;
        for (let i = rawHeaderRow.length - 1; i >= 0; i--) {
          if (rawHeaderRow[i] !== undefined && String(rawHeaderRow[i]).trim() !== '') {
            lastColIdx = i;
            break;
          }
        }

        if (lastColIdx === -1) return; // 헤더가 빈 경우 스킵

        const headers = [];
        for (let c = 0; c <= lastColIdx; c++) {
          const val = rawHeaderRow[c];
          const colName = (val !== undefined && String(val).trim() !== '') 
            ? String(val).trim() 
            : `[열_${c + 1}]`;
          headers.push(colName);
        }

        const signature = this.generateSignature(headers);
        if (!signature) return;

        // 데이터 행 추출 (헤더 다음 행부터)
        const dataRows = [];
        for (let r = headerRowIdx + 1; r < rawMatrix.length; r++) {
          const rowArr = rawMatrix[r] || [];
          // 완전히 빈 행인지 검사
          const isRowEmpty = rowArr.every(cell => cell === undefined || cell === null || String(cell).trim() === '');
          if (this.options.skipEmptyRows && isRowEmpty) continue;

          // 헤더 개수에 맞춰 row 객체 또는 배열 생성
          const rowObj = {};
          let hasData = false;
          headers.forEach((h, colIdx) => {
            const cellVal = rowArr[colIdx] !== undefined ? rowArr[colIdx] : '';
            rowObj[h] = cellVal;
            if (cellVal !== '') hasData = true;
          });

          if (!this.options.skipEmptyRows || hasData) {
            dataRows.push(rowObj);
          }
        }

        const item = {
          fileId: file.id,
          fileName: file.name,
          sheetName: sheet.name,
          rowCount: dataRows.length,
          data: dataRows
        };

        if (!groupMap.has(signature)) {
          groupMap.set(signature, {
            id: 'grp_' + Math.random().toString(36).substr(2, 9),
            name: '',
            signature: signature,
            headers: [...headers],
            items: [item]
          });
        } else {
          groupMap.get(signature).items.push(item);
        }
      });
    });

    // 그룹 리스트 정리 및 병합 데이터 생성
    let groupIndex = 1;
    this.groups = Array.from(groupMap.values()).map(group => {
      // 대표 그룹명 설정 (예: 규격 1 [이름, 연락처, 이메일...])
      const sampleCols = group.headers.slice(0, 3).join(', ') + (group.headers.length > 3 ? ` 외 ${group.headers.length - 3}개` : '');
      group.name = `규격 ${groupIndex} [${sampleCols}]`;
      group.customName = group.name;
      group.exportFileName = `통합_규격${groupIndex}_데이터.xlsx`;
      groupIndex++;

      // 병합 데이터 구축
      this.compileGroupMergedData(group);
      return group;
    });

    return this.groups;
  }

  /**
   * 단일 그룹의 병합 데이터 조립
   */
  compileGroupMergedData(group) {
    let merged = [];
    const seenHashes = new Set();

    group.items.forEach(item => {
      item.data.forEach(row => {
        const enrichedRow = {};

        // 출처 파일명 열 추가 옵션
        if (this.options.addSourceColumn) {
          enrichedRow[this.options.sourceColumnName] = item.fileName;
        }
        // 출처 시트명 열 추가 옵션
        if (this.options.addSheetColumn) {
          enrichedRow[this.options.sheetColumnName] = item.sheetName;
        }

        // 원래 컬럼 데이터 복사
        group.headers.forEach(h => {
          enrichedRow[h] = row[h] !== undefined ? row[h] : '';
        });

        // 중복 제거 검사
        if (this.options.removeDuplicates) {
          // 원래 컬럼 값 기준으로만 해시 검사 (출처 파일명 제외)
          const rowSignature = group.headers.map(h => String(row[h] || '')).join('###');
          if (seenHashes.has(rowSignature)) {
            return; // 중복 건너뜀
          }
          seenHashes.add(rowSignature);
        }

        merged.push(enrichedRow);
      });
    });

    group.mergedData = merged;
    group.totalRows = merged.length;
    group.totalFiles = new Set(group.items.map(i => i.fileName)).size;
  }

  /**
   * 최종 출력용 헤더 목록 (출처 열 포함 여부 반영)
   */
  getGroupDisplayHeaders(group) {
    const headers = [];
    if (this.options.addSourceColumn) headers.push(this.options.sourceColumnName);
    if (this.options.addSheetColumn) headers.push(this.options.sheetColumnName);
    return [...headers, ...group.headers];
  }

  /**
   * 단일 그룹 엑셀 다운로드
   */
  exportGroupToExcel(group, customFileName = null) {
    const displayHeaders = this.getGroupDisplayHeaders(group);
    const ws = XLSX.utils.json_to_sheet(group.mergedData, { header: displayHeaders });

    // 컬럼 너비 자동 조정
    this.autoFitColumns(ws, group.mergedData, displayHeaders);

    const wb = XLSX.utils.book_new();
    const cleanSheetName = (group.customName || '통합데이터').replace(/[:\\\/\?\*\[\]]/g, '_').substring(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, cleanSheetName);

    const fileName = customFileName || group.exportFileName || `${group.name}.xlsx`;
    XLSX.writeFile(wb, fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`);
  }

  /**
   * 단일 그룹 CSV 다운로드 (UTF-8 with BOM for Korean Excel compatibility)
   */
  exportGroupToCsv(group, customFileName = null) {
    const displayHeaders = this.getGroupDisplayHeaders(group);
    const ws = XLSX.utils.json_to_sheet(group.mergedData, { header: displayHeaders });
    const csvContent = XLSX.utils.sheet_to_csv(ws);
    
    // UTF-8 BOM 추가 (Excel에서 한글 깨짐 방지)
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const fileName = (customFileName || group.exportFileName || group.name).replace(/\.xlsx$/i, '') + '.csv';
    
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  /**
   * 모든 그룹을 하나의 엑셀 파일 내 시트별로 분리하여 다운로드
   */
  exportAllGroupsAsMultiSheetExcel(fileName = '모든규격_통합본.xlsx') {
    if (this.groups.length === 0) return;

    const wb = XLSX.utils.book_new();
    const usedSheetNames = new Set();

    this.groups.forEach((group, idx) => {
      const displayHeaders = this.getGroupDisplayHeaders(group);
      const ws = XLSX.utils.json_to_sheet(group.mergedData, { header: displayHeaders });
      this.autoFitColumns(ws, group.mergedData, displayHeaders);

      // 시트명 31자 제한 및 중복 방지
      let sheetName = (group.customName || `규격_${idx + 1}`)
        .replace(/[:\\\/\?\*\[\]]/g, '_')
        .substring(0, 28);
      
      let finalSheetName = sheetName;
      let counter = 1;
      while (usedSheetNames.has(finalSheetName) || !finalSheetName) {
        finalSheetName = `${sheetName.substring(0, 25)}_${counter++}`;
      }
      usedSheetNames.add(finalSheetName);

      XLSX.utils.book_append_sheet(wb, ws, finalSheetName);
    });

    XLSX.writeFile(wb, fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`);
  }

  /**
   * 전체 그룹 각각을 개별 엑셀 파일로 ZIP 압축 다운로드 (JSZip 사용)
   */
  async exportAllAsZip(zipFileName = '규격별_엑셀_통합본_전체.zip', onProgress = () => {}) {
    if (!window.JSZip) {
      throw new Error('JSZip 라이브러리가 로드되지 않았습니다.');
    }
    if (this.groups.length === 0) return;

    const zip = new JSZip();

    for (let i = 0; i < this.groups.length; i++) {
      const group = this.groups[i];
      const displayHeaders = this.getGroupDisplayHeaders(group);
      const ws = XLSX.utils.json_to_sheet(group.mergedData, { header: displayHeaders });
      this.autoFitColumns(ws, group.mergedData, displayHeaders);

      const wb = XLSX.utils.book_new();
      const sheetName = (group.customName || '통합데이터').replace(/[:\\\/\?\*\[\]]/g, '_').substring(0, 31);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);

      // 엑셀 바이너리 버퍼 생성
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const fileName = group.exportFileName || `규격_${i + 1}_통합.xlsx`;
      zip.file(fileName, excelBuffer);

      onProgress(i + 1, this.groups.length, fileName);
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(zipBlob);
    link.download = zipFileName.endsWith('.zip') ? zipFileName : `${zipFileName}.zip`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  /**
   * 컬럼 너비 자동 계산 헬퍼
   */
  autoFitColumns(worksheet, data, headers) {
    if (!worksheet || !headers || headers.length === 0) return;

    const colWidths = headers.map(header => {
      let maxLen = this.getStringDisplayLength(header);
      // 상위 최대 100개 행을 샘플링하여 최대 너비 계산
      const sampleRows = data.slice(0, 100);
      sampleRows.forEach(row => {
        const val = row[header];
        if (val !== undefined && val !== null) {
          const len = this.getStringDisplayLength(String(val));
          if (len > maxLen) maxLen = len;
        }
      });
      return { wch: Math.min(Math.max(maxLen + 4, 10), 60) };
    });

    worksheet['!cols'] = colWidths;
  }

  /**
   * 한글/전각 문자를 고려한 글자 폭 계산
   */
  getStringDisplayLength(str) {
    let len = 0;
    for (let i = 0; i < str.length; i++) {
      const code = str.charCodeAt(i);
      // 한글 및 전각 문자
      if ((code >= 0xac00 && code <= 0xd7a3) || (code >= 0x1100 && code <= 0x11ff) || code > 0xff) {
        len += 2;
      } else {
        len += 1;
      }
    }
    return len;
  }

  /**
   * 파일 크기 포맷팅
   */
  formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }
}

window.ExcelCore = ExcelCore;
