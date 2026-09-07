/**
 * UIController - 화면 렌더링, 사용자 인터랙션, 모달, 토스트 제어
 */
class UIController {
  constructor(excelCore) {
    this.core = excelCore;
    this.currentPreviewGroup = null;
    this.previewSearchQuery = '';
    this.previewCurrentPage = 1;
    this.previewPageSize = 25;

    this.initDOMElements();
    this.bindEvents();
    this.initTheme();
  }

  initDOMElements() {
    // Dropzone & File List
    this.dropzone = document.getElementById('dropzone');
    this.fileInput = document.getElementById('fileInput');
    this.uploadedFilesContainer = document.getElementById('uploadedFilesContainer');
    this.filesGrid = document.getElementById('filesGrid');
    this.fileCountLabel = document.getElementById('fileCountLabel');
    this.clearAllBtn = document.getElementById('clearAllBtn');
    this.sampleDataBtn = document.getElementById('sampleDataBtn');
    this.themeToggleBtn = document.getElementById('themeToggleBtn');

    // Options
    this.headerRowInput = document.getElementById('optHeaderRow');
    this.addSourceColCheck = document.getElementById('optAddSourceCol');
    this.addSheetColCheck = document.getElementById('optAddSheetCol');
    this.removeDuplicatesCheck = document.getElementById('optRemoveDuplicates');
    this.skipEmptyRowsCheck = document.getElementById('optSkipEmptyRows');
    this.sheetModeSelect = document.getElementById('optSheetMode');

    // Groups & Results
    this.groupsContainer = document.getElementById('groupsContainer');
    this.groupsSummaryBar = document.getElementById('groupsSummaryBar');
    this.totalGroupsCountEl = document.getElementById('totalGroupsCount');
    this.totalMergedRowsEl = document.getElementById('totalMergedRows');
    this.emptyState = document.getElementById('emptyState');

    // Batch Actions Bar
    this.batchActionsBar = document.getElementById('batchActionsBar');
    this.batchGroupCountEl = document.getElementById('batchGroupCount');
    this.batchTotalRowsEl = document.getElementById('batchTotalRows');
    this.downloadZipBtn = document.getElementById('downloadZipBtn');
    this.downloadMultiSheetBtn = document.getElementById('downloadMultiSheetBtn');

    // Preview Modal
    this.previewModal = document.getElementById('previewModal');
    this.modalCloseBtn = document.getElementById('modalCloseBtn');
    this.modalTitle = document.getElementById('modalTitle');
    this.modalTableSearch = document.getElementById('modalTableSearch');
    this.modalTableHead = document.getElementById('modalTableHead');
    this.modalTableBody = document.getElementById('modalTableBody');
    this.modalTotalRows = document.getElementById('modalTotalRows');
    this.modalFilteredRows = document.getElementById('modalFilteredRows');
    this.modalPageInfo = document.getElementById('modalPageInfo');
    this.modalPrevPage = document.getElementById('modalPrevPage');
    this.modalNextPage = document.getElementById('modalNextPage');
    this.modalPageSize = document.getElementById('modalPageSize');
    this.modalExportExcelBtn = document.getElementById('modalExportExcelBtn');
    this.modalExportCsvBtn = document.getElementById('modalExportCsvBtn');

    // Toast Container
    this.toastContainer = document.getElementById('toastContainer');
  }

  bindEvents() {
    // Theme Toggle
    this.themeToggleBtn?.addEventListener('click', () => this.toggleTheme());

    // File Input / Drag and Drop
    this.dropzone.addEventListener('click', () => this.fileInput.click());
    this.fileInput.addEventListener('change', (e) => this.handleFilesSelected(e.target.files));

    ['dragenter', 'dragover'].forEach(name => {
      this.dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(name => {
      this.dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.dropzone.classList.remove('dragover');
      });
    });

    this.dropzone.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      if (dt && dt.files && dt.files.length > 0) {
        this.handleFilesSelected(dt.files);
      }
    });

    // Clear All
    this.clearAllBtn?.addEventListener('click', () => {
      if (this.core.files.length === 0) return;
      if (confirm('업로드된 모든 파일 및 분석 결과를 초기화하시겠습니까?')) {
        this.core.clearAll();
        this.renderAll();
        this.showToast('모든 파일이 초기화되었습니다.', 'info');
      }
    });

    // Sample Data Generation
    this.sampleDataBtn?.addEventListener('click', async () => {
      try {
        this.showToast('샘플 엑셀 파일 6개(3가지 규격)를 생성 중입니다...', 'info');
        const sampleFiles = await SampleGenerator.generateSampleFiles();
        await this.core.addFiles(sampleFiles);
        this.renderAll();
        this.showToast('샘플 데이터가 성공적으로 로드되었습니다!', 'success');
      } catch (err) {
        this.showToast(`샘플 로드 실패: ${err.message}`, 'error');
      }
    });

    // Options Listeners
    this.headerRowInput?.addEventListener('change', (e) => {
      const val = parseInt(e.target.value, 10) || 1;
      this.core.updateOptions({ headerRowIndex: val });
      this.renderGroups();
      this.showToast(`헤더 기준 행이 ${val}행으로 변경되었습니다.`, 'info');
    });

    this.addSourceColCheck?.addEventListener('change', (e) => {
      this.core.updateOptions({ addSourceColumn: e.target.checked });
      this.renderGroups();
    });

    this.addSheetColCheck?.addEventListener('change', (e) => {
      this.core.updateOptions({ addSheetColumn: e.target.checked });
      this.renderGroups();
    });

    this.removeDuplicatesCheck?.addEventListener('change', (e) => {
      this.core.updateOptions({ removeDuplicates: e.target.checked });
      this.renderGroups();
      if (e.target.checked) {
        this.showToast('동일한 행 데이터에 대한 중복 제거가 적용되었습니다.', 'info');
      }
    });

    this.skipEmptyRowsCheck?.addEventListener('change', (e) => {
      this.core.updateOptions({ skipEmptyRows: e.target.checked });
      this.renderGroups();
    });

    this.sheetModeSelect?.addEventListener('change', (e) => {
      this.core.updateOptions({ sheetProcessingMode: e.target.value });
      this.renderGroups();
    });

    // Batch Downloads
    this.downloadZipBtn?.addEventListener('click', async () => {
      try {
        this.downloadZipBtn.disabled = true;
        this.downloadZipBtn.innerHTML = `<span>압축 중...</span>`;
        await this.core.exportAllAsZip('규격별_엑셀_통합본_전체.zip', (cur, total, name) => {
          this.showToast(`ZIP 압축 중: ${cur}/${total} (${name})`, 'info');
        });
        this.showToast('ZIP 압축 파일 다운로드가 완료되었습니다.', 'success');
      } catch (err) {
        this.showToast(`압축 다운로드 실패: ${err.message}`, 'error');
      } finally {
        this.downloadZipBtn.disabled = false;
        this.downloadZipBtn.innerHTML = `
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
          전체 ZIP 압축 다운로드
        `;
      }
    });

    this.downloadMultiSheetBtn?.addEventListener('click', () => {
      try {
        this.core.exportAllGroupsAsMultiSheetExcel('모든규격_통합본_시트별분리.xlsx');
        this.showToast('모든 규격이 포함된 다중 시트 엑셀 파일이 다운로드되었습니다.', 'success');
      } catch (err) {
        this.showToast(`다운로드 실패: ${err.message}`, 'error');
      }
    });

    // Modal Events
    this.modalCloseBtn?.addEventListener('click', () => this.closePreviewModal());
    this.previewModal?.addEventListener('click', (e) => {
      if (e.target === this.previewModal) this.closePreviewModal();
    });

    this.modalTableSearch?.addEventListener('input', (e) => {
      this.previewSearchQuery = e.target.value.toLowerCase();
      this.previewCurrentPage = 1;
      this.renderModalTableData();
    });

    this.modalPageSize?.addEventListener('change', (e) => {
      this.previewPageSize = parseInt(e.target.value, 10);
      this.previewCurrentPage = 1;
      this.renderModalTableData();
    });

    this.modalPrevPage?.addEventListener('click', () => {
      if (this.previewCurrentPage > 1) {
        this.previewCurrentPage--;
        this.renderModalTableData();
      }
    });

    this.modalNextPage?.addEventListener('click', () => {
      this.previewCurrentPage++;
      this.renderModalTableData();
    });

    this.modalExportExcelBtn?.addEventListener('click', () => {
      if (this.currentPreviewGroup) {
        this.core.exportGroupToExcel(this.currentPreviewGroup);
        this.showToast(`'${this.currentPreviewGroup.customName}' 엑셀 다운로드 시작!`, 'success');
      }
    });

    this.modalExportCsvBtn?.addEventListener('click', () => {
      if (this.currentPreviewGroup) {
        this.core.exportGroupToCsv(this.currentPreviewGroup);
        this.showToast(`'${this.currentPreviewGroup.customName}' CSV 다운로드 시작!`, 'success');
      }
    });
  }

  /**
   * 테마 초기화 및 토글
   */
  initTheme() {
    const savedTheme = localStorage.getItem('app_theme') || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);
    this.updateThemeIcon(savedTheme);
  }

  toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme') || 'dark';
    const next = current === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('app_theme', next);
    this.updateThemeIcon(next);
  }

  updateThemeIcon(theme) {
    if (!this.themeToggleBtn) return;
    if (theme === 'light') {
      this.themeToggleBtn.innerHTML = `
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg>
      `;
      this.themeToggleBtn.title = '다크 모드로 전환';
    } else {
      this.themeToggleBtn.innerHTML = `
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>
      `;
      this.themeToggleBtn.title = '라이트 모드로 전환';
    }
  }

  /**
   * 파일 선택 핸들러
   */
  async handleFilesSelected(files) {
    if (!files || files.length === 0) return;
    const validFiles = Array.from(files).filter(f => 
      f.name.endsWith('.xlsx') || f.name.endsWith('.xls') || f.name.endsWith('.csv')
    );

    if (validFiles.length === 0) {
      this.showToast('지원되지 않는 파일 형식입니다. (.xlsx, .xls, .csv 가능)', 'error');
      return;
    }

    try {
      this.showToast(`${validFiles.length}개의 파일을 분석 중입니다...`, 'info');
      await this.core.addFiles(validFiles);
      this.renderAll();
      this.showToast(`${validFiles.length}개 파일 분석 및 규격 분류가 완료되었습니다!`, 'success');
    } catch (err) {
      this.showToast(`파일 처리 중 오류: ${err.message}`, 'error');
    } finally {
      this.fileInput.value = ''; // Reset file input
    }
  }

  /**
   * 전체 렌더링
   */
  renderAll() {
    this.renderUploadedFiles();
    this.renderGroups();
  }

  /**
   * 업로드된 파일 목록 렌더링
   */
  renderUploadedFiles() {
    const files = this.core.files;
    if (files.length === 0) {
      this.uploadedFilesContainer.style.display = 'none';
      return;
    }

    this.uploadedFilesContainer.style.display = 'block';
    this.fileCountLabel.textContent = `총 ${files.length}개 파일`;
    this.filesGrid.innerHTML = '';

    files.forEach(file => {
      const chip = document.createElement('div');
      chip.className = 'file-chip';
      const sheetCountText = file.sheets.length > 1 ? ` · 시트 ${file.sheets.length}개` : '';
      
      chip.innerHTML = `
        <div class="file-chip-info">
          <div class="file-icon">
            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="8" y1="13" x2="16" y2="13"></line><line x1="8" y1="17" x2="16" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
          </div>
          <div class="file-name-block">
            <div class="file-name" title="${file.name}">${file.name}</div>
            <div class="file-meta">${file.sizeFormatted}${sheetCountText}</div>
          </div>
        </div>
        <button class="file-remove-btn" title="파일 삭제" data-file-id="${file.id}">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
        </button>
      `;

      chip.querySelector('.file-remove-btn').addEventListener('click', (e) => {
        e.stopPropagation();
        this.core.removeFile(file.id);
        this.renderAll();
        this.showToast(`'${file.name}' 파일이 삭제되었습니다.`, 'info');
      });

      this.filesGrid.appendChild(chip);
    });
  }

  /**
   * 규격 그룹 리스트 렌더링
   */
  renderGroups() {
    const groups = this.core.groups;
    this.groupsContainer.innerHTML = '';

    if (groups.length === 0) {
      this.emptyState.style.display = 'block';
      this.groupsSummaryBar.style.display = 'none';
      this.batchActionsBar.style.display = 'none';
      return;
    }

    this.emptyState.style.display = 'none';
    this.groupsSummaryBar.style.display = 'flex';
    this.batchActionsBar.style.display = 'flex';

    // Summary Stats
    const totalRows = groups.reduce((acc, g) => acc + g.totalRows, 0);
    this.totalGroupsCountEl.textContent = `${groups.length}개`;
    this.totalMergedRowsEl.textContent = `${totalRows.toLocaleString()}개`;
    this.batchGroupCountEl.textContent = `${groups.length}개`;
    this.batchTotalRowsEl.textContent = `${totalRows.toLocaleString()}개`;

    // Group Cards
    groups.forEach((group, index) => {
      const card = document.createElement('div');
      card.className = 'group-card';

      // Header Column Tags
      const displayHeaders = this.core.getGroupDisplayHeaders(group);
      const colTagsHtml = displayHeaders.map(h => {
        const isSource = h === this.core.options.sourceColumnName || h === this.core.options.sheetColumnName;
        return `<span class="col-tag ${isSource ? 'source-tag' : ''}">${h}</span>`;
      }).join('');

      // Files badges
      const filesBadgesHtml = group.items.map(item => `
        <span class="group-file-badge" title="${item.fileName} (${item.rowCount}행)">
          📄 ${item.fileName} (${item.rowCount.toLocaleString()}행)
        </span>
      `).join('');

      card.innerHTML = `
        <div class="group-card-header">
          <div class="group-title-area">
            <span class="group-badge">규격 ${index + 1}</span>
            <input type="text" class="group-name-input" value="${group.customName}" title="클릭하여 그룹명 수정" />
            <span style="font-size:0.75rem; color:var(--text-muted); font-weight:600;">(총 ${group.totalRows.toLocaleString()}행 · 파일 ${group.items.length}개)</span>
          </div>
          <div class="group-card-actions">
            <button class="btn btn-secondary preview-btn" data-group-id="${group.id}">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
              미리보기
            </button>
            <button class="btn btn-primary export-excel-btn" data-group-id="${group.id}">
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
              엑셀 다운로드
            </button>
          </div>
        </div>

        <div class="group-columns-wrapper">
          <div class="column-tags">${colTagsHtml}</div>
        </div>

        <div class="group-files-info">
          <div class="group-files-list">
            <span style="font-size:0.72rem; font-weight:600; color:var(--text-muted);">포함 파일:</span>
            ${filesBadgesHtml}
          </div>
        </div>
      `;

      // Group Name Change
      const nameInput = card.querySelector('.group-name-input');
      nameInput.addEventListener('change', (e) => {
        group.customName = e.target.value.trim() || group.name;
        group.exportFileName = `${group.customName}.xlsx`;
      });

      // Preview Button
      card.querySelector('.preview-btn').addEventListener('click', () => {
        this.openPreviewModal(group);
      });

      // Export Button
      card.querySelector('.export-excel-btn').addEventListener('click', () => {
        this.core.exportGroupToExcel(group);
        this.showToast(`'${group.customName}' 엑셀 다운로드 시작!`, 'success');
      });

      this.groupsContainer.appendChild(card);
    });
  }

  /**
   * 미리보기 모달 열기
   */
  openPreviewModal(group) {
    this.currentPreviewGroup = group;
    this.previewSearchQuery = '';
    this.previewCurrentPage = 1;
    this.modalTableSearch.value = '';

    this.modalTitle.textContent = `${group.customName} - 병합 데이터 미리보기`;
    this.modalTotalRows.textContent = `${group.totalRows.toLocaleString()}개`;

    // Render Table Header
    const displayHeaders = this.core.getGroupDisplayHeaders(group);
    this.modalTableHead.innerHTML = `
      <tr>
        <th class="row-index-cell">#</th>
        ${displayHeaders.map(h => {
          const isSource = h === this.core.options.sourceColumnName || h === this.core.options.sheetColumnName;
          return `<th class="${isSource ? 'col-source-header' : ''}">${h}</th>`;
        }).join('')}
      </tr>
    `;

    this.renderModalTableData();
    this.previewModal.classList.add('active');
    document.body.style.overflow = 'hidden';
  }

  /**
   * 미리보기 모달 닫기
   */
  closePreviewModal() {
    this.previewModal.classList.remove('active');
    document.body.style.overflow = '';
    this.currentPreviewGroup = null;
  }

  /**
   * 미리보기 모달 테이블 데이터 렌더링 (검색 & 페이징)
   */
  renderModalTableData() {
    if (!this.currentPreviewGroup) return;

    const group = this.currentPreviewGroup;
    const displayHeaders = this.core.getGroupDisplayHeaders(group);
    let rows = group.mergedData;

    // Search Filter
    if (this.previewSearchQuery) {
      const q = this.previewSearchQuery;
      rows = rows.filter(row => {
        return displayHeaders.some(h => {
          const val = row[h];
          return val !== undefined && val !== null && String(val).toLowerCase().includes(q);
        });
      });
    }

    this.modalFilteredRows.textContent = rows.length.toLocaleString();

    // Pagination
    const totalPages = Math.max(1, Math.ceil(rows.length / this.previewPageSize));
    if (this.previewCurrentPage > totalPages) this.previewCurrentPage = totalPages;

    const startIndex = (this.previewCurrentPage - 1) * this.previewPageSize;
    const pageRows = rows.slice(startIndex, startIndex + this.previewPageSize);

    this.modalPageInfo.textContent = `${this.previewCurrentPage} / ${totalPages} 페이지`;
    this.modalPrevPage.disabled = this.previewCurrentPage <= 1;
    this.modalNextPage.disabled = this.previewCurrentPage >= totalPages;

    // Render Rows
    if (pageRows.length === 0) {
      this.modalTableBody.innerHTML = `
        <tr>
          <td colspan="${displayHeaders.length + 1}" style="text-align:center; padding: 32px; color:var(--text-muted);">
            검색 결과가 없습니다.
          </td>
        </tr>
      `;
      return;
    }

    this.modalTableBody.innerHTML = pageRows.map((row, idx) => {
      const globalIdx = startIndex + idx + 1;
      const cells = displayHeaders.map(h => {
        const isSource = h === this.core.options.sourceColumnName || h === this.core.options.sheetColumnName;
        const val = row[h] !== undefined && row[h] !== null ? String(row[h]) : '';
        return `<td class="${isSource ? 'col-source-cell' : ''}" title="${val}">${val}</td>`;
      }).join('');

      return `
        <tr>
          <td class="row-index-cell">${globalIdx}</td>
          ${cells}
        </tr>
      `;
    }).join('');
  }

  /**
   * 토스트 알림 표시
   */
  showToast(message, type = 'info') {
    if (!this.toastContainer) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let iconHtml = '';
    if (type === 'success') {
      iconHtml = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#10b981" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>`;
    } else if (type === 'error') {
      iconHtml = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#ef4444" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>`;
    } else {
      iconHtml = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>`;
    }

    toast.innerHTML = `
      <div class="toast-icon">${iconHtml}</div>
      <div class="toast-message">${message}</div>
    `;

    this.toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(40px)';
      setTimeout(() => toast.remove(), 300);
    }, 3500);
  }
}

window.UIController = UIController;
