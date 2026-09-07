/**
 * SampleGenerator - 1-클릭으로 규격이 다른 여러 엑셀 샘플 파일 생성
 */
class SampleGenerator {
  /**
   * 샘플 엑셀 파일 6개(3가지 규격)를 생성하여 File 객체 배열로 반환
   */
  static async generateSampleFiles() {
    const files = [];

    // [규격 1] 회원 명단 데이터 (3개 파일)
    const memberHeaders = ['회원번호', '이름', '연락처', '이메일', '가입일자', '거주지역', '회원등급'];
    
    const members1 = [
      ['M1001', '김민준', '010-1234-5678', 'minjun.kim@example.com', '2024-01-05', '서울 강남구', 'VIP'],
      ['M1002', '이서연', '010-2345-6789', 'sy.lee@example.com', '2024-01-12', '경기 성남시', 'GOLD'],
      ['M1003', '박도윤', '010-3456-7890', 'doyun.park@example.com', '2024-01-19', '인천 연수구', 'SILVER'],
      ['M1004', '최지우', '010-4567-8901', 'jiwoo.choi@example.com', '2024-01-25', '서울 송파구', 'VIP'],
    ];

    const members2 = [
      ['M1005', '정예준', '010-5678-9012', 'yejun.jung@example.com', '2024-02-03', '부산 해운대구', 'GOLD'],
      ['M1006', '강하은', '010-6789-0123', 'haeun.kang@example.com', '2024-02-14', '대구 수성구', 'SILVER'],
      ['M1007', '조시우', '010-7890-1234', 'siwoo.cho@example.com', '2024-02-22', '서울 마포구', 'GOLD'],
    ];

    const members3 = [
      ['M1008', '윤서아', '010-8901-2345', 'seoa.yoon@example.com', '2024-03-08', '대전 유성구', 'VIP'],
      ['M1009', '장서준', '010-9012-3456', 'seojun.jang@example.com', '2024-03-15', '경기 수원시', 'SILVER'],
      ['M1010', '임지아', '010-0123-4567', 'jia.lim@example.com', '2024-03-29', '광주 서구', 'GOLD'],
    ];

    files.push(this.createExcelFile('2024_01월_회원명단.xlsx', '회원목록', [memberHeaders, ...members1]));
    files.push(this.createExcelFile('2024_02월_회원명단.xlsx', '회원목록', [memberHeaders, ...members2]));
    files.push(this.createExcelFile('2024_03월_회원명단.xlsx', '회원목록', [memberHeaders, ...members3]));

    // [규격 2] 주문/매출 거래 데이터 (2개 파일)
    const salesHeaders = ['주문번호', '주문일시', '상품코드', '상품명', '수량', '단가', '결제금액', '결제수단'];

    const sales1 = [
      ['ORD-20240101-01', '2024-01-01 10:23', 'PRD-A101', '인체공학 무선 마우스', 2, 45000, 90000, '신용카드'],
      ['ORD-20240102-02', '2024-01-02 14:15', 'PRD-B202', '기계식 게이밍 키보드', 1, 128000, 128000, '카카오페이'],
      ['ORD-20240105-03', '2024-01-05 16:40', 'PRD-C303', '27인치 4K 모니터', 1, 450000, 450000, '계좌이체'],
      ['ORD-20240108-04', '2024-01-08 19:12', 'PRD-A101', '인체공학 무선 마우스', 3, 45000, 135000, '네이버페이'],
    ];

    const sales2 = [
      ['ORD-20240203-01', '2024-02-03 11:05', 'PRD-D404', 'USB-C 멀티 허브', 2, 38000, 76000, '신용카드'],
      ['ORD-20240210-02', '2024-02-10 13:45', 'PRD-B202', '기계식 게이밍 키보드', 1, 128000, 128000, '신용카드'],
      ['ORD-20240218-03', '2024-02-18 20:30', 'PRD-E505', '노이즈캔슬링 헤드셋', 1, 289000, 289000, '토스페이'],
    ];

    files.push(this.createExcelFile('2024_01월_상품매출내역.xlsx', '매출데이터', [salesHeaders, ...sales1]));
    files.push(this.createExcelFile('2024_02월_상품매출내역.xlsx', '매출데이터', [salesHeaders, ...sales2]));

    // [규격 3] 인사 조직 현황 데이터 (1개 파일)
    const empHeaders = ['사번', '성명', '소속본부', '팀명', '직급', '내선번호', '입사년월'];
    const empData = [
      ['EMP202101', '한동훈', '기술개발본부', '프론트엔드팀', '선임연구원', '02-555-1001', '2021-03'],
      ['EMP202203', '송민경', '기술개발본부', '백엔드팀', '책임연구원', '02-555-1002', '2022-07'],
      ['EMP202305', '백현우', '사업기획본부', '마케팅팀', '주임', '02-555-2001', '2023-01'],
      ['EMP202309', '유지선', '경영지원본부', '인사총무팀', '대리', '02-555-3001', '2023-09'],
    ];

    files.push(this.createExcelFile('2024_상반기_임직원현황.xlsx', '사원명부', [empHeaders, ...empData]));

    return files;
  }

  /**
   * 2차원 배열 데이터로 엑셀 File 인스턴스 생성
   */
  static createExcelFile(fileName, sheetName, dataAoA) {
    const ws = XLSX.utils.aoa_to_sheet(dataAoA);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    
    return new File([excelBuffer], fileName, {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      lastModified: Date.now()
    });
  }
}

window.SampleGenerator = SampleGenerator;
