// 공장별 CAPA (공시·회사 자료 기준). 키는 data.js DATA_CAPA의 mill 이름.
// 2026-10-29 조사. 이 세션에서는 DART·KIND·cninfo·회사 사이트 원문 접근이 막혀,
// 모든 수치는 검색 결과에 노출된 원문 발췌로 확인했다. 사용 전 url의 원문으로 재확인 필요.
//
// capa: 연간 톤. null = 공장 단위 수치를 찾지 못함
// group: 공장별로 나뉘지 않고 여러 공장 합계로만 공개된 경우 (같은 group의 공장끼리 한 번만 합산)
// basis: 공시 | 공시 인용 | 회사 자료 | 업계 추정 | 미확인
const CAPA_DISCLOSED_META = { checked: '2026-10-29', verified_against_original: false };

const CAPA_DISCLOSED_GROUPS = {
  hansol_print: { capa: 500000, label: '장항+신탄진 합산', basis: '공시 인용', scope: '인쇄용지',
    source: 'KRX 체크포인트 기업분석 보고서 (한솔제지)', period: '2023.06', url: 'https://ssl.pstatic.net/imgstock/upload/research/company/1687851822764.pdf',
    note: '사업보고서는 장항·천안·신탄진을 인쇄용지·특수지 한 줄로만 공개(공장별 미공개). 특수지 34만 톤은 별도.' },
  moorim_print: { capa: 840000, label: '진주+울산 합산', basis: '공시 인용', scope: '인쇄용지 (무림페이퍼+무림P&P)',
    source: '한국신용평가 무림페이퍼 신용평가 보고서', period: '2025년 말', url: '',
    note: '무림P&P 울산 단독은 2023 반기보고서에 FY2022 458,094톤(가동시간 기준 생산능력, 실제 생산 498,470톤). 진주 단독 공시는 찾지 못함.' }
};

const CAPA_DISCLOSED = {
  // 한국
  'Janghang Mill': { capa: null, group: 'hansol_print', basis: '공시 인용' },
  'Shintanjin Mill': { capa: null, group: 'hansol_print', basis: '공시 인용' },
  'Jinju Mill': { capa: null, group: 'moorim_print', basis: '공시 인용' },
  'Ulsan Mill': { capa: null, group: 'moorim_print', basis: '공시 인용' },
  'Onsan Mill': { capa: 600000, basis: '공시', scope: '인쇄용지 (도공지 등)',
    source: '한국제지 제40기 사업보고서', period: '2023.12.31', url: 'https://kind.krx.co.kr/external/2024/03/14/001859/20240314004246/11011.htm',
    note: '회사 홈페이지도 연 60만 톤(PM 4대, 코터 2대). 실제 생산 2025년 398,578톤. 기존 23만 톤은 1995년 수치로 보임.' },

  // 유럽
  'Gratkorn Mill': { capa: 860000, basis: '회사 자료', scope: '그래픽용지 (라벨지 약 9만 톤 별도)',
    source: 'Sappi Gratkorn 공장 페이지', period: '2026', url: 'https://www.sappi.com/it-br/about-us/locations/mills/gratkorn-mill',
    note: 'Sappi·UPM 그래픽용지 합작에 포함(2026년 말 종결 예정).' },
  'Ehingen Mill': { capa: 280000, basis: '회사 자료', scope: '무목질 코트지',
    source: 'Sappi Ehingen 공장 페이지', period: '2026', url: 'https://www.sappi.com/en-gb/about-us/locations/mills/ehingen-mill', note: '' },
  'Maastricht Mill': { capa: 260000, basis: '공시', scope: '코트지+판지 합계 (그래픽 단독 미확인)',
    source: 'Sappi 2025 통합보고서, 사업장 표 (70/190/260천 톤)', period: 'FY2025', url: 'https://sappi-ir-reports.co.za/online-reports/air-2025/where-we-operate.php',
    note: '2017~18년 PM6 개조 후 대부분 판지 생산. 그래픽용지는 7만 톤일 가능성.' },
  'Kymi Paper Mill': { capa: 715000, basis: '회사 자료', scope: '코트·언코트 상질지·특수지 (PM8, PM9)',
    source: 'UPM Paper, UPM Kymi 공장 페이지', period: '2026', url: 'https://www.upmpaper.com/about-us/paper-mills-offices/upm-kymi/', note: '펄프 87만 톤 별도.' },
  'Nordland Mill': { capa: 843000, basis: '공시 인용', scope: '공장 전체 (상질지·특수지)',
    source: 'UPM EMAS Report 2025, Nordland', period: '2025', url: 'https://www.upm.com/siteassets/reporting-hub/emas-reports/emas-report-2025-upm-nordland-paper-mill-english.pdf',
    note: 'PM3(28만 톤) 2024년 말 폐쇄 반영. 합작에는 일부 라인만 포함.' },
  'Changshu Mill': { capa: 1220000, basis: '공시 인용', scope: '공장 전체 (상질지·라벨·글라신)',
    source: 'UPM Changshu EMAS Report 2021', period: '2021', url: 'https://www.upm.com/siteassets/documents/responsibility/1-fundamentals/emas-reports/upm-pulp-and-paper-mills-report/english/changshu_emas_2021_en.pdf',
    note: '기존 UPM 합계 1,220,000은 이 공장 수치와 같음.' },
  'Grycksbo Mill': { capa: 220000, basis: '공시', scope: '무목질 코트지',
    source: 'Arctic Paper Grycksbo EMAS 2024, 2025.03 보도자료', period: '2024~2025', url: 'https://www.arcticpaper.com/globalassets/arcticpaper.com/09-services/certificates/certificate/emas-certificate/emas2024_apg_eng-online_082025.pdf', note: '' },
  'Virton Mill': { capa: 390000, basis: '회사 자료', scope: '무목질 코트지 (PM1)',
    source: 'Burgo Virton 공장 페이지', period: '현재', url: 'https://www.burgo.com/en/paper-mills/virton', note: '회사 뉴스는 38만 톤 이상으로 표기.' },

  // 중국·인도네시아 (APP 계열은 비상장으로 법정 공시 없음)
  'Zhenjiang (Gold East Paper) Mill': { capa: 2000000, basis: '업계 추정', scope: '코트지, 공장 전체',
    source: 'Voith twogether 2021 외', period: '2015~2021', url: 'https://www.voith.com/corp-en/voith-paper_twogether21_en.pdf', note: 'PM1 50만 + PM2 50만 + PM3 110만 톤.' },
  'Yangpu (Hainan Jinhai Pulp & Paper Industry) Mill': { capa: 900000, basis: '공시 인용', scope: '문화용지 (140~300g 고평량)',
    source: '중국신문망 하이난 2021', period: '2021', url: 'http://www.hi.chinanews.com.cn/hnnew/2021-09-07/4_142718.html', note: '150만 톤 개조 계획(2023 착공), 완공 미확인. 화장지 70만 톤 별도.' },
  'Karawang No. 1 (PT Pindo Deli Pulp & Paper Mills) Mill': { capa: 1500000, basis: '회사 자료', scope: '전 지종, 카라왕 3개 공장 합계',
    source: 'Pindo Deli 회사 소개', period: '현재', url: 'https://pindodeli.co.id/about-us', note: '코트지 단독 수치 없음.' },
  'Shouguang (Shouguang Chenming Coated Woodfree Paper No. 3) Mill': { capa: 800000, basis: '공시 인용', scope: '코트지 1개 라인',
    source: '중국제지협회 방문 보고 외', period: '2020~2022', url: 'https://www.hstong.com/news/detail/25120415555592702', note: '2024.11 가동 중단 후 2025.03 재가동, 2026.03 전 기지 정상화 공시.' },
  'Shouguang (Shouguang Meilun Paper) Mill': { capa: 510000, basis: '공시 인용', scope: '고급 문화용지 (코트지 여부 미확인)',
    source: '중국제지협회 방문 보고 외', period: '2018~2022', url: 'https://www.hstong.com/news/detail/25120415555592702', note: '' },
  'Dongying Mill': { capa: 700000, basis: '업계 추정', scope: '고급 코트지',
    source: '중국증권보 2017', period: '2017', url: 'https://cs.com.cn/ssgs/ggjjd/20150607_65428/201703/t20170331_5227145.html', note: '연차보고서에 지종별 생산능력 없음.' },
  'Shandong Sun Paper Industry (Yanzhou) Mill': { capa: 800000, basis: '업계 추정', scope: '회사 전체 코트지 (옌저우가 주력)',
    source: '증권사 리포트 2020~2021', period: '2020~2021', url: 'https://img3.gelonghui.com/pdf/47f04-a051375f-2304-40ff-bb19-d237320b0e12.pdf', note: '' },
  'Nantong Mill': { capa: 400000, basis: '공시', scope: '고급 코트지 1개 라인',
    source: '장쑤왕자제지 환경영향평가서', period: '2021경', url: 'https://www.nexi.go.jp/environment/info/pdf/21-023_01.pdf', note: '두 번째 40만 톤 라인은 건설 취소. BHKP 47만 톤 별도.' },

  // 일본 (유가증권보고서·통합보고서에 공장별 생산능력 미공개)
  'Kasugai Mill': { capa: null, basis: '미확인', note: '2018년 생산 약 66만 톤(인쇄·정보용지 63%), 생산능력 미공개.' },
  'Yonago Mill': { capa: null, basis: '미확인', note: '고급 도공지 전용 공장. 일산 1,560톤(구 자료)만 확인.' },
  'Iwakuni Mill': { capa: null, basis: '미확인', note: '중기계획 2030에서 존속하는 그래픽용지 3개 거점 중 하나.' },
  'Ishinomaki Mill': { capa: null, basis: '미확인', note: 'N6(27.1만 톤) 2022년 중단. 존속 거점.' },
  'Hokkaido Mill - Shiraoi': { capa: null, basis: '미확인', note: 'PM8(6.4만 톤) 2025년 중단, FY2030/3까지 그래픽용지 철수 예정.' },
  'Harada (Nippon Paper Papylia) Mill': { capa: null, basis: '미확인', note: '특수지 공장.' },
  'Niigata Mill': { capa: 350000, basis: '회사 자료', scope: 'N9 온머신 코트지 PM만 (공장 전체 아님)',
    source: '호쿠에쓰 통합보고서 2025', period: '2025', url: 'https://hokuetsucorp.com/pdf/cr/cr2025_jp.pdf', note: '공장 전체 7개 PM, 전 지종 약 100만 톤. 1대는 골판지 원지로 전환.' }
};
