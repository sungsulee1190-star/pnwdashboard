// 경쟁사 공시·뉴스 (주차별 헤드라인용). 최신 항목을 위에 추가.
// date: 기사·공시 게재일 (가격 인상은 적용일, summary에 명시)
// cat: 한국 | 중국 | 유럽 | 일본 | 시장
const DATA_COMP_NEWS = [
  {
    date: "2026-10-28", cat: "중국", comp: "Sun Paper (태양지업)",
    title: "태양지업 3분기 순이익 9.71억 위안, 전년 대비 34.9% 증가",
    summary: "3분기 매출 111.02억 위안(+13.0%), 지배주주 순이익 9.71억 위안(+34.9%). 1~3분기 누적 순이익은 29.22억 위안(+16.9%). 실적은 포장지 사업이 끌었고 문화용지(코트지·백상지)는 약세가 이어졌다는 분석. 펄프·제지 일관 설비 증설도 계속.",
    source: "제몐(界面新闻), 동방재부", url: "https://www.jiemian.com/article/15257523.html"
  },
  {
    date: "2026-10-28", cat: "유럽", comp: "UPM",
    title: "UPM 3분기 비교 EBIT 1.31억 유로(-8%), 고객 재고조정과 원가 상승",
    summary: "3분기 매출 22.02억 유로(전년 22.26억), 비교 EBIT 1.31억 유로로 8% 감소. 고객사 재고조정에 따른 출하 감소와 변동비 상승이 원인. 영업현금흐름은 4.20억 유로로 개선. Sappi와 추진 중인 그래픽용지 합작법인은 기업결합 승인 절차 진행 중.",
    source: "UPM 보도자료 (PR Newswire)", url: "https://www.prnewswire.com/news-releases/upm-interim-report-q3-2026-strong-cash-flow-in-a-challenging-quarter-amid-destocking-and-cost-inflation-302921893.html"
  },
  {
    date: "2026-10-21", cat: "한국", comp: "한솔제지 (Hansol)",
    title: "한솔제지 대표, 국정감사에서 인쇄용지 담합 사과",
    summary: "한경록 대표가 국회 정무위 국감에서 담합에 사과. 리니언시로 과징금(약 1,425억 원) 전액 면제를 받았지만 피해 보상 방안은 제시하지 못했다는 지적. 출판·인쇄업계는 12월 손해배상 청구를 예고.",
    source: "뉴스핌, 딜사이트", url: "https://www.newspim.com/news/view/20261022000936"
  },
  {
    date: "2026-10-18", cat: "한국", comp: "무림 / 한국제지",
    title: "무림·한국제지, 실적 급증에도 과징금 반영 시 적자 전환 분석",
    summary: "공정위는 6개 제지사가 2021년 2월~2024년 12월 7차례 가격 인상을 합의했다며 총 3,383억 원 과징금 부과. 일부 제지사는 법원에서 집행정지 결정을 받음. 영업 실적은 늘었지만 과징금을 반영하면 적자로 돌아선다는 분석.",
    source: "한국경제", url: "https://www.hankyung.com/article/2026101887821"
  },
  {
    date: "2026-10-12", cat: "유럽", comp: "Sappi / Burgo / Cartiere del Garda",
    title: "유럽 코트지 가격 인상 이어져, Sappi 매엽 코트지 5~7% (10/12 적용)",
    summary: "Sappi Europe은 무목질 코트지 매엽(sheets)을 10월 12일 신규 주문·출하분부터 5~7% 인상. Cartiere del Garda는 양면 코트지를 10월 5일 출하분부터 톤당 50유로 인상(에너지·물류·원료비). Burgo는 9월 그래픽·특수지 7~8% 인상을 발표. 날짜는 적용일 기준.",
    source: "Sappi 보도자료 외", url: "https://www.sappi.com/en-gb/about-us/news-and-events/sappi-europe-announces-price-increase-woodfree-coated-sheets"
  },
  {
    date: "2026-10-12", cat: "한국", comp: "제지업계 전반",
    title: "출판계, 인쇄용지 담합 제지사 상대 공동 손해배상 소송 추진",
    summary: "약 4년간 담합으로 인쇄용지 가격이 71% 올랐다며 출판계가 공동 소송에 나섬. 일부 제지사의 집행정지로 가격 정상화가 늦어지고 있다는 주장.",
    source: "아시아경제", url: "https://view.asiae.co.kr/article/2026101214244935617"
  }
];
