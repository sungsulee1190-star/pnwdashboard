// 경쟁사 분기·연도 실적 (공시 기준). 2026-10-29 확인.
// 이 세션에서 공시 원문(DART·cninfo·JPX·IDX 등) 접근이 막혀, 공시를 인용한 언론·IR 보도의 검색 발췌로 확인했다. 원문 대조 필요.
// p는 달력 기준 분기(25.4Q = 2025년 10~12월). 회계연도가 다른 회사는 fiscal_label에 자사 분기 표기.
// totalOp는 op_label 기준(중국 3사는 영업이익 대신 지배주주순이익). printRev/printOp는 인쇄용지·그래픽 부문(공시된 경우만).
// '계산'이라고 적힌 분기는 공시된 누계에서 앞 누계를 뺀 값.
// 3분기 미발표 회사는 발표 후 갱신 (주간 예약 작업이 확인).
const DATA_COMP_FIN = {
 "hansol": {
  "unit": "십억원",
  "basis": "연결",
  "op_label": "영업이익",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": 562.1,
    "totalOp": 13.9,
    "printRev": null,
    "printOp": null,
    "status": "잠정공시",
    "date": "2026-01-23",
    "source": "아시아경제",
    "url": "https://view.asiae.co.kr/article/2026012311450829829",
    "note": "전년 동기 적자에서 흑자 전환(2024년 환경사업 대손상각 기저효과)"
   },
   {
    "p": "26.1Q",
    "totalRev": 559.8,
    "totalOp": 11.2,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-05-14",
    "source": "헤럴드경제",
    "url": "https://m.news.nate.com/view/20260514n33807",
    "note": "펄프가 상승, 인쇄·특수지 판가 3.9% 하락. 담합 과징금 1,425.8억은 리니언시로 면제"
   },
   {
    "p": "26.2Q",
    "totalRev": 592.7,
    "totalOp": 50.8,
    "printRev": 167.4,
    "printOp": null,
    "status": "잠정공시",
    "date": "2026-07-30",
    "source": "뉴스핌",
    "url": "https://www.newspim.com/news/view/20260730001306",
    "note": "미국 상호관세 무효화에 따른 관세환급 일회성 포함. 인쇄용지 매출 1,674억(언론 1곳)"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-11-16",
    "source": "",
    "url": "",
    "note": "작년 3분기 잠정실적은 10/30 공시. 분기보고서 기한 11/16"
   }
  ],
  "yearly": [
   {
    "p": "2025년",
    "totalRev": 2290.0,
    "totalOp": 49.9,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "뉴스핌",
    "url": "https://www.newspim.com/news/view/20260123000779",
    "note": "연결. 별도 기준 보도(2조1,691억·441억)와 차이, 사업보고서 확인 필요"
   }
  ],
  "checked": "2026-10-29"
 },
 "moorim": {
  "unit": "십억원",
  "basis": "무림페이퍼 연결(무림P&P 포함)",
  "op_label": "영업이익",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-02-20",
    "source": "디지털투데이",
    "url": "https://www.digitaltoday.co.kr/news/articleView.html?idxno=632871",
    "note": "분기 단독 수치 보도 없음(연간만 공시). 무림P&P 회수보일러 교체로 가동 중단"
   },
   {
    "p": "26.1Q",
    "totalRev": 307.2,
    "totalOp": 0.16,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-05-15",
    "source": "이데일리",
    "url": "https://news.nate.com/view/20260515n29667",
    "note": "펄프 등 원부자재 상승, 자회사 감가상각 부담. 순손실 78억"
   },
   {
    "p": "26.2Q",
    "totalRev": 337.7,
    "totalOp": 28.0,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-08-14",
    "source": "이데일리",
    "url": "https://edaily.co.kr/News/Read?mediaCodeNo=257&newsId=06350086645547320",
    "note": "담합 과징금 861억을 영업외로 반영해 순손실 582억. 정정공시로 과징금 감액, 환입 가능성"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-11-16",
    "source": "",
    "url": "",
    "note": "작년 3분기는 11/14경 공시"
   }
  ],
  "yearly": [
   {
    "p": "2025년",
    "totalRev": 1266.9,
    "totalOp": 5.95,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "디지털투데이",
    "url": "https://www.digitaltoday.co.kr/news/articleView.html?idxno=632871",
    "note": "연결 영업이익 -93.3%, 순손실 241억. 무림P&P 별도: 매출 7,316억, 영업손실 245억"
   }
  ],
  "checked": "2026-10-29"
 },
 "hankuk": {
  "unit": "십억원",
  "basis": "연결 (상장코드 027970)",
  "op_label": "영업이익",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": 181.0,
    "totalOp": -3.26,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-02",
    "source": "씽크풀(연간 공시에서 계산)",
    "url": "https://m.thinkpool.com/stockDiscuss/027970/cont/11864173",
    "note": "유형자산 손상차손 417억, 중대재해·통상임금 비용"
   },
   {
    "p": "26.1Q",
    "totalRev": 199.0,
    "totalOp": 4.6,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-05-14",
    "source": "이데일리",
    "url": "https://edaily.co.kr/News/Read?mediaCodeNo=257&newsId=06448486645448920",
    "note": "담합 과징금 490.6억을 영업외로 반영해 순손실 476억(6회 분납)"
   },
   {
    "p": "26.2Q",
    "totalRev": 220.4,
    "totalOp": 4.7,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-08-14",
    "source": "뉴스1",
    "url": "https://m.news.nate.com/view/20260814n20576",
    "note": "북미 수출 호조로 인쇄용지 회복"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-11-16",
    "source": "",
    "url": "",
    "note": "1·2분기 모두 기한일(5/14, 8/14)에 공시"
   }
  ],
  "yearly": [
   {
    "p": "2025년",
    "totalRev": 753.7,
    "totalOp": 3.1,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "뉴스1",
    "url": "https://www.news1.kr/industry/sb-founded/6136369",
    "note": "연결 영업이익 -83.8%, 순손실 약 335억"
   }
  ],
  "checked": "2026-10-29"
 },
 "sunpaper": {
  "unit": "억 위안",
  "basis": "연결",
  "op_label": "지배주주순이익",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": 102.56,
    "totalOp": 7.51,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-04-10",
    "source": "증권지성",
    "url": "https://stock.stockstar.com/RB2026041200000907.shtml",
    "note": "문화용지 가격 약세"
   },
   {
    "p": "26.1Q",
    "totalRev": 100.83,
    "totalOp": 9.06,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-04-29",
    "source": "신랑재경",
    "url": "https://finance.sina.com.cn/stock/aigc/stockfs/2026-04-29/doc-inhwefuz0932201.shtml",
    "note": "문화용지 가격 저점권"
   },
   {
    "p": "26.2Q",
    "totalRev": 107.79,
    "totalOp": 10.45,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-08-25",
    "source": "격륭회",
    "url": "https://m.gelonghui.com/news/5305854",
    "note": "상반기 문화용지 매출 54.2억(백상지 40.0, 코트지 14.2·매출 -29.7%). 포장지가 견인"
   },
   {
    "p": "26.3Q",
    "totalRev": 111.02,
    "totalOp": 9.71,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-10-28",
    "source": "제몐",
    "url": "https://www.jiemian.com/article/15257523.html",
    "note": "매출 +13.0%, 순이익 +34.9%. 문화용지 가격 하락을 포장지가 상쇄"
   }
  ],
  "yearly": [
   {
    "p": "2025년",
    "totalRev": 391.92,
    "totalOp": 32.51,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "매일경제신문",
    "url": "https://www.nbd.com.cn/articles/2026-04-11/4335140.html",
    "note": "매출 -3.8%, 순이익 +4.8%"
   }
  ],
  "checked": "2026-10-29"
 },
 "chenming": {
  "unit": "억 위안",
  "basis": "연결",
  "op_label": "지배주주순이익",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": 28.39,
    "totalOp": -22.88,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-03-31",
    "source": "증권시보(연간-누적 계산)",
    "url": "https://www.stcn.com/article/detail/3716991.html",
    "note": "재가동 시작으로 매출 회복"
   },
   {
    "p": "26.1Q",
    "totalRev": 29.81,
    "totalOp": -5.12,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-04-29",
    "source": "매일경제신문",
    "url": "https://www.nbd.com.cn/articles/2026-04-29/4374772.html",
    "note": "3/13 5대 기지 전면 재가동"
   },
   {
    "p": "26.2Q",
    "totalRev": 38.84,
    "totalOp": -2.74,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-08-28",
    "source": "21세기경제보도(상반기-1분기 계산)",
    "url": "https://m.21jingji.com/article/20260829/herald/588103c38903be09fcaca3e6cfacdd16.html",
    "note": "손실 축소, 본업은 적자 지속"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-10-31",
    "source": "",
    "url": "",
    "note": "3분기보고서 마감 10/31"
   }
  ],
  "yearly": [
   {
    "p": "2025년",
    "totalRev": 61.87,
    "totalOp": -82.96,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "제몐",
    "url": "https://www.jiemian.com/article/14112006.html",
    "note": "생산 중단 여파로 대규모 적자"
   }
  ],
  "checked": "2026-10-29"
 },
 "huatai": {
  "unit": "억 위안",
  "basis": "연결",
  "op_label": "지배주주순이익",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": 29.75,
    "totalOp": -1.39,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-04-18",
    "source": "동방재부(연간-누적 계산)",
    "url": "https://finance.eastmoney.com/a/202510303550329188.html",
    "note": "4분기 적자로 연간 적자"
   },
   {
    "p": "26.1Q",
    "totalRev": 31.24,
    "totalOp": 0.23,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-04-30",
    "source": "중재망",
    "url": "https://quote.cfi.cn/cwfxzb/1124/600308.html",
    "note": "흑자 복귀"
   },
   {
    "p": "26.2Q",
    "totalRev": 30.87,
    "totalOp": -0.04,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-08-22",
    "source": "중국증권보(상반기-1분기 계산)",
    "url": "https://epaper.cs.com.cn/zgzqb/images/2026-08/22/B060/zqB06022.pdf",
    "note": "상반기 순이익 -72%"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-10-31",
    "source": "",
    "url": "",
    "note": "3분기보고서 마감 10/31"
   }
  ],
  "yearly": [
   {
    "p": "2025년",
    "totalRev": 124.56,
    "totalOp": -0.66,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "매일경제신문",
    "url": "https://www.nbd.com.cn/articles/2026-04-18/4345678.html",
    "note": "매출 -4.5%, 무배당"
   }
  ],
  "checked": "2026-10-29"
 },
 "app": {
  "unit": "백만 달러",
  "basis": "Indah Kiat Tbk(INKP) 기준",
  "fy_note": "APP 중국 법인은 비상장으로 공시 없음",
  "op_label": "영업이익",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": null,
    "totalOp": 202.91,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-03-22",
    "source": "BCA Sekuritas(연간-누적 계산)",
    "url": "https://bcasekuritas.co.id/latest-news/news/penjualan-neto-inkp-capai-usd236-miliar-hingga-september-2025",
    "note": "분기 매출은 반올림치만 있어 비움"
   },
   {
    "p": "26.1Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-04-30",
    "source": "IndoPremier",
    "url": "https://www.indopremier.com/ipotnews/newsDetail.php?jdl=Financial+Statements+1Q+2026+of+INKP&news_id=488846&group_news=RESEARCHNEWS&news_date=&taging_subtype=&name=&search=&q=&halaman=",
    "note": "달러 수치 없음. 루피아 기준 매출 13.75조, 순이익 2.63조(+11%)"
   },
   {
    "p": "26.2Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-07-31",
    "source": "Databoks",
    "url": "https://databoks.katadata.co.id/en/market/statistics/6aa2076adc1e2/indah-kiat-inkp-net-profit-soars-127-in-the-first-half-of-2026",
    "note": "분기 단독 없음. 상반기 누적 매출 1,700, 순이익 372(+127%, 환차익 포함). Tjiwi Kimia 상반기 매출 528.6"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-10-31",
    "source": "",
    "url": "",
    "note": "IDX 마감 10월 말"
   }
  ],
  "yearly": [
   {
    "p": "2025년",
    "totalRev": 3170,
    "totalOp": 667.71,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "IDN Financials",
    "url": "https://www.idnfinancials.com/news/62396/inkp-aim-to-post-us453-mn-in-2025-profit-amid-revenue-pressure",
    "note": "INKP 순이익 453. Tjiwi Kimia: 매출 984.7, 영업이익 66.2, 순이익 275.8"
   }
  ],
  "checked": "2026-10-29"
 },
 "sappi": {
  "unit": "백만 달러",
  "basis": "그룹",
  "fy_note": "회계연도 9월 결산",
  "op_label": "조정 EBITDA",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": null,
    "totalOp": 90,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-02-04",
    "source": "JSE SENS",
    "url": "https://senspdf.jse.co.za/documents/2026/JSE/ISSE/SAVVI/SAPQ1DEC25.pdf",
    "note": "가격 약세·랜드 강세, 순손실 37",
    "fiscal_label": "FY26 1Q"
   },
   {
    "p": "26.1Q",
    "totalRev": 1334,
    "totalOp": 52,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-05-07",
    "source": "PaperAge",
    "url": "https://www.paperage.com/2026news/05-07-2026sappi-reports-results-for-its-financial-quarter-ended-31-march-2026.html",
    "note": "전 지역 판가 하락, 손상차손으로 순손실 413",
    "fiscal_label": "FY26 2Q"
   },
   {
    "p": "26.2Q",
    "totalRev": 1334,
    "totalOp": 53,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-08-06",
    "source": "BusinessDay",
    "url": "https://www.businessday.co.za/companies/earnings/2026-08-06-stronger-rand-and-scheduled-ngodwana-mill-maintenance-weigh-on-sappi/",
    "note": "Ngodwana 정비, 비용 인플레, 순손실 181",
    "fiscal_label": "FY26 3Q"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-11-06",
    "source": "Sappi IR",
    "url": "https://www.sappi.com/en-gb/investors/next-event",
    "note": "4분기·연간 실적 11/6 발표 예정",
    "fiscal_label": "FY26 4Q"
   }
  ],
  "yearly": [
   {
    "p": "FY2025",
    "totalRev": 5420,
    "totalOp": 501,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "MarketScreener",
    "url": "https://www.marketscreener.com/news/sappi-limited-reports-earnings-results-for-the-full-year-ended-september-30-2025-ce7d5cd3d081f227",
    "note": "2024.10~2025.9, 순손실 177"
   }
  ],
  "checked": "2026-10-29"
 },
 "upm": {
  "unit": "백만 유로",
  "basis": "그룹 (부문: Communication Papers)",
  "fy_note": "2026년 2분기부터 합판 사업 제외(계속사업)",
  "op_label": "비교 EBIT",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": 2312,
    "totalOp": 355,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-02-04",
    "source": "UPM 연간 실적",
    "url": "https://www.bolsamania.com/nota-de-prensa/mercados/upm-financial-statements-release-2025-positive-finish-to-the-year---strong-cash-flow-and-decisive-strategic-actions--21606377.html",
    "note": "Kaukas·Ettringen 지류 생산 중단으로 CP 생산능력 13% 감축"
   },
   {
    "p": "26.1Q",
    "totalRev": 2505,
    "totalOp": 274,
    "printRev": null,
    "printOp": 13,
    "status": "공시",
    "date": "2026-04-29",
    "source": "UPM 1분기 보고서",
    "url": "https://news.cision.com/upm-kymmene-oyj/r/upm-interim-report-q1-2026--good-start-to-the-year---excellent-performance-in-decarbonization-soluti,c4341929",
    "note": "CP 비교 EBIT 13. 유럽 그래픽용지 수요 -4%"
   },
   {
    "p": "26.2Q",
    "totalRev": 2355,
    "totalOp": 212,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-07-23",
    "source": "UPM 반기 보고서",
    "url": "https://finance.yahoo.com/markets/stocks/articles/upm-half-financial-report-2026-065600772.html",
    "note": "Sappi와 그래픽용지 합작 본계약. CP EBIT은 출처 충돌로 비움"
   },
   {
    "p": "26.3Q",
    "totalRev": 2202,
    "totalOp": 131,
    "printRev": null,
    "printOp": -8,
    "status": "공시",
    "date": "2026-10-28",
    "source": "UPM 3분기 보고서",
    "url": "https://www.prnewswire.com/news-releases/upm-interim-report-q3-2026-strong-cash-flow-in-a-challenging-quarter-amid-destocking-and-cost-inflation-302921893.html",
    "note": "CP 비교 EBIT -8로 적자 전환(에너지 비용). 유럽 그래픽용지 수요 -7%"
   }
  ],
  "yearly": [
   {
    "p": "2025년",
    "totalRev": 9656,
    "totalOp": 921,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "PaperAge",
    "url": "https://www.paperage.com/2026news/02-04-2026upm-reports-fourth-quarter-and-full-year-2025-results.html",
    "note": "CP 출하 -11%(289만 톤)"
   }
  ],
  "checked": "2026-10-29"
 },
 "arctic": {
  "unit": "백만 즐로티",
  "basis": "그룹 (부문: 제지, 펄프 Rottneros 제외)",
  "op_label": "EBIT",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": 745.1,
    "totalOp": -72.2,
    "printRev": 517.9,
    "printOp": null,
    "status": "공시",
    "date": "2026-02-19",
    "source": "Cision(잠정)",
    "url": "https://news.cision.com/arctic-paper/r/positioned-for-recovery-through-strategic-investments,c4309747",
    "note": "수요 부진·가격 경쟁, 펄프 부진"
   },
   {
    "p": "26.1Q",
    "totalRev": 814.0,
    "totalOp": -36.2,
    "printRev": 573.4,
    "printOp": -2.0,
    "status": "공시",
    "date": "2026-05-12",
    "source": "1분기 보고서",
    "url": "https://mfn.se/cis/a/arctic-paper/arctic-paper-focus-on-margin-recovery-and-cost-discipline-871eab88",
    "note": "4월 5~7%, 6월 4~8% 판가 인상 발표"
   },
   {
    "p": "26.2Q",
    "totalRev": 750.8,
    "totalOp": -29.7,
    "printRev": null,
    "printOp": -31.7,
    "status": "공시",
    "date": "2026-08-11",
    "source": "Quartr(반기 보고서)",
    "url": "https://quartr.com/events/arctic-paper-arp-q2-2026_3gqGigKF",
    "note": "제지 부문 적자 확대"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-11-05",
    "source": "Bankier",
    "url": "https://www.bankier.pl/wiadomosc/ARCTIC-PAPER-S-A-Terminy-publikacji-raportow-okresowych-w-2026-roku-9069488.html",
    "note": "3분기 보고서 11/5 예정"
   }
  ],
  "yearly": [
   {
    "p": "2025년",
    "totalRev": 3197.6,
    "totalOp": -166.0,
    "printRev": 2224.4,
    "printOp": 25.1,
    "status": "공시",
    "source": "Cision(잠정)",
    "url": "https://news.cision.com/arctic-paper/r/positioned-for-recovery-through-strategic-investments,c4309747",
    "note": "제지 부문 매출 2,224.4, EBIT 25.1. 순손실 175"
   }
  ],
  "checked": "2026-10-29"
 },
 "burgo": {
  "unit": "백만 유로",
  "basis": "그룹",
  "fy_note": "비상장, 연간 실적만 공개",
  "op_label": "조정 EBITDA",
  "quarterly": [],
  "yearly": [
   {
    "p": "2025년",
    "totalRev": 1395,
    "totalOp": 100,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "Burgo 지속가능성보고서 2025",
    "url": "https://www.burgo.com/sites/default/files/bilanci/Bilancio_Sostenibilita_2025.pdf",
    "note": "약 수치(circa). EBIT 46, 순이익 33. 판가·에너지·원료비 부담"
   }
  ],
  "checked": "2026-10-29"
 },
 "oji": {
  "unit": "억 엔",
  "basis": "연결 (부문: 인쇄정보미디어)",
  "fy_note": "회계연도 3월 결산",
  "op_label": "영업이익",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": 4779.7,
    "totalOp": 99.4,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-02-06",
    "source": "가부탄",
    "url": "https://s.kabutan.jp/news/k202602060045/",
    "note": "해외 펄프 시황 악화, 내수 물량 감소",
    "fiscal_label": "FY2025 3Q"
   },
   {
    "p": "26.1Q",
    "totalRev": 4687.6,
    "totalOp": 79.1,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-05-15",
    "source": "결산단신(연간-누적 계산)",
    "url": "https://www2.jpx.co.jp/disc/38610/081220260508519342_tse-acedjpsm-38610-20260515338610-ixbrl.htm",
    "note": "구조개혁 추진",
    "fiscal_label": "FY2025 4Q"
   },
   {
    "p": "26.2Q",
    "totalRev": 4685,
    "totalOp": 39,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-08-05",
    "source": "Monex",
    "url": "https://scouter.monex.co.jp/news/detail/3861_20260805_act_20260805_130011_3",
    "note": "상반기 정기수리 집중",
    "fiscal_label": "FY2026 1Q"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-11-06",
    "source": "",
    "url": "",
    "note": "중간결산 11/6 예정(제3자 일정)",
    "fiscal_label": "FY2026 2Q"
   }
  ],
  "yearly": [
   {
    "p": "FY2025",
    "totalRev": 18617.1,
    "totalOp": 345.8,
    "printRev": 2132,
    "printOp": 75.1,
    "status": "공시",
    "source": "결산단신",
    "url": "https://www2.jpx.co.jp/disc/38610/081220260508519342_tse-acedjpsm-38610-20260515338610-ixbrl.htm",
    "note": "2025.4~2026.3. 영업이익 -48.9%. 인쇄정보미디어 수치는 2차 출처 1곳"
   }
  ],
  "checked": "2026-10-29"
 },
 "nippon": {
  "unit": "억 엔",
  "basis": "연결 (부문: 종이·판지)",
  "fy_note": "회계연도 3월 결산",
  "op_label": "영업이익",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": 3003.1,
    "totalOp": 60.3,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-02-06",
    "source": "가부탄",
    "url": "https://s.kabutan.jp/stocks/3863/news/k202602060115/",
    "note": "수출 시황 악화로 종이·판지 매출 감소",
    "fiscal_label": "FY2025 3Q"
   },
   {
    "p": "26.1Q",
    "totalRev": 3030.8,
    "totalOp": 101.7,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-05-15",
    "source": "결산단신(연간-누적 계산)",
    "url": "https://www2.jpx.co.jp/disc/38630/081220260514535670_tse-acedjpsm-38630-20260515338630-ixbrl.htm",
    "note": "배당 10→15엔",
    "fiscal_label": "FY2025 4Q"
   },
   {
    "p": "26.2Q",
    "totalRev": 3144.7,
    "totalOp": 29.1,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-08-05",
    "source": "결산단신",
    "url": "https://www2.jpx.co.jp/disc/38630/081220260804507863_tse-qcedjpsm-38630-20260805338630-ixbrl.htm",
    "note": "영업이익 -47%, 연간 전망 미정",
    "fiscal_label": "FY2026 1Q"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-11-06",
    "source": "",
    "url": "",
    "note": "중간결산 11/6 예정(제3자 일정)",
    "fiscal_label": "FY2026 2Q"
   }
  ],
  "yearly": [
   {
    "p": "FY2025",
    "totalRev": 11926.1,
    "totalOp": 252.1,
    "printRev": 5579,
    "printOp": null,
    "status": "공시",
    "source": "결산단신",
    "url": "https://www2.jpx.co.jp/disc/38630/081220260514535670_tse-acedjpsm-38630-20260515338630-ixbrl.htm",
    "note": "영업이익 +27.9%. 종이·판지 매출은 2차 출처"
   }
  ],
  "checked": "2026-10-29"
 },
 "hokuetsu": {
  "unit": "억 엔",
  "basis": "연결",
  "fy_note": "회계연도 3월 결산",
  "op_label": "영업이익",
  "quarterly": [
   {
    "p": "25.4Q",
    "totalRev": 729.2,
    "totalOp": 33.4,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-02-13",
    "source": "민카부",
    "url": "https://s.minkabu.jp/stock/3865/settlement",
    "note": "원연료가 상승",
    "fiscal_label": "FY2025 3Q"
   },
   {
    "p": "26.1Q",
    "totalRev": 750.3,
    "totalOp": 21.8,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-05-15",
    "source": "결산단신(연간-누적 계산)",
    "url": "https://www2.jpx.co.jp/disc/38650/081220260515536184_tse-acedjpsm-38650-20260515338650-ixbrl.htm",
    "note": "연간 영업이익 -61.8%",
    "fiscal_label": "FY2025 4Q"
   },
   {
    "p": "26.2Q",
    "totalRev": 725.6,
    "totalOp": 6.9,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "date": "2026-08-07",
    "source": "결산단신",
    "url": "https://www2.jpx.co.jp/disc/38650/081220260806511464_tse-qcedjpsm-38650-20260807338650-ixbrl.htm",
    "note": "원연료가 급등, 영업이익 -78.9%",
    "fiscal_label": "FY2026 1Q"
   },
   {
    "p": "26.3Q",
    "totalRev": null,
    "totalOp": null,
    "printRev": null,
    "printOp": null,
    "status": "미발표",
    "date": "2026-11-14",
    "source": "",
    "url": "",
    "note": "작년 중간결산 11/14, 확정일 미확인",
    "fiscal_label": "FY2026 2Q"
   }
  ],
  "yearly": [
   {
    "p": "FY2025",
    "totalRev": 2877.4,
    "totalOp": 75.4,
    "printRev": null,
    "printOp": null,
    "status": "공시",
    "source": "결산단신",
    "url": "https://www2.jpx.co.jp/disc/38650/081220260515536184_tse-acedjpsm-38650-20260515338650-ixbrl.htm",
    "note": "FY2026 전망: 매출 3,050, 영업이익 30"
   }
  ],
  "checked": "2026-10-29"
 }
};
