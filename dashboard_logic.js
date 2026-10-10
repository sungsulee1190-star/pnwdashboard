/** PNW Dashboard aggregation and evidence-based email drafts. */
const DASHBOARD_REGION_MAP = {
    '01_북미':'북미','02_오세아':'오세아니아','03_서유럽_영국':'유럽','04_서유럽':'유럽','05_동유럽':'유럽','05_유럽_기타':'유럽',
    '06_아프리카_남아공':'아프리카','07_아프리카_기타':'아프리카','08_남미_멕시코':'중남미','09_남미_기타':'중남미',
    '10_중국':'동북아','11_몽골':'동북아','12_대만':'동북아','13_동남아_말싱':'동남아','14_동남아_인니':'동남아',
    '15_동남아_베트남':'동남아','16_동남아-태국':'동남아','17_동남아_필리핀':'동남아','18_동남아_기타':'동남아',
    '19_인도':'인도/서남아','21_방글라':'인도/서남아','21_서남아_2':'인도/서남아','22_중동(기타)':'중동','23_중동(이란)':'중동','24_일본':'동북아','25_대한민국':'한국',
    '01_미국':'북미','02_호주':'오세아니아','03_서유럽':'유럽','04_동유럽':'유럽','05_아프리카':'아프리카','06_남미':'중남미',
    '07_로컬':'한국','07_일본':'동북아','08_중국':'동북아','09_홍콩':'동북아','09_동남아':'동남아','10_일본':'동북아','10_서남아':'인도/서남아',
    '11_중동(기타)':'중동','12_이란':'중동','12_중동(이란)':'중동','13_인도':'인도/서남아','14_방글라데시':'인도/서남아','15_태국':'동남아',
    '16_말레이시아':'동남아','17_싱가포르':'동남아','18_인도네시아':'동남아','19_베트남':'동남아','20_필리핀':'동남아','21_동남아(기타)':'동남아','22_서남아(기타)':'인도/서남아'
};
const DASHBOARD_REGIONS = new Set(Object.values(DASHBOARD_REGION_MAP));
function cleanRegion(region) {
    if (typeof region !== 'string') return '기타';
    const key = region.trim();
    return DASHBOARD_REGIONS.has(key) ? key : (Object.hasOwn(DASHBOARD_REGION_MAP, key) ? DASHBOARD_REGION_MAP[key] : '기타');
}
function validPeriod(year, from, to) {
    return Number.isInteger(year) && year >= 1 && year <= 9999 && Number.isInteger(from) && Number.isInteger(to) && from >= 1 && to <= 12 && from <= to;
}
function getLatestPeriod(dataArray) {
    const rows = (Array.isArray(dataArray) ? dataArray : []).filter(d => d && validPeriod(d.y, d.m, d.m));
    if (!rows.length) return null;
    const latest = rows.reduce((a,b) => a.y * 12 + a.m >= b.y * 12 + b.m ? a : b);
    return { year: latest.y, month: latest.m };
}
function filterDataByPeriod(dataArray, year, monthFrom, monthTo) {
    if (!validPeriod(year, monthFrom, monthTo)) throw new RangeError('기간 오류: 연도와 월 범위를 확인하세요.');
    return (Array.isArray(dataArray) ? dataArray : []).filter(d => d && d.y === year && d.m >= monthFrom && d.m <= monthTo);
}
function volumeSum(data) {
    const rows = (Array.isArray(data) ? data : []).filter(d => d && Number.isFinite(d.v));
    return rows.length ? rows.reduce((sum,d) => sum + d.v, 0) : null;
}
function calculateYoY(currentData, previousData) {
    const current = volumeSum(currentData), previous = volumeSum(previousData);
    const diff = current === null || previous === null ? null : current - previous;
    return { current, previous, diff, rate: diff === null || previous === 0 ? null : diff / previous * 100 };
}
function calculateMoM(currentData, previousMonthData) { return calculateYoY(currentData, previousMonthData); }
function rowRegion(row) { return cleanRegion(row._rg || row.rg); }
function aggregateByRegion(dataArray) {
    const map = new Map();
    (Array.isArray(dataArray) ? dataArray : []).forEach(d => {
        if (!d || !Number.isFinite(d.v)) return;
        const region = rowRegion(d);
        if (!map.has(region)) map.set(region, { region, volume: 0, count: 0 });
        map.get(region).volume += d.v;
        map.get(region).count++;
    });
    return [...map.values()].sort((a,b) => b.volume - a.volume);
}
function aggregateByCountry(dataArray, targetRegion = 'all') {
    const map = new Map();
    (Array.isArray(dataArray) ? dataArray : []).forEach(d => {
        if (!d || !Number.isFinite(d.v)) return;
        const region = rowRegion(d);
        if (targetRegion !== 'all' && region !== targetRegion) return;
        const country = d.c || 'Unknown';
        if (!map.has(country)) map.set(country, { country, region, volume: 0 });
        map.get(country).volume += d.v;
    });
    return [...map.values()].sort((a,b) => b.volume - a.volume);
}
function calculateContribution(currentRegions, prevRegions, totalDiff) {
    const current = new Map((currentRegions || []).map(r => [r.region,r]));
    const previous = new Map((prevRegions || []).map(r => [r.region,r]));
    return [...new Set([...current.keys(),...previous.keys()])].map(region => {
        const r = current.get(region) || { region, volume: null, count: 0 };
        const prevVolume = Number.isFinite(previous.get(region)?.volume) ? previous.get(region).volume : null;
        const diff = Number.isFinite(r.volume) && prevVolume !== null ? r.volume - prevVolume : null;
        const contribution = diff === null || !Number.isFinite(totalDiff) ? null : totalDiff !== 0 ? diff / Math.abs(totalDiff) * 100 : 0;
        return { ...r, prevVolume, diff, contribution };
    }).sort((a,b) => a.diff === null ? (b.diff === null ? 0 : 1) : b.diff === null ? -1 : b.diff - a.diff);
}
function buildNarrative(regionData) {
    const rows = regionData || [];
    const growth = rows.filter(r => r.diff > 0).sort((a,b) => b.diff - a.diff).slice(0,2);
    const decline = rows.filter(r => r.diff < 0).sort((a,b) => a.diff - b.diff).slice(0,2);
    return [growth.length ? growth.map(r => r.region).join(', ') + ' 지역의 증가량이 컸습니다.' : '',
        decline.length ? decline.map(r => r.region).join(', ') + ' 지역의 감소량이 컸습니다.' : ''].filter(Boolean).join(' ');
}
function escapeDraftText(value) {
    return String(value ?? '').replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' })[c]);
}
function draftPlainText(html) {
    return html.replace(/<br\s*\/?\s*>/gi,'\n').replace(/<\/(?:h[1-6]|p|div|tr|table)>/gi,'\n')
        .replace(/<\/(?:td|th)>/gi,'\t').replace(/<[^>]*>/g,'')
        .replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&nbsp;/g,' ').replace(/&amp;/g,'&')
        .replace(/[ \t]+\n/g,'\n').replace(/\n{3,}/g,'\n\n').trim();
}
function buildEmailDraft(options = {}) {
    const { year, monthFrom, monthTo, includeCN, includeKR, includeNews, includePrice } = options;
    const targetRegion = options.targetRegion || 'all';
    let html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#333">';
    const paragraph = text => { html += '<p style="white-space:pre-wrap">' + escapeDraftText(text) + '</p>'; };
    let section = 0;
    const heading = text => { html += '<h3>' + (++section) + '. ' + escapeDraftText(text) + '</h3>'; };
    const format = value => value.toLocaleString('en-US', { maximumFractionDigits: 2 });
    const table = (headers,rows) => {
        html += '<table border="1" style="border-collapse:collapse;text-align:right"><tr>' + headers.map(h => '<th>' + escapeDraftText(h) + '</th>').join('') + '</tr>';
        rows.forEach(row => { html += '<tr>' + row.map(value => '<td>' + escapeDraftText(value) + '</td>').join('') + '</tr>'; });
        html += '</table>';
    };
    const finish = () => { html += '</div>'; return { html, text: draftPlainText(html) }; };
    if (!validPeriod(year,monthFrom,monthTo)) {
        paragraph('기간 오류: 유효한 연도와 1~12월 범위를 선택하세요. 시작월은 종료월보다 늦을 수 없습니다.');
        return finish();
    }
    const period = year + '년 ' + (monthFrom === monthTo ? monthFrom : monthFrom + '~' + monthTo) + '월';
    html += '<h2>' + escapeDraftText(period) + ' 중국/한국 인쇄용지 수출 동향 업데이트</h2>';
    paragraph('수출 대상지역: ' + (targetRegion === 'all' ? '전체' : targetRegion));
    if (includeNews) {
        heading('중국 시장 월보 요약');
        const news = typeof DATA_NEWS !== 'undefined' ? DATA_NEWS : null;
        const items = Array.isArray(news?.items) ? news.items : [];
        if (!items.length) paragraph('월보 데이터 없음.');
        else {
            paragraph('아래는 저장된 최신 월보 요약입니다. 게시일과 본문 기준월은 수출 선택 기간과 다를 수 있습니다.');
            items.slice(0,3).forEach(item => {
                paragraph((item.title || '월보') + ' | 출처: ' + (item.source || '미상') + ' | 게시일: ' + (item.date || '미상'));
                paragraph(item.summary || '요약 데이터 없음.');
            });
        }
    }
    function exportSection(data,title,source,korean) {
        heading(title);
        paragraph('출처: ' + source + '; 선택 기간 ' + period + '.');
        const rows = (Array.isArray(data) ? data : []).filter(d => d && Number.isFinite(d.v) && (targetRegion === 'all' || rowRegion(d) === targetRegion));
        const current = filterDataByPeriod(rows,year,monthFrom,monthTo);
        const previous = year > 1 ? filterDataByPeriod(rows,year-1,monthFrom,monthTo) : [];
        const missing = selected => Array.from({length:monthTo-monthFrom+1},(_,i) => monthFrom+i).filter(m => !selected.some(d => d.m === m));
        const missingCurrent = missing(current), missingPrevious = missing(previous);
        if (missingCurrent.length) paragraph(missingCurrent.map(m => m + '월').join(', ') + ' 수출 데이터 없음. 표시 수출량은 관측된 월의 합계입니다.');
        if (!current.length) return;
        const change = calculateYoY(current,previous);
        const changes = calculateContribution(aggregateByRegion(current),aggregateByRegion(previous),change.diff).map(r => {
            const missingRegionCurrent = missing(current.filter(d => rowRegion(d) === r.region));
            const missingRegionPrevious = missing(previous.filter(d => rowRegion(d) === r.region));
            const regionalComparable = !missingRegionCurrent.length && !missingRegionPrevious.length;
            return { ...r, missingRegionCurrent, missingRegionPrevious, diff: regionalComparable ? r.diff : null, contribution: regionalComparable ? r.contribution : null };
        });
        paragraph((korean ? '한국산' : '중국산') + ' 수출량: ' + format(change.current) + ' MT.');
        const comparable = !missingCurrent.length && !missingPrevious.length && changes.every(r => r.diff !== null);
        if (!comparable) paragraph('전년 동기 비교 데이터 없음 또는 불완전' + (missingPrevious.length ? ' (' + (year-1) + '년 ' + missingPrevious.map(m => m+'월').join(', ') + ')' : '') + '; 증감률 산출 불가.');
        else if (change.previous === 0) paragraph('전년 동기 기준 수출량 0 MT: 증감량 ' + format(change.diff) + ' MT; 증감률 산출 불가.');
        else paragraph('전년 동기 대비 ' + (change.diff === 0 ? '변동 없음' : change.diff > 0 ? '증가' : '감소') + ' (' + change.rate.toFixed(1) + '%), 증감량 ' + format(change.diff) + ' MT.');
        changes.filter(r => r.diff === null).forEach(r => {
            const gaps = [r.missingRegionCurrent.length ? year + '년 ' + r.missingRegionCurrent.map(m => m+'월').join(', ') : '',
                r.missingRegionPrevious.length ? (year-1) + '년 ' + r.missingRegionPrevious.map(m => m+'월').join(', ') : ''].filter(Boolean).join('; ');
            paragraph(r.region + ': ' + gaps + ' 지역별 관측 데이터 없음. 표시 수출량은 관측된 월의 합계이며 지역별 증감 비교는 산출 불가.');
        });
        paragraph(buildNarrative(changes));
        table(['지역','수출량(MT)','YoY 증감률','증감량(MT)'],changes.slice(0,10).map(r => [r.region,r.volume === null ? '데이터 없음' : format(r.volume),
            r.diff === null ? '비교 데이터 없음' : r.prevVolume === 0 ? '기준 0: 산출 불가' : (r.diff/r.prevVolume*100).toFixed(1)+'%',r.diff === null ? '산출 불가' : format(r.diff)]));
        table(['상위 수출국','수출량(MT)'],aggregateByCountry(current).slice(0,5).map(r => [r.country,format(r.volume)]));
        if (korean) {
            if (!current.every(d => Number.isFinite(d.amt))) paragraph('수출 금액 데이터 없음 또는 불완전; 가중평균 FOB 산출 불가.');
            else {
                const amount = current.reduce((sum,d) => sum+d.amt,0);
                paragraph('수출 금액: ' + format(amount) + ' USD; 가중평균 FOB: ' + (change.current > 0 ? format(amount/change.current) + ' USD/MT' : '수출량 0: 산출 불가') + '. (FOB = 합계 금액 / 합계 수출량)');
            }
            const companies = new Map();
            current.forEach(d => companies.set(d.comp || '미상',(companies.get(d.comp || '미상') || 0)+d.v));
            table(['수출사','수출량(MT)'],[...companies].sort((a,b) => b[1]-a[1]).map(([company,volume]) => [company,format(volume)]));
        }
    }
    if (includeCN) exportSection(typeof DATA_EXPORT !== 'undefined' ? DATA_EXPORT : [],'중국산 수출 지역/국가별 변화','대시보드에 저장된 중국산 국가별 월별 수출실적',false);
    if (includeKR) exportSection(typeof DATA_KR_COUNTRY !== 'undefined' ? DATA_KR_COUNTRY : [],'한국산 수출 동향','대시보드에 저장된 한국산 수출사/국가별 월별 수출실적',true);
    if (includePrice) {
        heading('가격 동향');
        paragraph('출처: 대시보드에 저장된 중국 내수가 (CNY/MT). 수출 대상지역 필터와 별도로 중국 가격 관측지역을 표시합니다.');
        const prices = typeof DATA_PRICE !== 'undefined' && Array.isArray(DATA_PRICE.prices) ? DATA_PRICE.prices.filter(d => d && Number.isFinite(d.p) && validPeriod(d.y,d.m,d.m)) : [];
        const selected = filterDataByPeriod(prices,year,monthFrom,monthTo).sort((a,b) => a.m-b.m);
        if (!selected.length) paragraph('선택 기간 가격 데이터 없음.');
        else {
            const missingMonths = Array.from({length:monthTo-monthFrom+1},(_,i) => monthFrom+i).filter(m => !selected.some(d => d.m === m));
            if (missingMonths.length) paragraph(missingMonths.map(m => m+'월').join(', ') + ' 가격 데이터 없음.');
            table(['월','제품/규격','관측지역','가격','전월 대비'],selected.map(d => {
                const prevYear = d.m === 1 ? d.y-1 : d.y, prevMonth = d.m === 1 ? 12 : d.m-1;
                const prior = prices.find(p => p.y === prevYear && p.m === prevMonth && p.t === d.t && p.g === d.g && p.r === d.r);
                const change = !prior ? '전월 데이터 없음' : format(d.p-prior.p) + ' CNY/MT; ' + (prior.p === 0 ? '기준 0: 증감률 산출 불가' : ((d.p-prior.p)/prior.p*100).toFixed(1)+'%');
                return [d.y+'년 '+d.m+'월',((d.t || '미상')+' '+(d.g || '')).trim(),d.r || '미상',format(d.p)+' CNY/MT',change];
            }));
        }
    }
    return finish();
}
const DashboardLogic = { getLatestPeriod, filterDataByPeriod, calculateYoY, calculateMoM, cleanRegion, cleanRg: cleanRegion, aggregateByRegion, aggregateByCountry, calculateContribution, buildNarrative, buildEmailDraft };
if (typeof window !== 'undefined') window.DashboardLogic = DashboardLogic;
if (typeof module !== 'undefined' && module.exports) module.exports = DashboardLogic;

