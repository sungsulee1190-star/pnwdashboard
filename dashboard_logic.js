/**
 * PNW Dashboard - Top-Down Analysis Logic
 * 
 * Includes functions for data aggregation, Top-Down filtering,
 * and Email Draft generation.
 */

// 1. Data Utility Functions
function getLatestPeriod(dataArray) {
    if (!dataArray || dataArray.length === 0) return { year: 2026, month: 1 };
    let latest = dataArray[0];
    for (let i = 1; i < dataArray.length; i++) {
        if (dataArray[i].y > latest.y || (dataArray[i].y === latest.y && dataArray[i].m > latest.m)) {
            latest = dataArray[i];
        }
    }
    return { year: latest.y, month: latest.m };
}

function filterDataByPeriod(dataArray, year, monthFrom, monthTo) {
    return dataArray.filter(d => d.y === year && d.m >= monthFrom && d.m <= monthTo);
}

function calculateYoY(currentData, previousData) {
    let currSum = currentData.reduce((sum, item) => sum + (item.v || 0), 0);
    let prevSum = previousData.reduce((sum, item) => sum + (item.v || 0), 0);
    let diff = currSum - prevSum;
    let rate = prevSum === 0 ? 0 : (diff / prevSum) * 100;
    return { current: currSum, previous: prevSum, diff: diff, rate: rate };
}

function calculateMoM(currentData, previousMonthData) {
    let currSum = currentData.reduce((sum, item) => sum + (item.v || 0), 0);
    let prevSum = previousMonthData.reduce((sum, item) => sum + (item.v || 0), 0);
    let diff = currSum - prevSum;
    let rate = prevSum === 0 ? 0 : (diff / prevSum) * 100;
    return { current: currSum, previous: prevSum, diff: diff, rate: rate };
}

function aggregateByRegion(dataArray) {
    let map = {};
    dataArray.forEach(d => {
        let r = d._rg || "기타";
        if (r.includes("_")) r = r.split("_")[1]; // Simplify region name if it has prefixes
        if (!map[r]) map[r] = { region: r, volume: 0, count: 0 };
        map[r].volume += (d.v || 0);
        map[r].count += 1;
    });
    return Object.values(map).sort((a, b) => b.volume - a.volume);
}

function aggregateByCountry(dataArray, targetRegion = "all") {
    let map = {};
    dataArray.forEach(d => {
        let r = d._rg || "기타";
        if (r.includes("_")) r = r.split("_")[1];
        if (targetRegion !== "all" && r !== targetRegion) return;
        
        let c = d.c || "Unknown";
        if (!map[c]) map[c] = { country: c, region: r, volume: 0 };
        map[c].volume += (d.v || 0);
    });
    return Object.values(map).sort((a, b) => b.volume - a.volume);
}

function calculateContribution(currentRegions, prevRegions, totalDiff) {
    if (totalDiff === 0) return currentRegions.map(r => ({ ...r, contribution: 0 }));
    let prevMap = {};
    prevRegions.forEach(p => prevMap[p.region] = p.volume);
    
    return currentRegions.map(r => {
        let prevVol = prevMap[r.region] || 0;
        let diff = r.volume - prevVol;
        let contribution = (diff / Math.abs(totalDiff)) * 100;
        return { ...r, prevVolume: prevVol, diff: diff, contribution: contribution };
    }).sort((a, b) => b.diff - a.diff);
}

// 2. Email Drafting Functions
function buildNarrative(regionData, countryData, metaData) {
    let text = "";
    if (regionData && regionData.length > 0) {
        let topGrowth = regionData.filter(r => r.diff > 0).slice(0, 2);
        if (topGrowth.length > 0) {
            text += `${topGrowth.map(r => r.region).join(", ")} 지역이 성장을 견인했습니다. `;
        }
        let topDecline = regionData.filter(r => r.diff < 0).reverse().slice(0, 2);
        if (topDecline.length > 0) {
            text += `반면 ${topDecline.map(r => r.region).join(", ")} 지역은 감소폭이 컸습니다. `;
        }
    }
    return text;
}

function buildEmailDraft(options) {
    let { 
        year, monthFrom, monthTo, 
        targetRegion, 
        includeSupply, includePrice, includeKR, includeCN, includeNews 
    } = options;
    
    // Retrieve filtered data
    let currExport = filterDataByPeriod(typeof DATA_EXPORT !== 'undefined' ? DATA_EXPORT : [], year, monthFrom, monthTo);
    let prevExport = filterDataByPeriod(typeof DATA_EXPORT !== 'undefined' ? DATA_EXPORT : [], year - 1, monthFrom, monthTo);
    
    let cnYoY = calculateYoY(currExport, prevExport);
    
    let draftHtml = `<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">`;
    draftHtml += `<h2 style="color: #004488;">${year}년 ${monthFrom === monthTo ? monthFrom + '월' : monthFrom + '~' + monthTo + '월'} 중국/한국 인쇄용지 수출입 동향 업데이트</h2>`;
    
    if (includeNews && typeof DATA_NEWS !== 'undefined') {
        draftHtml += `<h3 style="color: #0066cc; border-bottom: 1px solid #eee; padding-bottom: 5px;">1. 중국 시장 월보 요약</h3>`;
        draftHtml += `<p style="white-space: pre-wrap;">${DATA_NEWS[0]?.summary || '데이터 없음'}</p>`;
    }
    
    if (includeCN) {
        draftHtml += `<h3 style="color: #0066cc; border-bottom: 1px solid #eee; padding-bottom: 5px;">2. 중국산 수출 지역/국가별 변화</h3>`;
        draftHtml += `<p>중국산 수출량은 <b>${cnYoY.current.toLocaleString()} MT</b>로 전년 동기 대비 <b>${cnYoY.diff > 0 ? '증가' : '감소'}(${cnYoY.rate.toFixed(1)}%)</b> 하였습니다.</p>`;
        
        let currRegions = aggregateByRegion(currExport);
        let prevRegions = aggregateByRegion(prevExport);
        let contributions = calculateContribution(currRegions, prevRegions, cnYoY.diff);
        
        draftHtml += `<p>${buildNarrative(contributions)}</p>`;
        
        // Add a small table for top 5 regions
        draftHtml += `<table border="1" style="border-collapse: collapse; width: 100%; max-width: 500px; text-align: right;">`;
        draftHtml += `<tr style="background-color: #f5f5f5;"><th style="text-align:center;">지역</th><th>수출량(MT)</th><th>YoY 증감율</th><th>증감량</th></tr>`;
        contributions.slice(0, 5).forEach(r => {
            let rateStr = r.prevVolume === 0 ? '-' : ((r.diff / r.prevVolume) * 100).toFixed(1) + '%';
            let color = r.diff > 0 ? 'red' : (r.diff < 0 ? 'blue' : 'black');
            draftHtml += `<tr><td style="text-align:center;">${r.region}</td><td>${r.volume.toLocaleString()}</td><td style="color:${color};">${rateStr}</td><td style="color:${color};">${r.diff.toLocaleString()}</td></tr>`;
        });
        draftHtml += `</table>`;
    }
    
    if (includeKR) {
        draftHtml += `<h3 style="color: #0066cc; border-bottom: 1px solid #eee; padding-bottom: 5px;">3. 한국산 수출 동향</h3>`;
        draftHtml += `<p>세부 데이터는 대시보드 한국산 탭을 참조 바랍니다.</p>`;
    }
    
    if (includePrice) {
        draftHtml += `<h3 style="color: #0066cc; border-bottom: 1px solid #eee; padding-bottom: 5px;">4. 가격 동향</h3>`;
        draftHtml += `<p>중국 내수가 및 수출 가격 동향은 전반적으로 변동성이 유지되고 있습니다.</p>`;
    }
    
    draftHtml += `</div>`;
    
    // Create plain text version
    let draftText = draftHtml.replace(/<br>/g, "\n").replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ");
    
    return { html: draftHtml, text: draftText };
}

// Ensure functions are available globally
window.DashboardLogic = {
    getLatestPeriod,
    filterDataByPeriod,
    calculateYoY,
    calculateMoM,
    aggregateByRegion,
    aggregateByCountry,
    calculateContribution,
    buildNarrative,
    buildEmailDraft
};
