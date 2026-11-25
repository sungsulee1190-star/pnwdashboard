// Export the rendered market map without sending any company data to a server.
(() => {
  const encode = text => new TextEncoder().encode(text);
  const xml = value => String(value).replace(/[&<>"']/g, c =>
    ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&apos;'})[c]);
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const packageRel = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml';

  function columnName(index) {
    let name = '';
    for (let n = index + 1; n; n = Math.floor((n - 1) / 26))
      name = String.fromCharCode(65 + (n - 1) % 26) + name;
    return name;
  }

  function worksheet(table) {
    const occupied = new Map(), merges = [];
    const rows = Array.from(table.rows, (row, r) => {
      let col = 0;
      const cells = Array.from(row.cells, cell => {
        while ((occupied.get(col) || 0) > r) col++;
        const address = columnName(col) + (r + 1);
        const end = col + cell.colSpan;
        for (let i = col; i < end; i++) occupied.set(i, r + cell.rowSpan);
        if (cell.colSpan > 1 || cell.rowSpan > 1)
          merges.push(address + ':' + columnName(end - 1) + (r + cell.rowSpan));
        col = end;
        const value = (cell.innerText ?? cell.textContent).trim();
        if (cell.tagName === 'TD' && /^-?(?:0|[1-9]\d{0,2}(?:,\d{3})+|[1-9]\d*)(?:\.\d+)?$/.test(value))
          return '<c r="' + address + '"><v>' + value.replaceAll(',', '') + '</v></c>';
        return '<c r="' + address + '" t="inlineStr"><is><t xml:space="preserve">' + xml(value) + '</t></is></c>';
      }).join('');
      return '<row r="' + (r + 1) + '">' + cells + '</row>';
    }).join('');
    const merged = merges.length ? '<mergeCells count="' + merges.length + '">' +
      merges.map(ref => '<mergeCell ref="' + ref + '"/>').join('') + '</mergeCells>' : '';
    return '<worksheet xmlns="' + ns + '"><sheetData>' + rows + '</sheetData>' + merged + '</worksheet>';
  }

  // OOXML is a ZIP package. Stored entries suffice for the small visible tables.
  function zip(entries) {
    const chunks = [], directory = [];
    let offset = 0, directorySize = 0;
    const header = length => new DataView(new ArrayBuffer(length));
    for (const [path, source] of entries) {
      const name = encode(path), data = encode(source);
      let crc = 0xffffffff;
      for (const byte of data) {
        crc ^= byte;
        for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
      }
      crc = (crc ^ 0xffffffff) >>> 0;
      const local = header(30), central = header(46);
      local.setUint32(0, 0x04034b50, true);
      central.setUint32(0, 0x02014b50, true);
      central.setUint16(4, 20, true);
      for (const [h, start] of [[local,4],[central,6]]) {
        h.setUint16(start,20,true);
        h.setUint16(start+8,33,true); // 1980-01-01
        h.setUint32(start+10,crc,true);
        h.setUint32(start+14,data.length,true);
        h.setUint32(start+18,data.length,true);
        h.setUint16(start+22,name.length,true);
      }
      central.setUint32(42,offset,true);
      chunks.push(local.buffer,name,data);
      directory.push(central.buffer,name);
      offset += 30 + name.length + data.length;
      directorySize += 46 + name.length;
    }
    const end = header(22);
    end.setUint32(0,0x06054b50,true);
    end.setUint16(8,entries.length,true);
    end.setUint16(10,entries.length,true);
    end.setUint32(12,directorySize,true);
    end.setUint32(16,offset,true);
    return new Blob([...chunks,...directory,end.buffer], {type:mime + '.sheet'});
  }

  window.downloadMarketMapWorkbook = (sheets, filename) => {
    if (!sheets.length) throw new Error('내려받을 표가 없습니다.');
    const items = sheets.map(({name, table}, i) => ({name, table, id:i+1}));
    const entries = [
      ['[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="' + mime + '.sheet.main+xml"/>' + items.map(s => '<Override PartName="/xl/worksheets/sheet' + s.id + '.xml" ContentType="' + mime + '.worksheet+xml"/>').join('') + '</Types>'],
      ['_rels/.rels','<Relationships xmlns="' + packageRel + '"><Relationship Id="rId1" Type="' + rel + '/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
      ['xl/workbook.xml','<workbook xmlns="' + ns + '" xmlns:r="' + rel + '"><sheets>' + items.map(s => '<sheet name="' + xml(s.name) + '" sheetId="' + s.id + '" r:id="rId' + s.id + '"/>').join('') + '</sheets></workbook>'],
      ['xl/_rels/workbook.xml.rels','<Relationships xmlns="' + packageRel + '">' + items.map(s => '<Relationship Id="rId' + s.id + '" Type="' + rel + '/worksheet" Target="worksheets/sheet' + s.id + '.xml"/>').join('') + '</Relationships>'],
      ...items.map(s => ['xl/worksheets/sheet' + s.id + '.xml', worksheet(s.table)])
    ];
    const url = URL.createObjectURL(zip(entries));
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  };
})();
