(function () {
  const databaseName = 'ledger-loan-manager';
  const storeName = 'file-handles';
  const handleKey = 'excel-workbook';
  const encoder = new TextEncoder();
  const decoder = new TextDecoder();
  const spreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const loanHeaders = ['Loan ID', 'Borrower', 'Loan type', 'Principal (INR)', 'Annual interest rate (%)', 'Term (months)', 'Monthly EMI (INR)', 'Disbursement date', 'EMIs paid', 'Outstanding principal (INR)', 'Status', 'Notes'];
  const paymentHeaders = ['Payment ID', 'Loan ID', 'Borrower', 'Installment', 'Amount (INR)', 'Payment date', 'Reference'];
  const crcTable = Array.from({ length: 256 }, (_, value) => {
    let crc = value;
    for (let bit = 0; bit < 8; bit += 1) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
    return crc >>> 0;
  });

  function xmlEscape(value) {
    return String(value ?? '').replace(/[&<>"']/g, (character) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;'
    })[character]);
  }

  function columnName(index) {
    let name = '';
    let number = index + 1;
    while (number > 0) {
      number -= 1;
      name = String.fromCharCode(65 + number % 26) + name;
      number = Math.floor(number / 26);
    }
    return name;
  }

  function worksheetXml(headers, records) {
    const rows = [headers, ...records];
    const body = rows.map((row, rowIndex) => {
      const cells = row.map((value, columnIndex) => {
        const reference = `${columnName(columnIndex)}${rowIndex + 1}`;
        if (typeof value === 'number' && Number.isFinite(value)) {
          return `<c r="${reference}"${rowIndex === 0 ? ' s="1"' : ''}><v>${value}</v></c>`;
        }
        return `<c r="${reference}" t="inlineStr"${rowIndex === 0 ? ' s="1"' : ''}><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
      }).join('');
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    }).join('');
    const lastColumn = columnName(headers.length - 1);
    const lastRow = Math.max(1, rows.length);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="19"/><sheetData>${body}</sheetData><autoFilter ref="A1:${lastColumn}${lastRow}"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
  }

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }

  function concatenate(parts) {
    const size = parts.reduce((total, part) => total + part.length, 0);
    const output = new Uint8Array(size);
    let offset = 0;
    for (const part of parts) {
      output.set(part, offset);
      offset += part.length;
    }
    return output;
  }

  function zip(files) {
    const localParts = [];
    const centralParts = [];
    let localOffset = 0;
    for (const [path, content] of files) {
      const name = encoder.encode(path);
      const bytes = encoder.encode(content);
      const checksum = crc32(bytes);
      const localHeader = new Uint8Array(30);
      const localView = new DataView(localHeader.buffer);
      localView.setUint32(0, 0x04034b50, true);
      localView.setUint16(4, 20, true);
      localView.setUint16(10, 0, true);
      localView.setUint16(12, 33, true);
      localView.setUint32(14, checksum, true);
      localView.setUint32(18, bytes.length, true);
      localView.setUint32(22, bytes.length, true);
      localView.setUint16(26, name.length, true);
      localParts.push(localHeader, name, bytes);

      const centralHeader = new Uint8Array(46);
      const centralView = new DataView(centralHeader.buffer);
      centralView.setUint32(0, 0x02014b50, true);
      centralView.setUint16(4, 20, true);
      centralView.setUint16(6, 20, true);
      centralView.setUint16(12, 0, true);
      centralView.setUint16(14, 33, true);
      centralView.setUint32(16, checksum, true);
      centralView.setUint32(20, bytes.length, true);
      centralView.setUint32(24, bytes.length, true);
      centralView.setUint16(28, name.length, true);
      centralView.setUint32(42, localOffset, true);
      centralParts.push(centralHeader, name);
      localOffset += localHeader.length + name.length + bytes.length;
    }

    const centralDirectory = concatenate(centralParts);
    const end = new Uint8Array(22);
    const endView = new DataView(end.buffer);
    endView.setUint32(0, 0x06054b50, true);
    endView.setUint16(8, files.length, true);
    endView.setUint16(10, files.length, true);
    endView.setUint32(12, centralDirectory.length, true);
    endView.setUint32(16, localOffset, true);
    return concatenate([...localParts, centralDirectory, end]);
  }

  function buildWorkbook({ loans, payments }) {
    const contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
    const packageRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    const workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Loans" sheetId="1" r:id="rId1"/><sheet name="Repayments" sheetId="2" r:id="rId2"/></sheets></workbook>';
    const workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF23755B"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    const files = [
      ['[Content_Types].xml', contentTypes],
      ['_rels/.rels', packageRels],
      ['xl/workbook.xml', workbook],
      ['xl/_rels/workbook.xml.rels', workbookRels],
      ['xl/styles.xml', styles],
      ['xl/worksheets/sheet1.xml', worksheetXml(loanHeaders, loans)],
      ['xl/worksheets/sheet2.xml', worksheetXml(paymentHeaders, payments)]
    ];
    return zip(files);
  }

  function columnIndex(reference) {
    const letters = String(reference).match(/^[A-Z]+/)?.[0];
    if (!letters) throw new Error('Workbook contains an invalid cell reference.');
    return [...letters].reduce((index, letter) => index * 26 + letter.charCodeAt(0) - 64, 0) - 1;
  }

  function parseWorksheet(xml, expectedHeaders) {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    if (document.querySelector('parsererror')) throw new Error('Workbook contains invalid worksheet data.');
    const rows = [...document.getElementsByTagNameNS(spreadsheetNamespace, 'row')].map((row) => {
      const values = [];
      for (const cell of row.getElementsByTagNameNS(spreadsheetNamespace, 'c')) {
        const index = columnIndex(cell.getAttribute('r'));
        const inlineText = cell.getElementsByTagNameNS(spreadsheetNamespace, 't');
        const value = inlineText.length
          ? [...inlineText].map((text) => text.textContent).join('')
          : cell.getElementsByTagNameNS(spreadsheetNamespace, 'v')[0]?.textContent || '';
        values[index] = value;
      }
      return values;
    });
    if (!rows.length || expectedHeaders.some((header, index) => rows[0][index] !== header)) return null;
    return rows.slice(1).filter((row) => row.some((value) => value !== undefined && value !== '')).map((row) => expectedHeaders.map((_, index) => row[index] ?? ''));
  }

  async function readWorkbook(handle) {
    const file = await handle.getFile();
    if (!file.size) return null;
    const bytes = new Uint8Array(await file.arrayBuffer());
    const view = new DataView(bytes.buffer);
    const files = new Map();
    let offset = 0;
    while (offset + 30 <= bytes.length && view.getUint32(offset, true) === 0x04034b50) {
      const method = view.getUint16(offset + 8, true);
      const flags = view.getUint16(offset + 6, true);
      const expectedCrc = view.getUint32(offset + 14, true);
      const size = view.getUint32(offset + 22, true);
      const nameSize = view.getUint16(offset + 26, true);
      const extraSize = view.getUint16(offset + 28, true);
      if (method !== 0 || flags & 0x08) return null;
      const nameStart = offset + 30;
      const contentStart = nameStart + nameSize + extraSize;
      const name = decoder.decode(bytes.slice(nameStart, nameStart + nameSize));
      const content = bytes.slice(contentStart, contentStart + size);
      if (content.length !== size || crc32(content) !== expectedCrc) throw new Error('Workbook archive is incomplete or damaged.');
      files.set(name, decoder.decode(content));
      offset = contentStart + size;
    }

    const loansXml = files.get('xl/worksheets/sheet1.xml');
    const paymentsXml = files.get('xl/worksheets/sheet2.xml');
    if (!loansXml || !paymentsXml) return null;
    const loanRows = parseWorksheet(loansXml, loanHeaders);
    const paymentRows = parseWorksheet(paymentsXml, paymentHeaders);
    if (!loanRows || !paymentRows) return null;
    const loans = loanRows.map((row) => ({
      id: row[0],
      borrower: row[1],
      type: row[2],
      principal: Number(row[3]),
      rate: Number(row[4]),
      term: Number(row[5]),
      emi: Number(row[6]),
      startDate: row[7],
      notes: row[11],
      createdAt: new Date().toISOString()
    }));
    if (loans.some((loan) => !loan.id || !loan.borrower || !Number.isFinite(loan.principal) || !Number.isFinite(loan.rate) || !Number.isInteger(loan.term) || !loan.startDate)) {
      throw new Error('Workbook has incomplete loan records.');
    }
    const payments = paymentRows.map((row) => ({
      id: row[0],
      loanId: row[1],
      installment: Number(row[3]),
      amount: Number(row[4]),
      paidDate: row[5],
      note: row[6]
    }));
    if (payments.some((payment) => !payment.id || !payment.loanId || !Number.isInteger(payment.installment) || !Number.isFinite(payment.amount) || !payment.paidDate)) {
      throw new Error('Workbook has incomplete repayment records.');
    }
    return { loans, payments };
  }

  function openDatabase() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(databaseName, 1);
      request.onupgradeneeded = () => request.result.createObjectStore(storeName);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function saveHandle(handle) {
    const database = await openDatabase();
    await new Promise((resolve, reject) => {
      const transaction = database.transaction(storeName, 'readwrite');
      transaction.objectStore(storeName).put(handle, handleKey);
      transaction.oncomplete = resolve;
      transaction.onerror = () => reject(transaction.error);
    });
    database.close();
  }

  async function loadHandle() {
    const database = await openDatabase();
    const handle = await new Promise((resolve, reject) => {
      const request = database.transaction(storeName, 'readonly').objectStore(storeName).get(handleKey);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
    database.close();
    return handle;
  }

  async function sync(handle, records, requestPermission = false) {
    let permission = await handle.queryPermission({ mode: 'readwrite' });
    if (permission !== 'granted' && requestPermission) {
      permission = await handle.requestPermission({ mode: 'readwrite' });
    }
    if (permission !== 'granted') throw new Error('Reconnect the Excel workbook to allow automatic updates.');
    const bytes = buildWorkbook(records);
    const writable = await handle.createWritable();
    try {
      await writable.write(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      await writable.close();
    } catch (error) {
      await writable.abort();
      throw error;
    }
  }

  window.excelStore = { buildWorkbook, loadHandle, readWorkbook, saveHandle, sync };
})();
