const ExcelJS = require('exceljs');
const path = require('path');

async function createTemplate() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Pranay HOME');

    const COLORS = {
        LABEL_BG: 'FFFCE4D6',
        HEADER_BG: 'FFF4B084',
        RESULT_YELLOW: 'FFFFFF00',
        RESULT_GREEN: 'FFC6E0B4',
        BORDER: 'FF000000'
    };

    const styleCell = (cell, options = {}) => {
        cell.font = { name: 'Arial', size: 10, bold: !!options.bold };
        cell.border = {
            top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        cell.alignment = { vertical: 'middle', horizontal: options.align || 'left' };
        if (options.fill) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: options.fill } };
        }
        if (options.numFmt) cell.numFmt = options.numFmt;
    };

    sheet.getColumn('B').width = 25;
    sheet.getColumn('C').width = 15;
    sheet.getColumn('D').width = 15;
    sheet.getColumn('E').width = 15;
    sheet.getColumn('F').width = 15;
    sheet.getColumn('G').width = 15;
    sheet.getColumn('H').width = 15;
    sheet.getColumn('I').width = 15;
    sheet.getColumn('J').width = 15;

    const labels = ['Consumer Name', 'Consumer No', 'Fixed Charges', 'Sanct. Load (kW)', 'Connection Type'];
    labels.forEach((label, i) => {
        const row = i + 1;
        styleCell(sheet.getCell(`B${row}`), { bold: true, fill: COLORS.LABEL_BG });
        sheet.getCell(`B${row}`).value = label;
        styleCell(sheet.getCell(`D${row}`), { align: 'center' });
        styleCell(sheet.getCell(`H${row}`), { align: 'center' });
    });

    const tableHeaders = ['Sr.No', 'Month', 'Units', 'Bill Amount', 'Unit Cost', 'Month', 'Units', 'Bill Amount', 'Unit Cost'];
    tableHeaders.forEach((header, i) => {
        const cell = sheet.getCell(8, i + 1); // Starting at A? No, subagent said Sr No is in Column A? 
        // Wait, subagent said A is Sr No.
        // Let's check subagent report: "Column A contains the Serial Numbers (Sr.No)."
        // "Units (Critical) D9:D20"
        // "Month Name B9:B20" (Consumer 1)
        // "Month Name G9:G20" (Consumer 2)
        
        // This means:
        // A: Sr No
        // B: Month 1
        // C: ? (Maybe Sr No for consumer 1?)
        // D: Units 1
        // E: Amount 1
        // F: Cost 1
        // G: Month 2
        // H: Units 2
        // I: Amount 2
        // J: Cost 2
    });

    // Let's use the subagent's exact mapping for the table:
    const headerMapping = [
        { col: 'A', val: 'Sr.No' },
        { col: 'B', val: 'Month' },
        { col: 'D', val: 'Units' },
        { col: 'E', val: 'Bill Amount' },
        { col: 'F', val: 'Unit Cost' },
        { col: 'G', val: 'Month' },
        { col: 'H', val: 'Units' },
        { col: 'I', val: 'Bill Amount' },
        { col: 'J', val: 'Unit Cost' }
    ];

    headerMapping.forEach(h => {
        const cell = sheet.getCell(`${h.col}8`);
        cell.value = h.val;
        styleCell(cell, { bold: true, fill: COLORS.HEADER_BG, align: 'center' });
    });

    for (let i = 0; i < 12; i++) {
        const row = 9 + i;
        styleCell(sheet.getCell(`A${row}`), { align: 'center' });
        sheet.getCell(`A${row}`).value = i + 1;
        ['B','D','E','F','G','H','I','J'].forEach(col => styleCell(sheet.getCell(`${col}${row}`), { align: 'center' }));
    }

    const results = [
        { label: 'Average', c1: 'D24', c2: 'H24', f1: 'IFERROR(AVERAGE(D9:D20), 0)', f2: 'IFERROR(AVERAGE(H9:H20), 0)', fmt: '0.00' },
        { label: 'kW', c1: 'D25', c2: 'H25', f1: 'D24/106.06', f2: 'H24/106.06', fmt: '0.00' },
        { label: 'Solar Panels', c1: 'D26', c2: 'H26', f1: 'D25/0.6', f2: 'H25/0.6', fmt: '0.00' },
        { label: 'Solar capacity', c1: 'D27', c2: 'H27', f1: 'ROUND(D26*0.6, 1)', f2: 'ROUND(H26*0.6, 1)', fmt: '0.0', fill: COLORS.RESULT_YELLOW },
        { label: 'Number of Panels', c1: 'D28', c2: 'H28', f1: 'ROUNDUP(D26, 0)', f2: 'ROUNDUP(H26, 0)', fmt: '0', fill: COLORS.RESULT_GREEN }
    ];

    results.forEach((item) => {
        styleCell(sheet.getCell(item.c1), { align: 'right', bold: true, fill: item.fill, numFmt: item.fmt });
        sheet.getCell(item.c1).value = { formula: item.f1 };
        styleCell(sheet.getCell(item.c2), { align: 'right', bold: true, fill: item.fill, numFmt: item.fmt });
        sheet.getCell(item.c2).value = { formula: item.f2 };
        styleCell(sheet.getCell(`B${sheet.getCell(item.c1).row}`), { bold: true, fill: COLORS.LABEL_BG });
        sheet.getCell(`B${sheet.getCell(item.c1).row}`).value = item.label;
    });

    styleCell(sheet.getCell('B29'), { bold: true });
    sheet.getCell('B29').value = 'Total solar capacity';
    styleCell(sheet.getCell('D29'), { align: 'center', bold: true });
    sheet.getCell('D29').value = { formula: 'D27+H27' };

    styleCell(sheet.getCell('B30'), { bold: true });
    sheet.getCell('B30').value = 'Number of solar panels';
    styleCell(sheet.getCell('D30'), { align: 'center', bold: true });
    sheet.getCell('D30').value = { formula: 'D28+H28' };

    await workbook.xlsx.writeFile(path.join(__dirname, 'Solar_Load_Template.xlsx'));
    console.log('Template created with coordinates matching the reference spreadsheet.');
}

createTemplate();
