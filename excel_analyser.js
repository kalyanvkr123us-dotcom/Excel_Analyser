const fs = require('fs');
const path = require('path');
const childProcess = require('child_process');
const XLSX = require('xlsx');

const defaultSourcePath = path.join(__dirname, 'Source Data', 'GoCardless Generator Jan JNL.xlsx');
const inputPath = process.argv[2] || defaultSourcePath;
const outputPath = process.argv[3] || inputPath;
const outputSheetBaseName = 'AI analyser';
const journalSheetBaseName = 'AI Journal';

const notes = [
	['Notes', 'Category', 'Explanation'],
	['[1]', 'Payments', 'Total payments received from branch members.'],
	['[2]', 'Surcharge', '£1 surcharge for each branch based on number of contributions'],
	['[3]', 'Post-Surcharge', 'Amount remaining after deduction of £1 surcharge'],
	['[4]', 'MidTierCharge', '55% charge on the next £1,000'],
	['[5]', 'PostMidTier', 'Amount remaining after deduction of 55% charge'],
	['[6]', 'TopTierCharge', '60% charge on the residual amount (if more than £1k)'],
	['[7]', 'Residual', 'Amount remaining for the branch after deduction of all charges']
];

function text(value) {
	return value === null || value === undefined ? '' : String(value).trim();
}

function branchKey(value) {
	return text(value).split(':', 1)[0].trim();
}

function fallbackBranchKey(value) {
	return branchKey(value).replace(/^(refund for|failure fee for)\s+/i, '').trim();
}

function number(value) {
	if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
	const cleaned = text(value).replace(/[£$,]/g, '');
	if (!cleaned) return 0;
	const parsed = Number(cleaned);
	return Number.isFinite(parsed) ? parsed : 0;
}

function formula(value) {
	return { formula: value };
}

function excelColumnName(index) {
	let name = '';
	let column = index + 1;
	while (column > 0) {
		const remainder = (column - 1) % 26;
		name = String.fromCharCode(65 + remainder) + name;
		column = Math.floor((column - 1) / 26);
	}
	return name;
}

function monthEndSerial(value) {
	const date = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
	const monthEnd = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0));
	return (monthEnd.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
}

function monthNarration(value) {
	const date = new Date(Date.UTC(1899, 11, 30) + value * 86400000);
	const month = date.toLocaleString('en-US', { month: 'short', timeZone: 'UTC' });
	return `Income Allocation ${month}-${String(date.getUTCFullYear()).slice(-2)}`;
}

function formatCsvDate(value) {
	const date = new Date(Date.UTC(1899, 11, 30) + Number(value) * 86400000);
	if (!Number.isFinite(Number(value)) || Number.isNaN(date.getTime())) return value;
	return `${String(date.getUTCDate()).padStart(2, '0')}/${String(date.getUTCMonth() + 1).padStart(2, '0')}/${date.getUTCFullYear()}`;
}

function findColumn(headers, name) {
	const wanted = name.toLowerCase();
	const index = headers.findIndex((header) => text(header).toLowerCase() === wanted);
	if (index < 0) throw new Error(`Column "${name}" was not found.`);
	return index;
}

function rowsFromSheet(workbook, sheetName, headerRow = 0) {
	const sheet = workbook.Sheets[sheetName];
	if (!sheet) throw new Error(`Sheet "${sheetName}" was not found.`);
	return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false })
		.slice(headerRow);
}

function mappingFromMapSheet(workbook) {
	const rows = rowsFromSheet(workbook, 'Map');
	const headerIndex = rows.findIndex((row) => text(row[0]).toLowerCase() === 'resources.description');
	if (headerIndex < 0) throw new Error('The Map tab does not contain a resources.description header.');

	const headers = rows[headerIndex];
	const descriptionColumn = findColumn(headers, 'resources.description');
	const codeColumn = findColumn(headers, 'Code');
	const mapping = new Map();

	rows.slice(headerIndex + 1).forEach((row) => {
		const description = branchKey(row[descriptionColumn]);
		const code = text(row[codeColumn]);
		if (description && code) {
			const isCourierDescription = /^couriers\s+(and|&)\s+logistics/i.test(description);
			mapping.set(description, isCourierDescription ? 'CLB' : code);
		}
	});
	return mapping;
}

function accountMappingFromACMap(workbook) {
	const rows = rowsFromSheet(workbook, 'AC_Map');
	const mapping = new Map();

	rows.forEach((row) => {
		const accountCode = text(row[0]);
		const branchCode = text(row[2]);
		if (accountCode && branchCode) mapping.set(branchCode, accountCode);
	});
	return mapping;
}

function analyse(workbook) {
	const mapping = mappingFromMapSheet(workbook);
	const accountMapping = accountMappingFromACMap(workbook);
	const inputRows = rowsFromSheet(workbook, 'Input');
	if (!inputRows.length) throw new Error('The Input tab is empty.');

	const headers = inputRows[0];
	const descriptionColumn = findColumn(headers, 'resources.description');
	const statusColumn = findColumn(headers, 'resources.status');
	const createdAtColumn = findColumn(headers, 'resources.created_at');
	const grossColumn = findColumn(headers, 'gross_amount');
	const feesColumn = findColumn(headers, 'gocardless_fees');
	const groups = new Map();
	const unmapped = new Map();
	const monthCounts = new Map();
	const descriptionColumnName = excelColumnName(descriptionColumn);
	let goCardlessFees = 0;
	let analysedRowCount = 0;

	inputRows.slice(1).forEach((row, rowIndex) => {
		if (text(row[statusColumn]).toLowerCase() === 'settled') return;
		const description = text(row[descriptionColumn]);
		if (!description) return;
		analysedRowCount += 1;
		const createdAt = number(row[createdAtColumn]);
		const createdAtDate = new Date(Date.UTC(1899, 11, 30) + createdAt * 86400000);
		const monthKey = `${createdAtDate.getUTCFullYear()}-${createdAtDate.getUTCMonth()}`;
		const monthEntry = monthCounts.get(monthKey) || { count: 0, serial: createdAt };
		monthEntry.count += 1;
		monthEntry.serial = Math.max(monthEntry.serial, createdAt);
		monthCounts.set(monthKey, monthEntry);
		goCardlessFees += number(row[feesColumn]);
		const descriptionKey = branchKey(description);
		const code = mapping.get(descriptionKey) || mapping.get(fallbackBranchKey(descriptionKey)) || 'UNMAPPED';
		if (code === 'UNMAPPED') {
			const entry = unmapped.get(description) || { count: 0, inputRow: rowIndex + 2 };
			entry.count += 1;
			unmapped.set(description, entry);
		}
		if (!groups.has(code)) groups.set(code, { payments: 0, contributions: 0 });
		const group = groups.get(code);
		group.payments += number(row[grossColumn]);
		group.contributions += 1;
	});

	const table = [['MAP.Code', 'Payments', 'Surcharge', 'Post-Surcharge', 'MidTierCharge', 'PostMidTier', 'TopTierCharge', 'Residual']];
	const journalRows = [];
	const firstDataRow = 6;
	let outputRow = firstDataRow;
	for (const [code, group] of groups) {
		const row = outputRow++;
		const normalizedCode = text(code).toUpperCase();
		const isGeneral = normalizedCode === 'GENERAL';
		const isIncome = normalizedCode === 'INCOME';
		const isNap = normalizedCode === 'NAP';
		const surcharge = isIncome ? 0 : group.contributions;
		const postSurcharge = group.payments - surcharge;
		const midTierCharge = isGeneral || isIncome ? 0 : Math.min(Math.max(postSurcharge, 0), 1000) * 0.55;
		const postMidTier = Math.max(postSurcharge - midTierCharge, 0);
		const topTierCharge = isGeneral || isIncome ? 0 : Math.max(postSurcharge - 1000, 0) * 0.6;
		const residual = isIncome ? group.payments : Math.max(postMidTier - topTierCharge, 0);
		table.push([
			code,
			group.payments,
			surcharge,
			formula(`=B${row}-C${row}`),
			isGeneral || isIncome ? formula('=0') : formula(`=MIN(MAX(D${row},0),1000)*55%`),
			formula(`=MAX(D${row}-E${row},0)`),
			isGeneral || isIncome ? formula('=0') : formula(`=MAX(D${row}-1000,0)*60%`),
			isIncome ? formula(`=B${row}`) : formula(`=MAX(F${row}-G${row},0)`)
		]);
		journalRows.push({ code, residualAccount: isGeneral || isNap ? '500CU' : accountMapping.get(code) || '', surcharge, midTierCharge, topTierCharge, residual });
	}

	const totalRow = firstDataRow + groups.size;
	const lastDataRow = totalRow - 1;
	const total = [
		'Total',
		...['B', 'C', 'D', 'E', 'F', 'G', 'H'].map((column) => formula(`=SUM(${column}${firstDataRow}:${column}${lastDataRow})`))
	];
	table.push(total);

	const output = [
		['Income Distribution - Calculation'],
		[],
		['', '[1]', '[2]', '[3]', '[4]', '[5]', '[6]', '[7]'],
		['Category', 'Payments', 'Surcharge', 'Post-Surcharge', 'MidTierCharge', 'PostMidTier', 'TopTierCharge', 'Residual'],
		...table,
		[],
		...notes,
		[],
		['Mapping checks'],
		['Description', 'Remittance count'],
		...Array.from(unmapped, ([description, { count, inputRow }]) => [
			description,
			formula(`=HYPERLINK("#'Input'!$${descriptionColumnName}$${inputRow}",${count})`)
		])
	];
	const dominantMonth = Array.from(monthCounts.values()).reduce((current, candidate) => {
		if (!current || candidate.count > current.count || (candidate.count === current.count && candidate.serial > current.serial)) {
			return candidate;
		}
		return current;
	}, null);
	const journalDate = monthEndSerial(dominantMonth ? dominantMonth.serial : 0);
	const narration = monthNarration(journalDate);
	const allocationAmount = journalRows.reduce((sum, row) => sum + row.surcharge + row.midTierCharge + row.topTierCharge + row.residual, 0);
	const journalOutput = [
		['Journal Date', '', { value: journalDate }],
		[],
		['*Amount', 'Description', '*AccountCode', '*Date', '*Narration', '*TaxRate'],
		[goCardlessFees, 'GC Fees', '708CU', { value: journalDate }, narration, 'No VAT'],
		[-goCardlessFees, 'GC Fees', 133, { value: journalDate }, narration, 'No VAT'],
		[allocationAmount, 'Income Allocation', 133, { value: journalDate }, narration, 'No VAT'],
		...journalRows.flatMap(({ code, residualAccount, surcharge, midTierCharge, topTierCharge, residual }) => [
			...(surcharge ? [[-surcharge, `${code} - £1 Surcharge`, '500LD', { value: journalDate }, narration, 'No VAT']] : []),
			...(midTierCharge ? [[-midTierCharge, `${code} - 55% Charge`, '500CU', { value: journalDate }, narration, 'No VAT']] : []),
			...(topTierCharge ? [[-topTierCharge, `${code} - 60% Charge`, '500CU', { value: journalDate }, narration, 'No VAT']] : []),
			...(residual ? [[-residual, `${code} - Residual`, residualAccount, { value: journalDate }, narration, 'No VAT']] : [])
		])
	];
	return {
		output,
		journalOutput,
		unmappedCount: Array.from(unmapped.values()).reduce((sum, entry) => sum + entry.count, 0),
		rowCount: analysedRowCount
	};
}

function nextOutputSheetName(workbook) {
	if (!workbook.SheetNames.includes(outputSheetBaseName)) return outputSheetBaseName;

	let version = 1;
	while (workbook.SheetNames.includes(`${outputSheetBaseName} V${version}`)) version += 1;
	return `${outputSheetBaseName} V${version}`;
}

function versionedSheetName(baseName, analyserSheetName) {
	const version = analyserSheetName.match(/ V(\d+)$/);
	return version ? `${baseName} V${version[1]}` : baseName;
}

function appendSheetWithExcel(data, sheetName, filePrefix) {
	const dataPath = path.join(require('os').tmpdir(), `${filePrefix}-${process.pid}.json`);
	const scriptPath = path.join(__dirname, 'append_ai_analyser.ps1');
	fs.writeFileSync(dataPath, JSON.stringify(data), 'utf8');
	try {
		childProcess.execFileSync('powershell.exe', [
			'-NoProfile',
			'-ExecutionPolicy', 'Bypass',
			'-File', scriptPath,
			'-WorkbookPath', outputPath,
			'-DataPath', dataPath,
			'-SheetName', sheetName
		], { stdio: 'inherit' });
	} finally {
		if (fs.existsSync(dataPath)) fs.unlinkSync(dataPath);
	}
}

function exportJournalCsv(journalData, journalSheetName) {
	const csvRows = journalData.slice(2).map((row) => row.map((cell, columnIndex) => {
		if (cell && typeof cell === 'object' && Object.prototype.hasOwnProperty.call(cell, 'value')) {
			return columnIndex === 3 ? formatCsvDate(cell.value) : cell.value;
		}
		return cell;
	}));
	const csvSheet = XLSX.utils.aoa_to_sheet(csvRows);
	const csvPath = path.join(path.dirname(outputPath), `${journalSheetName}.csv`);
	fs.writeFileSync(csvPath, XLSX.utils.sheet_to_csv(csvSheet), 'utf8');
	console.log(`Created "${path.basename(csvPath)}" in ${path.dirname(outputPath)}`);
}

function main() {
	if (!fs.existsSync(inputPath)) throw new Error(`Workbook not found: ${inputPath}`);
	const workbook = XLSX.readFile(inputPath, { cellDates: false });
	const result = analyse(workbook);
	const outputSheetName = nextOutputSheetName(workbook);
	appendSheetWithExcel(result.output, outputSheetName, 'ai-analyser');
	const journalSheetName = versionedSheetName(journalSheetBaseName, outputSheetName);
	appendSheetWithExcel(result.journalOutput, journalSheetName, 'ai-journal');
	exportJournalCsv(result.journalOutput, journalSheetName);
	console.log(`Created "${outputSheetName}" in ${outputPath}`);
	console.log(`Created "${journalSheetName}" in ${outputPath}`);
	console.log(`Analysed ${result.rowCount} remittances; unmapped remittances: ${result.unmappedCount}.`);
}

try {
	main();
} catch (error) {
	console.error(`AI analyser failed: ${error.message}`);
	process.exitCode = 1;
}
