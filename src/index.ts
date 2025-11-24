import * as ExcelJS from "exceljs";

/**
 * Validates XLSX file format by checking if all required headers are present
 * @param buffer - The XLSX file buffer
 * @param requiredHeaders - Array of required header names
 * @param headerLineNumber - Row number containing the headers (default: 1)
 * @returns Object mapping header names to their column indices (1-based, compatible with ExcelJS getCell())
 * @throws Error if any required headers are missing
 */
export async function validateXlsxHeaders(
	buffer: Buffer,
	requiredHeaders: string[],
	headerLineNumber: number = 1
): Promise<Record<string, number>> {
	try {
		/* Load the workbook from buffer */
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(buffer as any);

		/* Get the first worksheet */
		const worksheet = workbook.getWorksheet(1);
		if (!worksheet) {
			throw new Error("XLSX file must contain at least one worksheet");
		}

		/* Reuse the worksheet validation logic */
		return validateWorksheetHeaders(
			worksheet,
			requiredHeaders,
			headerLineNumber
		);
	} catch (error) {
		if (error instanceof Error) {
			throw error;
		}
		throw new Error(
			`Failed to validate XLSX file: ${(error as Error).message}`
		);
	}
}

/**
 * Alternative function that accepts worksheet directly
 * @param worksheet - ExcelJS worksheet object
 * @param requiredHeaders - Array of required header names
 * @param headerLineNumber - Row number containing the headers (default: 1)
 * @returns Object mapping header names to their column indices (1-based, compatible with ExcelJS getCell())
 */
export function validateWorksheetHeaders(
	worksheet: ExcelJS.Worksheet,
	requiredHeaders: string[],
	headerLineNumber: number = 1
): Record<string, number> {
	/* Get the specified header row */
	const headerRow = worksheet.getRow(headerLineNumber);
	if (!headerRow || headerRow.cellCount === 0) {
		throw new Error(
			`Worksheet must contain a header row at line ${headerLineNumber}`
		);
	}

	/* Extract header values and create a map */
	const headerMap: Record<string, number> = {};
	const foundHeaders: string[] = [];

	headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
		const cellValue = cell.text.trim();
		if (cellValue) {
			/* Keep 1-based index for ExcelJS getCell() compatibility */
			headerMap[cellValue] = colNumber;
			foundHeaders.push(cellValue);
		}
	});

	/* Check required headers */
	const missingHeaders: string[] = [];
	const resultMap: Record<string, number> = {};

	for (const requiredHeader of requiredHeaders) {
		if (headerMap.hasOwnProperty(requiredHeader)) {
			resultMap[requiredHeader] = headerMap[requiredHeader] as number;
		} else {
			missingHeaders.push(requiredHeader);
		}
	}

	if (missingHeaders.length > 0) {
		throw new Error(
			`Missing required headers: ${missingHeaders.join(", ")}. ` +
				`Found headers: ${foundHeaders.join(", ")}`
		);
	}

	return resultMap;
}
