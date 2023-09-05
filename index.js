const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

// Create a new workbook and a new worksheet
let workbook = new ExcelJS.Workbook();
let worksheet = workbook.addWorksheet("Sheet1");
let outputDir = path.join(__dirname, "output");
let outputPath = path.join(outputDir, "output.xlsx");
let inputDir = path.join(__dirname, "inputs");
let inputData = [];

// Function to read and process an input file
async function processFile(filePath) {
	let workbook = new ExcelJS.Workbook();
	await workbook.xlsx.readFile(filePath);
	let worksheet = workbook.getWorksheet(2);
	const data = {
		B: worksheet.getCell("C11").value?.result || worksheet.getCell("C11").value,
		C: worksheet.getCell("D11").value?.result || worksheet.getCell("D11").value,
		D: worksheet.getCell("C13").value?.result || worksheet.getCell("C13").value,
		E: worksheet.getCell("D13").value?.result || worksheet.getCell("D13").value,
		F: worksheet.getCell("D14").value?.result || worksheet.getCell("D14").value,
		G: worksheet.getCell("D16").value?.result || worksheet.getCell("D16").value,
		H: worksheet.getCell("D18").value?.result || worksheet.getCell("D18").value,
		I: worksheet.getCell("C21").value?.result || worksheet.getCell("C21").value,
		J: worksheet.getCell("D21").value?.result || worksheet.getCell("D21").value,
		K: worksheet.getCell("B24").value?.result || worksheet.getCell("B24").value,
	};
	return data;
}

// Function to write data to the output file
async function writeDataToFile(data) {
	let workbook = new ExcelJS.Workbook();
	await workbook.xlsx.readFile(outputPath);
	let worksheet = workbook.getWorksheet("Sheet1");
	let initRowCount = worksheet.rowCount;
	console.log(
		`writing to output file, starting from row number ${
			initRowCount + 1
		} 📝 ...`
	);
	// // Write each row of data to the worksheet
	data.forEach((rowData, rowIndex) => {
		let row = worksheet.getRow(initRowCount + rowIndex + 1);
		row.getCell("B").value = rowData.B;
		row.getCell("C").value = rowData.C;
		row.getCell("D").value = rowData.D;
		row.getCell("E").value = rowData.E;
		row.getCell("F").value = rowData.F;
		row.getCell("G").value = rowData.G;
		row.getCell("H").value = rowData.H;
		row.getCell("I").value = rowData.I;
		row.getCell("J").value = rowData.J;
		row.getCell("K").value = rowData.K;

		row.commit();
	});
	await workbook.xlsx.writeFile(outputPath);
	console.log("SUCCESS ✅");
	console.log("PROGRAM END 🎬");
}

// main entry point function for the script
const main = async () => {
	console.log("SCRIPT STARTED");
	// Create the output directory if it does not exist
	if (!fs.existsSync(outputDir)) {
		console.log(
			"no output file found. Creating new file ./output/output.xlsx ..."
		);
		fs.mkdirSync(outputDir, { recursive: true });
		// Write 'Name' and 'Address' to cells A1 and B1 with bold formatting
		worksheet.getCell("A1").value = "Name";
		worksheet.getCell("A1").font = { bold: true };
		worksheet.getCell("B1").value = "C11";
		worksheet.getCell("B1").font = { bold: true };
		worksheet.getCell("C1").value = "D11";
		worksheet.getCell("C1").font = { bold: true };
		worksheet.getCell("D1").value = "C13";
		worksheet.getCell("D1").font = { bold: true };
		worksheet.getCell("E1").value = "D13";
		worksheet.getCell("E1").font = { bold: true };
		worksheet.getCell("F1").value = "D14";
		worksheet.getCell("F1").font = { bold: true };
		worksheet.getCell("G1").value = "D16";
		worksheet.getCell("G1").font = { bold: true };
		worksheet.getCell("H1").value = "D18";
		worksheet.getCell("H1").font = { bold: true };
		worksheet.getCell("I1").value = "C21";
		worksheet.getCell("I1").font = { bold: true };
		worksheet.getCell("J1").value = "D21";
		worksheet.getCell("J1").font = { bold: true };
		worksheet.getCell("K1").value = "B24";
		worksheet.getCell("K1").font = { bold: true };

		try {
			workbook.xlsx
				.writeFile(outputPath)
				.then(async () => {
					console.log(
						"Output file initiated 🎉 . Reading input files in ./inputs folder "
					);
				})
				.catch((err) => console.error(err));
		} catch (err) {
			console.error("err: ", err);
		}
	} else {
		console.log(`output file "./outputs/output.xlsx" found 🎉`);
	}
	console.log("starting to read input files 📖 ...");
	let files = fs.readdirSync(inputDir);
	const xlsxFiles = files.map((item) => {
		if (path.extname(item) === ".xlsx") return item;
	});
	if (!xlsxFiles?.length) {
		console.log("❌ No xlsx input files found.");
		console.log("PROGRAM END 🎬");
		return;
	} else {
		console.log(`${xlsxFiles.length} number of xlsx files found 🫡 ...`);
		for (let file of xlsxFiles) {
			let filePath = path.join(inputDir, file);
			let fileData = await processFile(filePath);
			inputData.push(fileData);
			console.log(`✅ read data from file ${path.basename(file)} ...`);
		}
	}

	writeDataToFile(inputData);
};

main();
