#!/usr/bin/env node

import fs from "fs";
import path from "path";
import XLSX from "xlsx";

interface CliOptions {
  file: string;
  outDir: string;
  column: number;
}

function printHelp(): void {
  console.log(
    `
Usage:
  organization-scaner --file <file.xlsx> --out-dir <directory> --column <number>
  organization-scaner <file.xlsx> <directory> <number>

Options:
  -f, --file      Excel file to read
  -d, --out-dir   Directory for generated files
  -c, --column    Excel column number to process, starting from 1
  -h, --help      Show this help
`.trim(),
  );
}

function parseColumn(value: string | undefined): number {
  if (!value) {
    throw new Error("Column number is required.");
  }

  const column = Number(value);
  if (!Number.isInteger(column) || column < 1) {
    throw new Error(
      "Column number must be a positive integer starting from 1.",
    );
  }

  return column;
}

function parseArgs(argv: string[]): CliOptions {
  const args = argv.slice(2);
  const result: Partial<CliOptions> = {};

  if (args.includes("--help") || args.includes("-h")) {
    printHelp();
    process.exit(0);
  }

  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === "--file" || arg === "-f") {
      result.file = args[i + 1];
      i += 1;
    } else if (arg === "--out-dir" || arg === "-d") {
      result.outDir = args[i + 1];
      i += 1;
    } else if (arg === "--column" || arg === "-c") {
      result.column = parseColumn(args[i + 1]);
      i += 1;
    }
  }

  if (!result.file && !result.outDir && !result.column && args.length >= 3) {
    result.file = args[0];
    result.outDir = args[1];
    result.column = parseColumn(args[2]);
  }

  if (!result.file || !result.outDir || !result.column) {
    throw new Error(
      "Arguments required: --file <file.xlsx> --out-dir <directory> --column <number>",
    );
  }

  return {
    file: result.file,
    outDir: result.outDir,
    column: result.column,
  };
}

function cellValueToString(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

function readColumnValues(filePath: string, column: number): string[] {
  const workbook = XLSX.readFile(filePath);
  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    throw new Error("Excel file does not contain worksheets.");
  }

  const worksheet = workbook.Sheets[firstSheetName];
  if (!worksheet) {
    throw new Error(`Worksheet "${firstSheetName}" was not found.`);
  }

  const rows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    blankrows: false,
  });
  const columnIndex = column - 1;

  return rows
    .map((row) => cellValueToString(Array.isArray(row) ? row[columnIndex] : ""))
    .filter((value) => value.length > 0);
}

async function parse(value: string): Promise<void> {
  console.log(value);
}

async function parseColumnValues(
  filePath: string,
  column: number,
): Promise<void> {
  const values = readColumnValues(filePath, column);

  for (const value of values) {
    await parse(value);
  }
}

async function main(): Promise<void> {
  try {
    const { file, outDir, column } = parseArgs(process.argv);
    const filePath = path.resolve(file);
    const outputDir = path.resolve(outDir);

    fs.mkdirSync(outputDir, { recursive: true });

    console.log(`Reading column ${column} from ${filePath}...`);
    await parseColumnValues(filePath, column);
    console.log(`Finished. Output directory: ${outputDir}`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error("Error:", message);
    printHelp();
    process.exit(1);
  }
}

main();
