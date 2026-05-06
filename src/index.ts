#!/usr/bin/env node

import fs from "fs";
import path from "path";
import puppeteer, { type Page } from "puppeteer";
import XLSX from "xlsx";
import StyledXLSX from "xlsx-js-style";

const YANDEX_MAPS_URL = "https://yandex.ru/maps/";
const SEARCH_INPUT_SELECTOR =
  'input[placeholder*="Поиск"], input[aria-label*="Поиск"], input.input__control';
const RESULTS_LIST_XPATH =
  "/html/body/div[1]/div[2]/div[10]/aside/div[1]/div[1]/div/div[1]/div/div/div[5]/div/div[5]/div/div/div/div[3]";
const REPORT_FILE_NAME = "result.xlsx";

interface CliOptions {
  file: string;
  outDir: string;
  column: number;
  startRow: number;
  additionalColumn: number;
}

interface SourceRow {
  address: string;
  additionalValue: string;
}

interface ParseResult {
  address: string;
  additionalValue: string;
  organizations: string[];
  error?: string;
}

function printHelp(): void {
  console.log(
    `
Usage:
  organization-scaner --file <file.xlsx> --out-dir <directory> --column <number> --row <number> [--additional <number>]
  organization-scaner <file.xlsx> <directory> <column-number> <row-number> [additional-column-number]

Options:
  -f, --file        Excel file to read
  -d, --out-dir     Directory for generated files
  -c, --column      Excel column number to process, starting from 1
  -r, --row         Excel row number to process, starting from 1
  -a, --additional  Additional Excel column number, 0 disables it
  -h, --help        Show this help
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

function parseStartRow(value: string | undefined): number {
  if (!value) {
    throw new Error("Start row number is required.");
  }

  const row = Number(value);
  if (!Number.isInteger(row) || row < 1) {
    throw new Error(
      "Start row number must be a positive integer starting from 1.",
    );
  }

  return row;
}

function parseAdditionalColumn(value: string | undefined): number {
  if (!value) {
    return 0;
  }

  const column = Number(value);
  if (!Number.isInteger(column) || column < 0) {
    throw new Error(
      "Additional column number must be 0 or a positive integer.",
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
    } else if (arg === "--row" || arg === "-r") {
      result.startRow = parseStartRow(args[i + 1]);
      i += 1;
    } else if (arg === "--additional" || arg === "-a") {
      result.additionalColumn = parseAdditionalColumn(args[i + 1]);
      i += 1;
    }
  }

  if (
    !result.file &&
    !result.outDir &&
    !result.column &&
    !result.startRow &&
    args.length >= 4
  ) {
    result.file = args[0];
    result.outDir = args[1];
    result.column = parseColumn(args[2]);
    result.startRow = parseStartRow(args[3]);
    result.additionalColumn = parseAdditionalColumn(args[4]);
  }

  if (!result.file || !result.outDir || !result.column || !result.startRow) {
    throw new Error(
      "Arguments required: --file <file.xlsx> --out-dir <directory> --column <number> --row <number>",
    );
  }

  return {
    file: result.file,
    outDir: result.outDir,
    column: result.column,
    startRow: result.startRow,
    additionalColumn: result.additionalColumn ?? 0,
  };
}

function cellValueToString(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

function formatProgress(
  done: number,
  total: number,
  foundRows: number,
  errorAddresses: number,
): string {
  const width = 28;
  const ratio = total === 0 ? 1 : done / total;
  const filled = Math.round(width * ratio);
  const bar = `${"#".repeat(filled)}${"-".repeat(width - filled)}`;
  const percent = Math.round(ratio * 100)
    .toString()
    .padStart(3, " ");

  return [
    `[${bar}] ${percent}%`,
    `обработано ${done}/${total}`,
    `кол-во ${foundRows}`,
    `Ошибки ${errorAddresses}`,
  ].join(" | ");
}

function writeProgress(
  done: number,
  total: number,
  foundRows: number,
  errorAddresses: number,
  currentAddress = "",
): void {
  const columns = process.stdout.columns || 120;
  const line = `${formatProgress(
    done,
    total,
    foundRows,
    errorAddresses,
  )} | ${currentAddress}`.slice(0, columns);

  process.stdout.write(`\r${line.padEnd(columns, " ")}`);
}

function readSourceRows(
  filePath: string,
  column: number,
  startRow: number,
  additionalColumn: number,
): SourceRow[] {
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
  const additionalColumnIndex = additionalColumn - 1;
  const startRowIndex = startRow - 1;

  return rows
    .slice(startRowIndex)
    .map((row) => {
      const cells = Array.isArray(row) ? row : [];

      return {
        address: cellValueToString(cells[columnIndex]),
        additionalValue:
          additionalColumn > 0
            ? cellValueToString(cells[additionalColumnIndex])
            : "",
      };
    })
    .filter((row) => row.address.length > 0);
}

async function waitForSearchResults(page: Page): Promise<void> {
  await page
    .waitForNetworkIdle({
      idleTime: 1_000,
      timeout: 30_000,
    })
    .catch(() => undefined);

  await page.waitForFunction(
    (resultsListXPath) => {
      const list = document.evaluate(
        resultsListXPath,
        document,
        null,
        XPathResult.FIRST_ORDERED_NODE_TYPE,
        null,
      ).singleNodeValue;

      if (list instanceof Element && list.querySelector("li")) {
        return true;
      }

      return document.querySelectorAll("li").length > 0;
    },
    {
      timeout: 30_000,
    },
    RESULTS_LIST_XPATH,
  );
}

async function revealAllOrganizationsInside(page: Page): Promise<void> {
  const wasClicked = await page.evaluate(() => {
    function text(value: string | null | undefined): string {
      return value?.replace(/\s+/g, " ").trim() ?? "";
    }

    const button = Array.from(
      document.querySelectorAll("aside [role='button'], aside button, aside a"),
    ).find(
      (element) => text(element.textContent) === "Все организации в этом доме",
    );

    if (button instanceof HTMLElement) {
      button.click();
      return true;
    }

    return false;
  });

  if (!wasClicked) {
    return;
  }

  await page
    .waitForNetworkIdle({
      idleTime: 1_000,
      timeout: 30_000,
    })
    .catch(() => undefined);
}

async function readResultNames(page: Page): Promise<string[]> {
  return page.evaluate((resultsListXPath) => {
    function text(value: string | null | undefined): string {
      return value?.replace(/\s+/g, " ").trim() ?? "";
    }

    function getElementByXPath(
      xpath: string,
      root: Document | Element,
    ): Element | null {
      const result = document.evaluate(
        xpath,
        root,
        null,
        XPathResult.FIRST_ORDERED_NODE_TYPE,
        null,
      );

      return result.singleNodeValue instanceof Element
        ? result.singleNodeValue
        : null;
    }

    const list = getElementByXPath(resultsListXPath, document);
    function findSectionStart(root: Element): Element | null {
      const elements = Array.from(root.querySelectorAll("*"));

      return (
        elements.find((element) => {
          const value = text(element.textContent);

          return (
            value === "Организации в здании" || value === "Организации внутри"
          );
        }) ?? null
      );
    }

    function isAfter(source: Element, target: Element): boolean {
      return Boolean(
        source.compareDocumentPosition(target) &
        Node.DOCUMENT_POSITION_FOLLOWING,
      );
    }

    function findOrganizationsInsideItems(root: Element): Element[] {
      const sectionStart = findSectionStart(root);

      return Array.from(root.querySelectorAll("li")).filter((item) => {
        if (item.closest(".toponym-service-orgs")) {
          return false;
        }

        if (!sectionStart) {
          return true;
        }

        return isAfter(sectionStart, item);
      });
    }

    const items = list
      ? findOrganizationsInsideItems(list)
      : Array.from(document.querySelectorAll("aside li")).filter((item) => {
          return !item.closest(".toponym-service-orgs");
        });

    const names = items
      .map((item) => {
        const elementFromSelector = item.querySelector(
          [
            '[class*="search-business-snippet-view__title"]',
            '[class*="search-business-snippet-view__head"]',
            '[class*="search-snippet-view__title"]',
            '[class*="business-snippet-view__title"]',
            '[class*="title"]',
          ].join(", "),
        );
        const elementFromXPath = getElementByXPath(
          ".//div/div/div/div/div[2]/div[1]",
          item,
        );

        return text(
          elementFromSelector?.textContent ?? elementFromXPath?.textContent,
        );
      })
      .filter((name) => name.length > 0);

    return Array.from(new Set(names));
  }, RESULTS_LIST_XPATH);
}

async function parse(value: string): Promise<string[]> {
  const browser = await puppeteer.launch({
    headless: true,
    defaultViewport: {
      width: 1366,
      height: 768,
    },
  });

  try {
    const page = await browser.newPage();
    await page.goto(YANDEX_MAPS_URL, {
      waitUntil: "networkidle2",
      timeout: 60_000,
    });

    await page.waitForSelector(SEARCH_INPUT_SELECTOR, {
      visible: true,
      timeout: 30_000,
    });
    await page.click(SEARCH_INPUT_SELECTOR, { clickCount: 3 });
    await page.keyboard.press("Backspace");
    await page.type(SEARCH_INPUT_SELECTOR, value, { delay: 20 });
    await page.keyboard.press("Enter");

    await waitForSearchResults(page);
    await revealAllOrganizationsInside(page);

    const names = await readResultNames(page);
    if (names.length === 0) {
      return [];
    }

    return names;
  } finally {
    await browser.close();
  }
}

function getResultBodyRows(result: ParseResult): string[] {
  if (result.organizations.length > 0) {
    return result.organizations;
  }

  return [result.error ? `Ошибка: ${result.error}` : "Организации не найдены"];
}

function buildReportRows(results: ParseResult[]): string[][] {
  const rows: string[][] = [];

  for (const result of results) {
    const bodyRows = getResultBodyRows(result);

    rows.push([result.address, "", ""]);

    for (let index = 0; index < bodyRows.length; index += 1) {
      rows.push([
        bodyRows[index],
        index === 0 ? String(result.organizations.length) : "",
        index === 0 ? result.additionalValue : "",
      ]);
    }

    rows.push([]);
  }

  return rows;
}

function getThinBorder(): Record<
  string,
  { style: string; color: { rgb: string } }
> {
  return {
    top: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
  };
}

function getCell(
  worksheet: XLSX.WorkSheet,
  row: number,
  column: number,
): XLSX.CellObject {
  const cellAddress = StyledXLSX.utils.encode_cell({ r: row, c: column });
  const cell = worksheet[cellAddress] as XLSX.CellObject | undefined;

  if (cell) {
    return cell;
  }

  const emptyCell: XLSX.CellObject = { t: "s", v: "" };
  worksheet[cellAddress] = emptyCell;

  return emptyCell;
}

function writeReport(results: ParseResult[], outputDir: string): string {
  const outputPath = path.join(outputDir, REPORT_FILE_NAME);
  const workbook = StyledXLSX.utils.book_new();
  const worksheet = StyledXLSX.utils.aoa_to_sheet(buildReportRows(results));
  const merges: XLSX.Range[] = [];
  const border = getThinBorder();

  let rowIndex = 0;
  for (const result of results) {
    const bodyRowsCount = getResultBodyRows(result).length;
    const tableStartRow = rowIndex;
    const bodyStartRow = rowIndex + 1;
    const tableEndRow = rowIndex + bodyRowsCount;
    const addressCell = getCell(worksheet, tableStartRow, 0);

    merges.push({
      s: { r: tableStartRow, c: 0 },
      e: { r: tableStartRow, c: 2 },
    });

    if (bodyRowsCount > 1) {
      merges.push({
        s: { r: bodyStartRow, c: 1 },
        e: { r: tableEndRow, c: 1 },
      });
      merges.push({
        s: { r: bodyStartRow, c: 2 },
        e: { r: tableEndRow, c: 2 },
      });
    }

    addressCell.s = {
      font: {
        bold: true,
        sz: 14,
      },
      alignment: {
        horizontal: "center",
        vertical: "center",
      },
      border,
    };

    for (let row = tableStartRow; row <= tableEndRow; row += 1) {
      for (let column = 0; column <= 2; column += 1) {
        const cell = getCell(worksheet, row, column);

        cell.s = {
          ...(cell.s ?? {}),
          border,
          alignment: {
            vertical: "center",
            wrapText: true,
            ...(column > 0 ? { horizontal: "center" } : {}),
          },
        };
      }
    }

    getCell(worksheet, bodyStartRow, 1).s = {
      ...(getCell(worksheet, bodyStartRow, 1).s ?? {}),
      font: {
        bold: true,
      },
      alignment: {
        horizontal: "center",
        vertical: "center",
      },
      border,
    };
    getCell(worksheet, bodyStartRow, 2).s = {
      ...(getCell(worksheet, bodyStartRow, 2).s ?? {}),
      alignment: {
        horizontal: "center",
        vertical: "center",
        wrapText: true,
      },
      border,
    };

    rowIndex += bodyRowsCount + 2;
  }

  worksheet["!merges"] = merges;
  worksheet["!cols"] = [{ wch: 80 }, { wch: 16 }, { wch: 28 }];
  StyledXLSX.utils.book_append_sheet(workbook, worksheet, "Result");
  StyledXLSX.writeFile(workbook, outputPath);

  return outputPath;
}

async function parseColumnValues(
  filePath: string,
  column: number,
  startRow: number,
  additionalColumn: number,
): Promise<ParseResult[]> {
  const rows = readSourceRows(filePath, column, startRow, additionalColumn);
  const results: ParseResult[] = [];
  let dataAddresses = 0;
  let foundRows = 0;
  let errorAddresses = 0;

  console.log(`Адресов к обработке: ${rows.length}`);
  writeProgress(0, rows.length, foundRows, errorAddresses);

  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];

    writeProgress(index, rows.length, foundRows, errorAddresses, row.address);

    try {
      const organizations = await parse(row.address);

      if (organizations.length > 0) {
        dataAddresses += 1;
        foundRows += organizations.length;
      }

      results.push({
        address: row.address,
        additionalValue: row.additionalValue,
        organizations,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);

      errorAddresses += 1;
      results.push({
        address: row.address,
        additionalValue: row.additionalValue,
        organizations: [],
        error: message,
      });
    }

    writeProgress(index + 1, rows.length, foundRows, errorAddresses);
  }

  process.stdout.write("\n");

  return results;
}

async function main(): Promise<void> {
  try {
    const { file, outDir, column, startRow, additionalColumn } = parseArgs(
      process.argv,
    );
    const filePath = path.resolve(file);
    const outputDir = path.resolve(outDir);

    fs.mkdirSync(outputDir, { recursive: true });

    console.log(`Файл: ${filePath}`);
    console.log(`Колонка с адресами: ${column}`);
    console.log(`Начальная строка: ${startRow}`);
    console.log(`Дополнительная колонка: ${additionalColumn}`);
    console.log(`Папка результата: ${outputDir}`);
    const results = await parseColumnValues(
      filePath,
      column,
      startRow,
      additionalColumn,
    );
    const reportPath = writeReport(results, outputDir);
    console.log(`Saved report to ${reportPath}`);
    console.log(`Finished. Output directory: ${outputDir}`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error("Error:", message);
    printHelp();
    process.exit(1);
  }
}

main();
