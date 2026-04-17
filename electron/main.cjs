const { app, BrowserWindow, dialog, ipcMain } = require("electron");
const fs = require("node:fs/promises");
const path = require("node:path");
const ExcelJS = require("exceljs");

const isDev = Boolean(process.env.VITE_DEV_SERVER_URL);
let mainWindow = null;

function getActiveWindow() {
  return BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0] || mainWindow || null;
}

function mapPrinter(printerInfo = {}) {
  return {
    name: String(printerInfo.name || ""),
    displayName: String(printerInfo.displayName || printerInfo.name || ""),
    description: String(printerInfo.description || ""),
    status: Number.isFinite(Number(printerInfo.status)) ? Number(printerInfo.status) : 0,
    isDefault: Boolean(printerInfo.isDefault),
    isVirtual: /pdf|xps|onenote/i.test(
      `${printerInfo.name || ""} ${printerInfo.displayName || ""} ${printerInfo.description || ""}`
    )
  };
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1560,
    height: 980,
    minWidth: 1280,
    minHeight: 820,
    title: "档案标签排版器",
    backgroundColor: "#efe7dc",
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  if (isDev) {
    mainWindow.loadURL(process.env.VITE_DEV_SERVER_URL);
  } else {
    mainWindow.loadFile(path.join(__dirname, "..", "dist", "index.html"));
  }

  mainWindow.on("closed", () => {
    mainWindow = null;
  });
}

function sanitizeHeader(header, index, seenHeaders) {
  const baseLabel = String(header || "").trim() || `列${index}`;
  const nextCount = (seenHeaders.get(baseLabel) || 0) + 1;
  seenHeaders.set(baseLabel, nextCount);

  const label = nextCount > 1 ? `${baseLabel} (${nextCount})` : baseLabel;
  return {
    key: `field_${index}`,
    label,
    excelIndex: index
  };
}

function cellToText(cell) {
  if (!cell) {
    return "";
  }

  if (typeof cell.text === "string" && cell.text.trim()) {
    return cell.text.trim();
  }

  const { value } = cell;

  if (value == null) {
    return "";
  }

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "object") {
    if (Array.isArray(value.richText)) {
      return value.richText.map((item) => item.text || "").join("").trim();
    }

    if (typeof value.text === "string") {
      return value.text.trim();
    }

    if (typeof value.result !== "undefined") {
      return String(value.result).trim();
    }
  }

  return String(value).trim();
}

function isBlankRow(row, columnCount) {
  for (let index = 1; index <= columnCount; index += 1) {
    if (cellToText(row.getCell(index))) {
      return false;
    }
  }

  return true;
}

function findHeaderRowNumber(worksheet) {
  let headerRowNumber = null;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (row.values.filter(Boolean).length > 0 && headerRowNumber == null) {
      headerRowNumber = rowNumber;
    }
  });

  return headerRowNumber || 1;
}

async function readWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  const extension = path.extname(filePath).toLowerCase();

  if (extension === ".csv") {
    await workbook.csv.readFile(filePath);
  } else {
    await workbook.xlsx.readFile(filePath);
  }

  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error("未在 Excel 文件中找到工作表。");
  }

  const headerRowNumber = findHeaderRowNumber(worksheet);
  const headerRow = worksheet.getRow(headerRowNumber);
  const seenHeaders = new Map();
  const columns = [];

  for (let index = 1; index <= headerRow.cellCount; index += 1) {
    const header = sanitizeHeader(cellToText(headerRow.getCell(index)), index, seenHeaders);
    columns.push(header);
  }

  const rows = [];

  for (let rowNumber = headerRowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);

    if (isBlankRow(row, columns.length)) {
      continue;
    }

    const record = {
      __rowNumber: rowNumber,
      __index: rows.length + 1
    };

    columns.forEach((column) => {
      record[column.key] = cellToText(row.getCell(column.excelIndex));
    });

    rows.push(record);
  }

  return {
    filePath,
    fileName: path.basename(filePath),
    sheetName: worksheet.name,
    rowCount: rows.length,
    columns,
    rows
  };
}

ipcMain.handle("workbook:choose", async () => {
  const browserWindow = getActiveWindow();
  const result = await dialog.showOpenDialog(browserWindow, {
    title: "选择学生信息 Excel",
    properties: ["openFile"],
    filters: [
      {
        name: "Excel 文件",
        extensions: ["xlsx", "xls", "csv"]
      }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  return readWorkbook(result.filePaths[0]);
});

ipcMain.handle("workbook:read", async (_event, filePath) => {
  if (!filePath) {
    throw new Error("文件路径不能为空。");
  }

  return readWorkbook(filePath);
});

ipcMain.handle("template:save", async (_event, payload) => {
  const browserWindow = getActiveWindow();
  const { canceled, filePath } = await dialog.showSaveDialog(browserWindow, {
    title: "保存排版模板",
    defaultPath: "档案标签模板.archive-label.json",
    filters: [
      {
        name: "标签模板文件",
        extensions: ["json"]
      }
    ]
  });

  if (canceled || !filePath) {
    return null;
  }

  await fs.writeFile(filePath, JSON.stringify(payload, null, 2), "utf8");
  return filePath;
});

ipcMain.handle("template:load", async () => {
  const browserWindow = getActiveWindow();
  const result = await dialog.showOpenDialog(browserWindow, {
    title: "选择排版模板",
    properties: ["openFile"],
    filters: [
      {
        name: "标签模板文件",
        extensions: ["json"]
      }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    return null;
  }

  const rawContent = await fs.readFile(result.filePaths[0], "utf8");
  return JSON.parse(rawContent);
});

ipcMain.handle("system:list-printers", async () => {
  const browserWindow = getActiveWindow();

  if (!browserWindow) {
    throw new Error("当前没有可用窗口，无法读取打印机列表。");
  }

  const printers = await browserWindow.webContents.getPrintersAsync();
  return printers.map((printer) => mapPrinter(printer));
});

ipcMain.handle("system:print", async (_event, payload = {}) => {
  const browserWindow = getActiveWindow();

  if (!browserWindow) {
    throw new Error("当前没有可打印的窗口。");
  }

  const {
    deviceName = "",
    silent = false,
    usePrinterDefaultPageSize = false
  } = payload || {};

  browserWindow.show();
  browserWindow.focus();

  const invokePrint = (options) =>
    new Promise((resolve, reject) => {
      setTimeout(() => {
        browserWindow.webContents.print(options, (success, errorType) => {
          if (!success) {
            reject(new Error(errorType || "打印任务未完成。"));
            return;
          }

          resolve(true);
        });
      }, 120);
    });

  const baseOptions = {
    silent: Boolean(silent),
    printBackground: true,
    landscape: false
  };

  if (deviceName) {
    baseOptions.deviceName = deviceName;
  }

  if (usePrinterDefaultPageSize) {
    baseOptions.usePrinterDefaultPageSize = true;
  } else {
    baseOptions.pageSize = "A4";
  }

  try {
    await invokePrint(baseOptions);
    return true;
  } catch (error) {
    if (!/Invalid printer settings/i.test(error.message || "")) {
      throw error;
    }

    await invokePrint({
      ...baseOptions,
      usePrinterDefaultPageSize: true,
      pageSize: undefined
    });
    return true;
  }
});

app.whenReady().then(() => {
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});
