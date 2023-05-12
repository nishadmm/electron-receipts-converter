const path = require("path");
const os = require("os");
const fs = require("fs");
const pdf = require("pdf-parse");
const pdf2table = require("pdf2table");
const resizeImg = require("resize-img");
const ExcelJS = require("exceljs");
const pdfToExcel = require("pdf-to-excel");

const pdfParse = require("pdf-parse");
const XLSX = require("xlsx");
// import PDFParser from "pdf2json";

const { app, BrowserWindow, Menu, ipcMain, shell } = require("electron");

const isDev = process.env.NODE_ENV !== "production";
const isMac = process.platform === "darwin";

let mainWindow;

// Create the main window
function createMainWindow() {
  mainWindow = new BrowserWindow({
    title: "Image Resizer",
    width: isDev ? 1000 : 500,
    height: 600,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  //   open dev tools if in dev env
  if (isDev) {
    mainWindow.webContents.openDevTools();
  }

  mainWindow.loadFile(path.join(__dirname, "./renderer/index.html"));
}

// Create the about window
function createAboutWindow() {
  const aboutWindow = new BrowserWindow({
    title: "About Image Resizer",
    width: 300,
    height: 400,
  });

  aboutWindow.loadFile(path.join(__dirname, "./renderer/about.html"));
}

// App is ready
app.whenReady().then(() => {
  createMainWindow();

  // Implement menu
  const mainMenu = Menu.buildFromTemplate(menu);
  Menu.setApplicationMenu(mainMenu);

  mainWindow.on("closed", () => (mainMenu = null));

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
    }
  });
});

// Menu template
const menu = [
  ...(isMac
    ? [
        {
          label: app.name,
          submenu: [
            {
              label: "About",
              click: createAboutWindow,
            },
          ],
        },
      ]
    : []),
  {
    label: "File",
    submenu: [
      {
        label: "quit",
        click: () => app.quit(),
        accelerator: "CmdOrCtrl+W",
      },
    ],
  },
  ...(isMac
    ? []
    : [
        {
          label: "Help",
          submenu: [
            {
              label: "About",
              click: createAboutWindow,
            },
          ],
        },
      ]),
  // {
  //   role: "fileMenu",
  // },
];

// Respond to ipcRenderer resize
ipcMain.on("receipt:convert", (e, options) => {
  options.dest = path.join(os.homedir(), "receipt-converter");
  resizeImage(options);
});

app.on("window-all-closed", () => {
  if (!isMac) {
    app.quit();
  }
});

function convertPDFtoExcel(pdfPath, dest, filesPath) {
  // Get filename
  // const filename = path.basename(pdfPath);

  // Create a new workbook and worksheet
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet([[""]]);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  let allColumns = [];

  // Read the PDF file and extract its text
  filesPath.forEach((singlePath, index) => {
    const pdfData = fs.readFileSync(singlePath);
    pdfParse(pdfData).then(function (data) {
      // Set the width of the first column to 100 units
      // worksheet["!cols"] = [{ width: 100 }, {}, {}];

      // Split the text into rows and columns
      const rows = data.text.split("\n");
      allColumns = [...allColumns, ...rows.map((row) => row.split(/\s+/))];

      if (index + 1 === filesPath.length) {
        // Add the rows and columns to the worksheet
        XLSX.utils.sheet_add_aoa(worksheet, allColumns);
        // Save the Excel file
        const excelData = XLSX.write(workbook, {
          bookType: "xlsx",
          type: "buffer",
        });
        fs.writeFileSync(path.join(dest, `${"receiptsExcel"}.xlsx`), excelData);

        // Send success to renderer
        mainWindow.webContents.send("receipt:done");

        // open the folder in the file explorer
        shell.openPath(dest);
      }
    });
  });

  // const pdfData = fs.readFileSync(pdfPath);
  // pdfParse(pdfData).then(function (data) {
  //   // Create a new workbook and worksheet
  //   const workbook = XLSX.utils.book_new();
  //   const worksheet = XLSX.utils.aoa_to_sheet([[""]]);

  //   // Set the width of the first column to 100 units
  //   // worksheet["!cols"] = [{ width: 100 }, {}, {}];

  //   // Split the text into rows and columns
  //   const rows = data.text.split("\n");
  //   const columns = rows.map((row) => row.split(/\s+/));

  //   // Add the rows and columns to the worksheet
  //   XLSX.utils.sheet_add_aoa(worksheet, columns);
  //   XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  //   // Save the Excel file
  //   const excelData = XLSX.write(workbook, {
  //     bookType: "xlsx",
  //     type: "buffer",
  //   });
  //   fs.writeFileSync(path.join(dest, `${"receiptsExcel"}.xlsx`), excelData);

  //   // Send success to renderer
  //   mainWindow.webContents.send("receipt:done");

  //   // open the folder in the file explorer
  //   shell.openPath(dest);
  // });
}

async function resizeImage({ imgPath, dest, filesPath }) {
  try {
    convertPDFtoExcel(imgPath, dest, filesPath);
  } catch (error) {
    console.log(error);
  }
}

// async function resizeImage({ imgPath, dest }) {
//   try {
//     let dataBuffer = fs.readFileSync(imgPath);

//     pdf(dataBuffer).then(function (data) {
//       // Convert the PDF to a table using pdf2table
//       const table = pdf2table.parse(data.text);

//       console.log(table, data);

//       const workbook = new ExcelJS.Workbook();
//       const worksheet = workbook.addWorksheet("Receipts");

//       // Loop through each row of the table
//       for (let i = 0; i < table.length; i++) {
//         const row = table[i];

//         // Loop through each cell of the row
//         for (let j = 0; j < row.length; j++) {
//           const cell = row[j];

//           // Set the value of the corresponding cell in the worksheet
//           worksheet.getCell(`${String.fromCharCode(65 + j)}${i + 1}`).value =
//             cell;
//         }
//       }

//       // // Define headers
//       // worksheet.columns = [
//       //   { header: "Receipt Number", key: "receiptNumber", width: 20 },
//       //   { header: "Date", key: "date", width: 15 },
//       //   { header: "Total", key: "total", width: 15 },
//       //   { header: "Items", key: "items", width: 50 },
//       // ];

//       // worksheet.addRow({
//       //   receiptNumber: "data",
//       //   date: "data",
//       //   total: "data",
//       //   items: "data",
//       // });

//       // // Resize image
//       // const newPath = await resizeImg(fs.readFileSync(imgPath), {
//       //   width: +width,
//       //   height: +height,
//       // });

//       workbook.xlsx
//         .writeFile("receipts.xlsx")
//         .then(() => {
//           // Send success to renderer
//           mainWindow.webContents.send("receipt:done");

//           // open the folder in the file explorer
//           // shell.openPath("");
//         })
//         .catch((error) => {
//           console.error(error);
//         });

//       // // Get filename
//       // const filename = path.basename(imgPath);

//       // // Create destination folder if it doesn't exist
//       // if (!fs.existsSync(dest)) {
//       //   fs.mkdirSync(dest);
//       // }

//       // // Write the file to the destination folder
//       // fs.writeFileSync(path.join(dest, filename), workbook);
//     });
//   } catch (error) {
//     console.log(error);
//   }
// }

// Create a new PDFParser instance
// let pdfParser = new PDFParser();
// async function resizeImage({ imgPath, dest }) {
//   try {
//     // Parse the PDF file
//     let dataBuffer = fs.readFileSync(imgPath);
//     pdfParser.parseBuffer(dataBuffer);

//     // Set up a callback for when the PDF has been parsed
//     pdfParser.on("pdfParser_dataReady", (pdfData) => {
//       // Create a new Excel workbook
//       const workbook = new ExcelJS.Workbook();

//       // Add a new worksheet to the workbook
//       const worksheet = workbook.addWorksheet("Sheet1");

//       // Loop through each page of the PDF
//       for (let i = 0; i < pdfData.formImage.Pages.length; i++) {
//         const page = pdfData.formImage.Pages[i];

//         // Loop through each text block on the page
//         for (let j = 0; j < page.Texts.length; j++) {
//           const text = page.Texts[j];
//           const textValue = decodeURIComponent(text.R[0].T);

//           // Set the value of the corresponding cell in the worksheet
//           worksheet.getCell(`A${j + 1}`).value = textValue;
//         }
//       }

//       // Write the workbook to a new Excel file
//       workbook.xlsx
//         .writeFile(app.getPath("documents") + "/output.xlsx")
//         .then(() => {
//           console.log("Excel file saved successfully!");

//           // Send success to renderer
//           mainWindow.webContents.send("receipt:done");

//           // open the folder in the file explorer
//           // shell.openPath("");
//         })
//         .catch((error) => {
//           console.error(error);
//         });
//     });
//   } catch (error) {
//     console.log(error);
//   }
// }
