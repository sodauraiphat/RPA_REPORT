require("dotenv").config();
const sql = require("mssql");
const ExcelJS = require("exceljs");
const cron = require("node-cron");
const fs = require("fs");
const path = require("path"); // added a path
const config = {
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_DATABASE,
  server: process.env.DB_SERVER,
  options: {
    instanceName: process.env.DB_INSTANCE,
    encrypt: false,
    trustServerCertificate: true,
  },
};

async function runRpa(daysToProcess = []) {
  let pool;
  try {
    const now = new Date();
    const monthNames = [
      "Jan",
      "Feb",
      "Mar",
      "Apr",
      "May",
      "Jun",
      "Jul",
      "Aug",
      "Sep",
      "Oct",
      "Nov",
      "Dec",
    ];
    const currentMonth = monthNames[now.getMonth()];
    const currentYear = now.getFullYear();

    const fileName = `A Daily production report ${currentMonth} ${currentYear}.xlsx`;
    const filePath = path.join(__dirname, "ChicagoReport", fileName);
    
    if (!fs.existsSync(filePath))
      throw new Error(`Template missing: ${fileName}`);

    // Load the existing updated file if it exists, otherwise the template
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    pool = await sql.connect(config);

    // If no specific days passed, default to only today
    if (daysToProcess.length === 0) {
      daysToProcess = [now.getDate()];
    }

    for (const day of daysToProcess) {
      const dateStr = `${currentYear}-${(now.getMonth() + 1).toString().padStart(2, "0")}-${day.toString().padStart(2, "0")}`;
      console.log(
        `[${new Date().toLocaleTimeString()}]  Syncing Data: ${dateStr} (Sheet ${day})`,
      );

      const result = await pool.request().query(`
                SELECT LineType, CustKey, OrderNumber, OrderQty,TotalBags, ProductCode, 
                       StartTime, EndTime, PartialBagsKG, TotalStdUnitKG 
                FROM [dbo].[Chicago_Report_INF] 
                WHERE CAST(ReportDate AS DATE) = '${dateStr}'
                ORDER BY StartTime ASC
            `);
      
      const worksheet = workbook.worksheets.find(
        (s) => s.name.trim() === day.toString(),
      );
      if (!worksheet) continue;

      // Clear rows 5-500 to ensure fresh data
      for (let i = 5; i <= 500; i++) {
        const row = worksheet.getRow(i);
        [
          "B",
          "C",
          "D",
          "E",
          "F",
          "H",
          "I",
          "J",
          "K",
          "L",
          "Y",
          "Z",
          "AA",
          "AB",
          "AC",
          "AE",
          "AG",
          "AF",
          "AH",
          "AI",
        ].forEach((col) => {
          row.getCell(col).value = null;
        });
      }
      
      //start the 5th row
      let leftRow = 5;
      let rightRow = 5;

      result.recordset.forEach((row) => {
        let hour = 0;
        if (row.StartTime instanceof Date) {
          hour = row.StartTime.getHours();
        } else if (typeof row.StartTime === "string") {
          const timePart = row.StartTime.includes(" ")
            ? row.StartTime.split(" ")[1]
            : row.StartTime;
          hour = parseInt(timePart.split(":")[0]);
        }

        const val = (v) => (v == null ? "" : v);
        const isDayShift = hour >= 7 && hour <= 18;
        const targetRow = isDayShift ? leftRow++ : rightRow++;
        const excelRow = worksheet.getRow(targetRow);

        const map = isDayShift  //Mapping data to each column
          ? {
              B: "LineType",
              C: "CustKey",
              D: "OrderNumber",
              E: "OrderQty",
              F: "ProductCode",
              H: "StartTime",
              I: "EndTime",
              J: "TotalBags",
              K: "PartialBagsKG",
              L: "TotalStdUnitKG",
            }
          : {
              Y: "LineType",
              Z: "CustKey",
              AA: "OrderNumber",
              AB: "OrderQty",
              AC: "ProductCode",
              AE: "StartTime",
              AF: "EndTime",
              AG: "TotalBags",
              AH: "PartialBagsKG",
              AI: "TotalStdUnitKG",
            };

        Object.keys(map).forEach((col) => {
          excelRow.getCell(col).value = val(row[map[col]]);
        });
      });
    }

    
    
    // Set Active Sheet to Today's Date
    const todayStr = new Date().getDate().toString();
    const activeSheetIndex = workbook.worksheets.findIndex(s => s.name.trim() === todayStr);

    if (activeSheetIndex !== -1) {
      workbook.views = [
        {
          x: 0, y: 0, width: 10000, height: 20000,
          firstSheet: 0, 
          activeTab: activeSheetIndex, 
          visibility: 'visible'
        }
      ];
    }

    //save the file locally
    await workbook.xlsx.writeFile(filePath);

    // --- ONLY EXPORT TO THIS PATH ---
    const finalDestination = "\\\\th-bp-filesvr.nwfth.com\\Production\\DPR\\" + fileName;

    try {
      fs.copyFileSync(filePath, finalDestination);
      console.log(`Successfully exported to: ${finalDestination}`);
    } catch (exportErr) {
      console.error(`Export failed: ${exportErr.message}`);
    }

    console.log(`completed ${fileName}`);
  } catch (err) {
    console.error("RPA Error:", err.message);
  } finally {
    if (pool) await pool.close();
  }
}

async function start() {
  const now = new Date();
  const today = now.getDate();

  console.log(`System Startup: Syncing month-to-date (Days 1 to ${today})...`);
  let syncDays = [];
  for (let d = 1; d <= today; d++) {
    syncDays.push(d);
  }
  await runRpa(syncDays);

  console.log(
    "Startup sync complete. Now monitoring today's data every 10 minutes...",
  );
  cron.schedule("*/10 * * * *", () => {
    runRpa([new Date().getDate()]);
  });
}

start();
