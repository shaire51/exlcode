const express = require("express"); // 建立伺服器
const cors = require("cors"); // 允許不同來源的前端呼叫 API
const XLSX = require("xlsx"); // 讀取 Excel
const path = require("path"); // 處理檔案路徑

const app = express();

app.use(cors());
app.use(express.json());

// 讀取 Excel 裡面指定的工作表
function readSheetFromExcel(sheetName) {
  const filePath = path.join(__dirname, "test.xlsx");

  console.log("讀取檔案位置:", filePath);

  const workbook = XLSX.readFile(filePath);

  console.log("所有工作表:", workbook.SheetNames);

  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    throw new Error(`找不到工作表：${sheetName}`);
  }

  const data = XLSX.utils.sheet_to_json(sheet, {
    defval: "", // 空白欄位也保留
  });

  return data;
}

function generateParkingOrder({
  currentApplicants,
  lastParkingList,
  parkingCount,
}) {
  const getId = (p) => String(p.員編).trim();

  const currentIds = new Set(currentApplicants.map(getId));
  const lastIds = new Set(lastParkingList.map(getId));

  // 上期有，本期沒有：未申請，從輪序移除
  const notAppliedThisSeason = lastParkingList.filter((p) => {
    return !currentIds.has(getId(p));
  });

  // 上期沒有，本期有：新增申請
  const newApplicants = currentApplicants.filter((p) => {
    return !lastIds.has(getId(p));
  });

  // 上期有，本期也有：保留原輪序
  const continuedList = lastParkingList.filter((p) => {
    return currentIds.has(getId(p));
  });

  // 確保照「順位」排序，避免 Excel 順序亂掉
  continuedList.sort((a, b) => Number(a.順位) - Number(b.順位));

  // 新申請者放到輪序最後面
  const newApplicantRows = newApplicants.map((p) => ({
    ...p,
    順位: "",
    車位號碼: "",
    上期狀態: "新增申請",
  }));

  const fullOrder = [...continuedList, ...newApplicantRows];

  // 本季基本沒車位人數：58 - 53 = 5
  const baseNoParkingCount = Math.max(
    currentApplicants.length - parkingCount,
    0,
  );

  // 上季替補人數
  const lastWaitingCount = lastParkingList.filter((p) => {
    return String(p.上期狀態).trim() === "後補";
  }).length;

  // 本季要被擠出去當替補的人數
  // 例：5 + 6 + 1 - 5 = 7
  const pushOutCount =
    baseNoParkingCount +
    lastWaitingCount +
    newApplicants.length -
    notAppliedThisSeason.length;

  const safePushOutCount = Math.max(pushOutCount, 0);

  // 重點：前面的人被擠出去當替補
  // 例：1~7 變替補
  const waiting = fullOrder.slice(0, safePushOutCount).map((p, index) => ({
    ...p,
    本期狀態: "替補",
    本期順位: index + 1,
  }));

  // 後面的人往前補成正選
  // 例：8 之後變正選
  const assigned = fullOrder.slice(safePushOutCount).map((p, index) => ({
    ...p,
    本期狀態: "正選",
    本期順位: index + 1,
  }));

  return {
    summary: {
      本季申請人數: currentApplicants.length,
      停車位數: parkingCount,
      本季基本沒車位人數: baseNoParkingCount,
      上季替補人數: lastWaitingCount,
      本季新增申請人數: newApplicants.length,
      本季未申請人數: notAppliedThisSeason.length,
      本季擠出替補人數: safePushOutCount,
    },
    assigned,
    waiting,
    notAppliedThisSeason,
    newApplicants,
  };
}

app.get("/api/a", (req, res) => {
  try {
    const currentApplicants = readSheetFromExcel("申請資料");
    const lastParkingList = readSheetFromExcel("車位資料");

    const result = generateParkingOrder({
      currentApplicants,
      lastParkingList,
      parkingCount: 53,
    });

    res.json(result);
  } catch (error) {
    console.error("讀取或計算失敗:", error);

    res.status(500).json({
      message: "讀取或計算失敗",
      error: error.message,
    });
  }
});

app.listen(3001, () => {
  console.log("Server running on http://localhost:3001/api/a");
});
