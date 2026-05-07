const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");
const path = require("path");

const app = express();

app.use(cors());
app.use(express.json());

// 讀取 Excel 指定工作表
function readSheetFromExcel(sheetName) {
  const filePath = path.join(__dirname, "test.xlsx");

  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    throw new Error(`找不到工作表：${sheetName}`);
  }

  return XLSX.utils.sheet_to_json(sheet, {
    defval: "",
  });
}

// 統一取得員編
function getId(person) {
  return String(person.員編 || "").trim();
}

// 產生車位輪序
function generateParkingOrder({
  currentApplicants,
  lastOrderList,
  parkingCount,
}) {
  const getId = (person) => String(person.員編 || "").trim();

  const currentIds = new Set(currentApplicants.map(getId));
  const lastOrderIds = new Set(lastOrderList.map(getId));

  // 1. 上期有，這期也有：保留輪序
  const continuedList = lastOrderList
    .filter((person) => currentIds.has(getId(person)))
    .sort((a, b) => Number(a.順位) - Number(b.順位));

  // 2. 上期有，這期沒有：未申請
  const notAppliedThisSeason = lastOrderList.filter((person) => {
    return !currentIds.has(getId(person));
  });

  // 3. 上期沒有，這期有：第一次申請
  const firstTimeApplicants = currentApplicants.filter((person) => {
    return !lastOrderIds.has(getId(person));
  });

  // 4. 本期替補人數
  const waitingCount = Math.max(continuedList.length - parkingCount, 0);

  // 5. 輪序前面 N 人變替補
  const waiting = continuedList.slice(0, waitingCount).map((person, index) => ({
    ...person,
    本期狀態: "替補",
    本期替補順位: index + 1,
  }));

  // 6. 剩下的人變正選
  const assigned = continuedList.slice(waitingCount).map((person, index) => ({
    ...person,
    本期狀態: "正選",
    本期正選順位: index + 1,
  }));

  // 7. 第一次申請者，排到下一期輪序最後
  const nextSeasonNewApplicants = firstTimeApplicants.map((person, index) => ({
    ...person,
    本期狀態: "第一次申請",
    備註: "本期不參與，排入下期輪序最後",
    下期新增順位: index + 1,
  }));

  // 8. 下一期輪序：正選在前，替補移到後面，第一次申請者再放最後
  const nextSeasonOrder = [
    ...assigned,
    ...waiting,
    ...nextSeasonNewApplicants,
  ].map((person, index) => ({
    ...person,
    下期順位: index + 1,
  }));

  return {
    summary: {
      本期申請人數: currentApplicants.length,
      上期輪序人數: lastOrderList.length,
      本期有效輪序人數: continuedList.length,
      停車位數: parkingCount,
      本期正選人數: assigned.length,
      本期替補人數: waiting.length,
      第一次申請人數: firstTimeApplicants.length,
      本期未申請人數: notAppliedThisSeason.length,
    },

    本期正選名單: assigned,
    本期替補名單: waiting,
    第一次申請名單: firstTimeApplicants,
    本期未申請名單: notAppliedThisSeason,
    下期輪序名單: nextSeasonOrder,
  };
}

app.get("/api/a", (req, res) => {
  try {
    const currentApplicants = readSheetFromExcel("申請資料");
    const lastOrderList = readSheetFromExcel("車位資料");

    const result = generateParkingOrder({
      currentApplicants,
      lastOrderList,
      parkingCount: 5,
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
