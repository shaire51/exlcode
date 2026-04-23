const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");
const path = require("path");

const app = express();
app.use(cors());
app.use(express.json());

// 讀 Excel 的函式
function readUsersFromExcel() {
  const filePath = path.join(__dirname, "users.xlsx");

  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets["Sheet1"];

  const data = XLSX.utils.sheet_to_json(sheet);
  return data;
}

// API：取得 Excel 資料
app.get("/api/users", (req, res) => {
  try {
    const users = readUsersFromExcel();
    res.json(users);
  } catch (error) {
    console.error("讀取 Excel 失敗:", error);
    res.status(500).json({ message: "讀取 Excel 失敗" });
  }
});

app.listen(3001, () => {
  console.log("Server running on http://localhost:3001");
});
