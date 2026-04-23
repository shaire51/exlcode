const express = require("express"); //建立伺服器
const cors = require("cors"); //不同來源的前端可以來呼叫你的 API。
const XLSX = require("xlsx"); //載入 xlsx 套件
const path = require("path"); //用來處理檔案路徑

const app = express();
app.use(cors()); //套用 cors 這個中介軟體
app.use(express.json()); //Express 內建的一個中介軟體，用來解析 JSON，全域套用 JSON 解析功能

function readUsersFromExcel() {
  const filePath = path.join(__dirname, "test.xlsx"); //path 模組裡的 join  ex: path.join("abc", "test.xlsx")=abc/test.xlsx   <__dirname = 當前目錄檔位置>
  console.log("讀取檔案位置:", filePath);

  const workbook = XLSX.readFile(filePath);
  console.log("所有工作表:", workbook.SheetNames);

  const firstSheetName = workbook.SheetNames[0];
  console.log("你是什麼?", firstSheetName);
  const sheet = workbook.Sheets[firstSheetName];
  console.log(sheet);

  const data = XLSX.utils.sheet_to_json(sheet);
  return data;
}

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
  console.log("Server running on http://localhost:3001/api/users");
});
