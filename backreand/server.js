const express = require("express"); //建立伺服器
const cors = require("cors"); //不同來源的前端可以來呼叫你的 API。
const XLSX = require("xlsx"); //載入 xlsx 套件
const path = require("path"); //用來處理檔案路徑

const app = express();
app.use(cors()); //套用 cors 這個中介軟體
app.use(express.json()); //Express 內建的一個中介軟體，用來解析 JSON，全域套用 JSON 解析功能

//從資料表拿資料
function readUsersFromExcel() {
  const filePath = path.join(__dirname, "test.xlsx"); //path 模組裡的 join  ex: path.join("abc", "test.xlsx")=abc/test.xlsx   <__dirname = 當前目錄檔位置>
  console.log("讀取檔案位置:", filePath);

  const workbook = XLSX.readFile(filePath); //readFile = 讀取某個檔案
  console.log("所有工作表:", workbook.SheetNames);

  const firstSheetName = workbook.SheetNames[0]; //讀 workbook 物件裡的 SheetNames 屬性 [0] = 第一個工作表
  console.log("你是什麼?", firstSheetName);
  const sheet = workbook.Sheets[firstSheetName]; //從整本 Excel 裡，取出剛剛那張工作表。

  const data = XLSX.utils.sheet_to_json(sheet); // sheet_to_json(...)把工作表轉成 JSON。
  return data;
}

app.get("/api/a", (req, res) => {
  try {
    const aa = readUsersFromExcel();
    res.json(aa);
  } catch {
    console.log("無法連線");
  }
});

app.listen(3001, () => {
  console.log("Server running on http://localhost:3001/api/a");
});
