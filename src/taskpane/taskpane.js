/* global document, Excel, Office */

let gameArea = { top: 0, left: 0, rows: 0, cols: 0 };
let player = { x: 0, y: 0 };
let bullets = [];
let enemies = [];
let gameLoopId = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = initGame;
  }
});

async function initGame() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();

    // 必须 load
    range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();

    // 保存屏幕范围
    gameArea.top = range.rowIndex;
    gameArea.left = range.columnIndex;
    gameArea.rows = range.rowCount;
    gameArea.cols = range.columnCount;

    // 清空区域
    range.values = Array(gameArea.rows).fill().map(() => Array(gameArea.cols).fill(""));

    // 添加边框
    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(edge => {
      range.format.borders.getItem(edge).style = "Continuous";
      range.format.borders.getItem(edge).weight = "Thick";
    });

    // 飞机初始位置：底部中央
    player.x = Math.floor(gameArea.cols / 2);
    player.y = gameArea.rows - 1;

    await context.sync();
  });

  // 绑定键盘事件
  document.addEventListener("keydown", handleKey);

  // 启动游戏循环
  if (gameLoopId) clearInterval(gameLoopId);
  gameLoopId = setInterval(gameLoop, 400); // 每 400ms 刷新一次
}

function handleKey(e) {
  if (e.key === "ArrowLeft" && player.x > 0) player.x--;
  if (e.key === "ArrowRight" && player.x < gameArea.cols - 1) player.x++;
  if (e.key === "ArrowUp" && player.y > 0) player.y--;
  if (e.key === "ArrowDown" && player.y < gameArea.rows - 1) player.y++;

  // 空格键发射子弹
  if (e.code === "Space") {
    bullets.push({ x: player.x, y: player.y - 1 });
  }
}

function gameLoop() {
  // 更新子弹位置
  bullets.forEach((b) => b.y--);
  bullets = bullets.filter((b) => b.y >= 0);

  // 敌机下落
  enemies.forEach((e) => e.y++);
  enemies = enemies.filter((e) => e.y < gameArea.rows);

  // 碰撞检测
  bullets.forEach((b, bi) => {
    enemies.forEach((e, ei) => {
      if (b.x === e.x && b.y === e.y) {
        bullets[bi].hit = true;
        enemies[ei].hit = true;
      }
    });
  });
  bullets = bullets.filter((b) => !b.hit);
  enemies = enemies.filter((e) => !e.hit);

  // 随机生成敌机（概率）
  if (Math.random() < 0.3) {
    enemies.push({ x: Math.floor(Math.random() * gameArea.cols), y: 0 });
  }

  render();
}

async function render() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRangeByIndexes(gameArea.top, gameArea.left, gameArea.rows, gameArea.cols);

    // 清空屏幕
    range.values = Array(gameArea.rows).fill().map(() => Array(gameArea.cols).fill(""));

    // 画飞机
    sheet.getCell(gameArea.top + player.y, gameArea.left + player.x).values = [["✈"]];

    // 画子弹
    bullets.forEach((b) => {
      if (b.y >= 0) {
        sheet.getCell(gameArea.top + b.y, gameArea.left + b.x).values = [["|"]];
      }
    });

    // 画敌机
    enemies.forEach((e) => {
      if (e.y < gameArea.rows) {
        sheet.getCell(gameArea.top + e.y, gameArea.left + e.x).values = [["●"]];
      }
    });

    await context.sync();
  });
}
