/* global document, Excel, Office */

let gameArea = { top: 0, left: 0, rows: 0, cols: 0 };
let player = { x: 0, y: 0 };
let bullets = [];
let enemies = [];
let gameLoopId = null;
let isPaused = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = initGame;
    document.getElementById("pause").onclick = togglePause;
    document.getElementById("restart").onclick = restartGame;
  }
});

async function initGame() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();

    range.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();

    gameArea.top = range.rowIndex;
    gameArea.left = range.columnIndex;
    gameArea.rows = range.rowCount;
    gameArea.cols = range.columnCount;

    range.values = Array(gameArea.rows).fill().map(() => Array(gameArea.cols).fill(""));

    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(edge => {
      range.format.borders.getItem(edge).style = "Continuous";
      range.format.borders.getItem(edge).weight = "Thick";
    });

    player.x = Math.floor(gameArea.cols / 2);
    player.y = gameArea.rows - 1;

    await context.sync();
  });

  bullets = [];
  enemies = [];
  isPaused = false;
  document.getElementById("status-label").innerText = "Game running...";

  document.addEventListener("keydown", handleKey);

  if (gameLoopId) clearInterval(gameLoopId);
  gameLoopId = setInterval(() => {
    if (!isPaused) gameLoop();
  }, 400);
}

function togglePause() {
  isPaused = !isPaused;
  document.getElementById("status-label").innerText = isPaused ? "Game paused" : "Game running...";
}

async function restartGame() {
  if (gameLoopId) clearInterval(gameLoopId);

  // 清空当前游戏区域
  await Excel.run(async (context) => {
    if (gameArea.rows > 0 && gameArea.cols > 0) {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRangeByIndexes(gameArea.top, gameArea.left, gameArea.rows, gameArea.cols);
      range.values = Array(gameArea.rows).fill().map(() => Array(gameArea.cols).fill(""));
    }
    await context.sync();
  });

  // 重新开始游戏
  initGame();
}

function handleKey(e) {
  if (isPaused) return; // 暂停时不响应操作

  if (e.key === "ArrowLeft" && player.x > 0) player.x--;
  if (e.key === "ArrowRight" && player.x < gameArea.cols - 1) player.x++;
  if (e.key === "ArrowUp" && player.y > 0) player.y--;
  if (e.key === "ArrowDown" && player.y < gameArea.rows - 1) player.y++;

  if (e.code === "Space") {
    bullets.push({ x: player.x, y: player.y - 1 });
  }
}

function gameLoop() {
  bullets.forEach((b) => b.y--);
  bullets = bullets.filter((b) => b.y >= 0);

  enemies.forEach((e) => e.y++);
  enemies = enemies.filter((e) => e.y < gameArea.rows);

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

  if (Math.random() < 0.3) {
    enemies.push({ x: Math.floor(Math.random() * gameArea.cols), y: 0 });
  }

  render();
}

async function render() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRangeByIndexes(gameArea.top, gameArea.left, gameArea.rows, gameArea.cols);

    range.values = Array(gameArea.rows).fill().map(() => Array(gameArea.cols).fill(""));

    sheet.getCell(gameArea.top + player.y, gameArea.left + player.x).values = [["✈"]];

    bullets.forEach((b) => {
      if (b.y >= 0) {
        sheet.getCell(gameArea.top + b.y, gameArea.left + b.x).values = [["|"]];
      }
    });

    enemies.forEach((e) => {
      if (e.y < gameArea.rows) {
        sheet.getCell(gameArea.top + e.y, gameArea.left + e.x).values = [["●"]];
      }
    });

    await context.sync();
  });
}
