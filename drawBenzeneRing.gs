function drawPixel(sheet, x, y) {
  sheet.getRange(y, x).setBackground("black");
}

function drawLine(sheet, x1, y1, x2, y2) {
  var dx = Math.abs(x2 - x1);
  var dy = Math.abs(y2 - y1);
  var sx = x1 < x2 ? 1 : -1;
  var sy = y1 < y2 ? 1 : -1;
  var err = dx - dy;

  while (true) {
    drawPixel(sheet, x1, y1);

    if (x1 === x2 && y1 === y2) break;
    var e2 = 2 * err;
    if (e2 > -dy) {
      err -= dy;
      x1 += sx;
    }
    if (e2 < dx) {
      err += dx;
      y1 += sy;
    }
  }
}

function drawBenzeneRing() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
  // sheet.setName("BenzeneRing");

  // init the sheet

  const baseLineLen = 29;

  sheet.insertRowsAfter(1, 101);
  sheet.insertColumnsAfter(1, 101);

  sheet.setColumnWidths(2, 100, 15);
  sheet.setRowHeights(2, 100, 15);

  // draw a benzene ring
  const startposx = 2;
  const startposy = Math.ceil((baseLineLen * 2) / 3);

  // draw outside lines
  const rightlinex = startposx + Math.ceil(Math.sqrt(3) * baseLineLen);
  const topx = startposx + Math.ceil((Math.sqrt(3) * baseLineLen) / 2);
  const topy = startposy - Math.ceil(baseLineLen / 2);
  const bottomx = topx;
  const bottomy = baseLineLen + startposy + Math.ceil(baseLineLen / 2);
  drawLine(sheet, startposx, startposy, startposx, baseLineLen + startposy);
  drawLine(sheet, rightlinex, startposy, rightlinex, baseLineLen + startposy);
  drawLine(sheet, startposx, startposy, topx, topy);
  drawLine(sheet, rightlinex, startposy, topx, topy);
  drawLine(sheet, startposx, baseLineLen + startposy, bottomx, bottomy);
  drawLine(sheet, rightlinex, baseLineLen + startposy, bottomx, bottomy);

  // draw inside lines
  const margin = Math.round(baseLineLen / 8);
  const ix11 = startposx + Math.ceil((margin * Math.sqrt(3)) / 2);
  const iy11 = startposy + margin / 2;
  const ix12 = topx;
  const iy12 = topy + margin;
  const ix21 = ix11;
  const iy21 = baseLineLen + startposy - margin / 2;
  const ix22 = bottomx;
  const iy22 = bottomy - margin;
  const ix31 = rightlinex - Math.ceil((margin * Math.sqrt(3)) / 2);
  const iy31 = startposy + margin / 2;
  const ix32 = ix31;
  const iy32 = baseLineLen + startposy - margin / 2;
  drawLine(sheet, ix11, iy11, ix12, iy12);
  drawLine(sheet, ix21, iy21, ix22, iy22);
  drawLine(sheet, ix31, iy31, ix32, iy32);
}
