<template>
  <div class="sheet">
    <div class="operation" :style="`max-width: ${width + 2}px`">
      <img
        class="mr15"
        src="../assets/images/clear_border.png"
        @click="cancelBorder"
        title="取消边框"
      />
      <img
        class="mr15"
        src="../assets/images/top_bottom.png"
        @click="addTBBorder"
        title="上下边框"
      />
      <img
        class="mr15"
        src="../assets/images/top_right_bottom.png"
        @click="addTBRBorder"
        title="上下右边框"
      />
      <img
        class="mr15"
        src="../assets/images/fill_main.png"
        @click="toMainColor"
        title="设为主栏"
      />
      <img
        class="mr15"
        src="../assets/images/fill_guest.png"
        @click="toGuestColor"
        title="设为宾栏"
      />
      <img
        class="mr15"
        src="../assets/images/fill.png"
        @click="cancelColor"
        title="取消填色"
      />
    </div>
    <div class="ova wrapper-content">
      <div ref="wrapper" class="canvas-wrapper" :style="wrapperStyle">
        <canvas id="canvas" ref="canvas"></canvas>
        <div class="left-row" :style="leftRowStyle">
          <div
            v-for="(row, rowIndex) in excelGrid.row"
            :key="rowIndex"
            class="left-item"
            :style="`height: ${defaultHeight}px`"
            @click="selectRowRange(rowIndex)"
          ></div>
        </div>
        <div class="top-column clearfix" :style="topColumnStyle">
          <div
            v-for="(col, colIndex) in excelGrid.column"
            :key="colIndex"
            class="top-item"
            :style="`width: ${col}px;height: ${defaultHeight}px`"
            @click="selectColRange(colIndex)"
          ></div>
        </div>
        <div
          class="edit-wrapper"
          :style="contentStyle"
          @dblclick="openEdit"
          @mousedown="startSelectRange"
          @mouseup="stopSelectRange"
          @mousemove="changeSelectRange"
        >
          <div class="active" :style="activeStyle"></div>
        </div>
        <div :style="inputStyle" class="text-wrapper">
          <textarea
            ref="textarea"
            rows="1"
            v-show="showTextArea"
            v-model="content"
            :style="textAreaStyle"
            @keydown.enter="stopDefault"
            @keyup.enter="valueChange"
            @blur="valueChange"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script>
const lineColor = '#777';
const fillColor = '#e8e8e8';
const borderColor = '#000';
const bgColor = '#fff';

const mainColor = '#99CCFF';
const guestColor = '#FFFF99';

const defaultWidth = 40;
const defaultHeight = 20;

const defaultFont = '14px serif';

const letterArray = [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'K',
  'L',
  'M',
  'N',
  'O',
  'P',
  'Q',
  'R',
  'S',
  'T',
  'U',
  'V',
  'W',
  'X',
  'Y',
  'Z',
];

const hidden = 'display: none';

export default {
  name: 'Sheet',

  props: {
    sheetData: {
      default() {
        return [];
      },
      type: Array,
    },
  },

  data() {
    return {
      dataSource: [],
      firstWidth: defaultWidth,
      defaultHeight,
      width: 300,
      height: 100,
      content: '',
      showTextArea: false,
      startRowIndex: -1,
      startColIndex: -1,
      endRowIndex: -1,
      endColIndex: -1,
      moving: false,
      cacheRowIndex: -1,
      cacheColIndex: -1,
      editPosition: [],
    };
  },

  computed: {
    excelGrid() {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      const row = [];
      const column = [];
      const data = [];
      this.dataSource.forEach((child, childIndex) => {
        row.push(defaultHeight);
        child.forEach((item, index) => {
          ctx.font = '12px 宋体';
          const { width } = ctx.measureText(item.value);
          const colWidth = width + 10 < defaultWidth ? defaultWidth : width + 10;
          if (!column[index]) {
            column.push(colWidth);
          } else if (colWidth > column[index]) {
            column[index] = colWidth;
          }
          data.push({
            ...item,
            rowIndex: childIndex,
            colIndex: index,
          });
        });
      });
      return {
        row,
        column,
        data,
      };
    },
    rowsHeight() {
      return defaultHeight * this.excelGrid.row.length;
    },
    columnsWidth() {
      if (!this.excelGrid.column.length) return this.firstWidth;
      const colWidth = this.excelGrid.column.reduce((a, b) => a + b);
      return colWidth + this.firstWidth;
    },
    wrapperStyle() {
      return `width: ${this.width + 2}px;height: ${this.height + 2}px;`;
    },
    leftRowStyle() {
      return `width: ${this.firstWidth}px;height: ${this.rowsHeight}px`;
    },
    topColumnStyle() {
      return `width: ${this.columnsWidth - this.firstWidth}px;height: ${defaultHeight}px;left: ${
        this.firstWidth
      }px;`;
    },
    rowsRange() {
      const ranges = [];
      this.excelGrid.row.forEach((item, index, arr) => {
        if (!index) {
          ranges.push(defaultHeight);
          return;
        }
        ranges.push(ranges[index - 1] + arr[index - 1]);
      });
      return ranges;
    },
    columnsRange() {
      const ranges = [];
      this.excelGrid.column.forEach((item, index, arr) => {
        if (!index) {
          ranges.push(this.firstWidth);
          return;
        }
        ranges.push(ranges[index - 1] + arr[index - 1]);
      });
      return ranges;
    },
    contentStyle() {
      return `width: ${this.columnsWidth - this.firstWidth}px;
      height: ${this.rowsHeight}px;
      left: ${this.firstWidth}px;
      top: ${defaultHeight}px;`;
    },
    activeStyle() {
      const {
        startRowIndex, startColIndex, endRowIndex, endColIndex,
      } = this;
      if (
        startRowIndex < 0
        || startColIndex < 0
        || endRowIndex < 0
        || endColIndex < 0
        || !this.excelGrid.column.length
        || !this.excelGrid.row.length
      ) return hidden;
      const width = this.excelGrid.column
        .slice(startColIndex, endColIndex + 1)
        .reduce((a, b) => a + b);
      const height = this.excelGrid.row
        .slice(startRowIndex, endRowIndex + 1)
        .reduce((a, b) => a + b);
      return `width: ${width - 1}px;
      height: ${height - 1}px;
      top: ${this.rowsRange[startRowIndex] - defaultHeight + 1}px;
      left: ${this.columnsRange[startColIndex] - this.firstWidth + 1}px`;
    },
    inputStyle() {
      const [rowIndex, colIndex] = this.editPosition;
      if (!this.showTextArea) return hidden;
      const width = this.excelGrid.column[colIndex];
      const height = this.excelGrid.row[rowIndex];
      const x = this.rowsRange[rowIndex];
      const y = this.columnsRange[colIndex];
      return `width: ${width}px;
      height: ${height}px;
      left: ${y}px;
      top: ${x}px;`;
    },
    textAreaStyle() {
      const [rowIndex, colIndex] = this.editPosition;
      if (rowIndex === undefined || colIndex === undefined) return hidden;
      const cell = this.excelGrid.data.find(
        d => d.rowIndex === rowIndex && d.colIndex === colIndex,
      );
      if (!cell) return hidden;
      return `background: ${this.getCellFillStyle(cell)}`;
    },
  },

  watch: {
    excelGrid(grid) {
      this.init(grid);
    },
    showTextArea(value) {
      if (value) {
        this.$nextTick(() => {
          this.$refs.textarea.focus();
        });
      }
    },
  },

  methods: {
    init(grid) {
      const { canvas } = this.$refs;
      if (!canvas) return;
      if (!canvas.getContext) return;
      const ctx = canvas.getContext('2d');
      canvas.width = this.columnsWidth;
      canvas.height = this.rowsHeight + defaultHeight;
      this.width = canvas.width + 2;
      this.height = canvas.height + 2;

      this.initRows(ctx, grid.row, canvas.width);

      this.initColumns(ctx, grid.column, canvas.height);

      this.renderContent(ctx, grid.data);
    },
    initRows(ctx, row) {
      row.forEach((rowItem, rowIndex) => {
        const number = rowIndex + 1;
        const itemHeight = number * defaultHeight;
        this.drawLine(ctx, 0, defaultWidth, itemHeight, itemHeight, lineColor);
        ctx.font = defaultFont;
        let { width } = ctx.measureText(number);
        if (width < defaultWidth) width = defaultWidth;
        if (this.firstWidth > width) this.firstWidth = width;
        ctx.fillStyle = fillColor;
        ctx.fillRect(0, itemHeight, width, defaultHeight);
        this.drawText(ctx, number, defaultHeight, defaultHeight * number + defaultHeight / 2, {
          font: defaultFont,
          fillStyle: lineColor,
        });
      });
    },
    initColumns(ctx, columns) {
      let x = this.firstWidth;
      columns.forEach((width, colIndex) => {
        this.drawLine(ctx, x, x, 0, defaultHeight, lineColor);
        ctx.fillStyle = fillColor;
        ctx.fillRect(x, 0, width, defaultHeight);
        ctx.fillText(this.computedColumnText(colIndex), x + width / 2, defaultHeight / 2);
        this.drawText(ctx, this.computedColumnText(colIndex), x + width / 2, defaultHeight / 2, {
          font: defaultFont,
          fillStyle: lineColor,
        });
        x += width;
      });
    },
    computedColumnText(index, value = '') {
      while (index > 25) {
        return this.computedColumnText(index - 26, `${value}${letterArray[index % 26]}`);
      }
      return `${value}${letterArray[index % 26]}`;
    },
    getCellFillStyle({ mainColumn, guestColumn }) {
      if (mainColumn) return mainColor;
      if (guestColumn) return guestColor;
      return bgColor;
    },
    getCellBorderStyle(boolean) {
      return boolean ? borderColor : 'transparent';
    },
    getCellX({ colIndex }) {
      const list = this.excelGrid.column.filter((c, ci) => ci < colIndex);
      if (!list.length) return this.firstWidth;
      return list.reduce((a, b) => a + b) + this.firstWidth;
    },
    getTextCellX({ colIndex }) {
      return this.getCellX({ colIndex });
    },
    getTextCellY({ rowIndex }) {
      return this.getCellY({ rowIndex }) + defaultHeight / 2;
    },
    getCellY({ rowIndex }) {
      return defaultHeight * (rowIndex + 1);
    },
    getCellWidth({ colIndex }) {
      return this.excelGrid.column[colIndex];
    },
    getCellHeight() {
      return defaultHeight;
    },
    drawText(ctx, text = '', x, y, options) {
      ctx.textBaseline = 'middle';
      ctx.textAlign = 'center';
      Object.entries(options).forEach(([key, value]) => {
        ctx[key] = value;
      });
      ctx.fillText(text, x, y);
    },
    drawLine(ctx, x0, x1, y0, y1, color, offset = 0) {
      ctx.lineWidth = 1;
      ctx.strokeStyle = color;
      ctx.beginPath();
      ctx.moveTo(x0 - offset, y0 - offset); // 设置0.5偏移量以处理线条宽度问题
      ctx.lineTo(x1 - offset, y1 - offset);
      ctx.closePath();
      ctx.stroke();
    },
    drawBorder(ctx, child) {
      const x = this.getCellX(child);
      const xWidth = this.getCellX(child) + this.getCellWidth(child);
      const y = this.getCellY(child);
      const yHeight = this.getCellY(child) + defaultHeight;
      const {
        topLine,
        rightLine,
        buttonLine,
        leftLine,
        value,
        mainColumn,
        guestColumn,
        rowIndex,
        colIndex,
      } = child;
      // 画上线
      if (rowIndex) {
        this.drawLine(ctx, x, xWidth, y, y, this.getCellBorderStyle(topLine), 0.5);
      } else {
        this.drawLine(ctx, x, xWidth, y, y, lineColor, 0.5);
      }
      // 画左线
      if (colIndex) {
        this.drawLine(ctx, x, x, y, yHeight, this.getCellBorderStyle(leftLine), 0.5);
      } else {
        this.drawLine(ctx, x, x, y, yHeight, lineColor, 0.5);
      }
      const isValue = !topLine
        && !rightLine
        && !buttonLine
        && !leftLine
        && value !== ''
        && !guestColumn
        && !mainColumn;
      if ((isValue && rowIndex !== this.excelGrid.row.length - 1) || !rowIndex) {
        // 首尾行不画下线
        this.drawLine(ctx, x, xWidth, yHeight, yHeight, lineColor, 0.5);
      } else {
        this.drawLine(ctx, x, xWidth, yHeight, yHeight, this.getCellBorderStyle(buttonLine), 0.5);
      }
      if (
        (isValue && colIndex && colIndex !== this.excelGrid.column.length - 1)
        || (!isValue && topLine && !rightLine && buttonLine && !mainColumn && !guestColumn)
      ) {
        this.drawLine(ctx, xWidth, xWidth, y, yHeight, lineColor, 0.5);
      } else {
        this.drawLine(ctx, xWidth, xWidth, y, yHeight, this.getCellBorderStyle(rightLine), 0.5);
      }
    },
    renderContent(ctx, data) {
      data.forEach((child) => {
        ctx.fillStyle = this.getCellFillStyle(child);
        ctx.fillRect(
          this.getCellX(child),
          this.getCellY(child),
          this.getCellWidth(child),
          this.getCellHeight(child),
        );
        this.drawText(ctx, child.value, this.getTextCellX(child) + 5, this.getTextCellY(child), {
          font: '12px 宋体',
          fillStyle: '#000',
          textAlign: 'left',
        });
        this.drawBorder(ctx, child);
      });
    },
    getPosition(x, y) {
      const { x: parentX, y: parentY } = this.$refs.wrapper.getBoundingClientRect();
      const selfX = x - parentX;
      const colIndex = this.columnsRange.findIndex(
        (r, ri) => r + this.excelGrid.column[ri] >= selfX,
      );
      const selfY = y - parentY;
      const rowIndex = this.rowsRange.findIndex((c, ci) => c + this.excelGrid.row[ci] >= selfY);
      return {
        rowIndex,
        colIndex,
      };
    },
    startSelectRange(e) {
      const { clientX, clientY } = e;
      const { rowIndex, colIndex } = this.getPosition(clientX, clientY);
      this.startRowIndex = rowIndex;
      this.endRowIndex = rowIndex;
      this.startColIndex = colIndex;
      this.endColIndex = colIndex;
      // 缓存起始点
      this.cacheRowIndex = rowIndex;
      this.cacheColIndex = colIndex;
      this.moving = true;
    },
    stopSelectRange() {
      this.moving = false;
    },
    changeSelectRange({ clientX, clientY }) {
      if (!this.moving) return;
      const { rowIndex, colIndex } = this.getPosition(clientX, clientY);
      if (rowIndex < this.cacheRowIndex) {
        this.startRowIndex = rowIndex;
        this.endRowIndex = this.cacheRowIndex;
      } else {
        this.endRowIndex = rowIndex;
      }
      if (colIndex < this.cacheColIndex) {
        this.startColIndex = colIndex;
        this.endColIndex = this.cacheColIndex;
      } else {
        this.endColIndex = colIndex;
      }
    },
    selectRowRange(index) {
      this.startColIndex = 0;
      this.endColIndex = this.columnsRange.length - 1;
      this.startRowIndex = index;
      this.endRowIndex = index;
    },
    selectColRange(index) {
      this.startRowIndex = 0;
      this.endRowIndex = this.rowsRange.length - 1;
      this.startColIndex = index;
      this.endColIndex = index;
    },
    openEdit({ clientX, clientY }) {
      const { rowIndex, colIndex } = this.getPosition(clientX, clientY);
      const cell = this.excelGrid.data.find(
        d => d.rowIndex === rowIndex && d.colIndex === colIndex,
      );
      if (!cell) return;
      this.showTextArea = true;
      this.content = cell.value;
      this.editPosition = [rowIndex, colIndex];
    },
    valueChange(e) {
      this.changeDataSource({ value: e.target.value }, this.editPosition);
      this.showTextArea = false;
      this.editPosition = [];
    },
    toMainColor() {
      this.changeSelectedValue({ mainColumn: true, guestColumn: false });
    },
    toGuestColor() {
      this.changeSelectedValue({ guestColumn: true, mainColumn: false });
    },
    cancelColor() {
      this.changeSelectedValue({ mainColumn: false, guestColumn: false });
    },
    addTBRBorder() {
      this.changeSelectedValue({
        rightLine: true,
        topLine: true,
        buttonLine: true,
      });
    },
    addTBBorder() {
      this.changeSelectedValue({
        topLine: true,
        buttonLine: true,
      });
    },
    cancelBorder() {
      this.changeSelectedValue({
        topLine: false,
        rightLine: false,
        buttonLine: false,
        leftLine: false,
      });
    },
    changeSelectedValue(map) {
      for (let i = this.startRowIndex; i <= this.endRowIndex; i += 1) {
        for (let j = this.startColIndex; j <= this.endColIndex; j += 1) {
          this.changeDataSource(map, [i, j]);
        }
      }
    },
    changeDataSource(keyValues, position) {
      const [rowIndex, colIndex] = position;
      this.dataSource.forEach((child, childIndex) => {
        if (rowIndex !== childIndex) return child;
        return child.map((item, index) => {
          if (index === colIndex) {
            Object.entries(keyValues).forEach(([key, value]) => {
              item[key] = value;
            });
          }
          return item;
        });
      });
    },
    stopDefault(e) {
      e.preventDefault();
    },
  },

  mounted() {
    this.dataSource = JSON.parse(JSON.stringify(this.sheetData));
  },
};
</script>

<style scoped lang="less">
.sheet {
  width: 100%;
  overflow: auto;
}
.operation {
  margin: 0 auto 15px;
  img {
    cursor: pointer;
  }
}
.canvas-wrapper {
  position: relative;
  margin: 0 auto;
}
#canvas {
  border: 1px solid #000;
}
.left-row {
  position: absolute;
  left: 0;
  top: 20px;
  .left-item {
    width: 100%;
  }
}
.top-column {
  position: absolute;
  top: 0;
  .top-item {
    float: left;
  }
}
.edit-wrapper {
  position: absolute;
  cursor: cell;
}
.text-wrapper {
  position: absolute;
  overflow: hidden;
  textarea {
    width: 100%;
    height: 100%;
    border: none;
    outline: none;
    overflow: hidden;
    resize: none;
    padding: 0 5px;
    font-size: 12px;
    word-break: keep-all;
    white-space: nowrap;
  }
}
.wrapper-content {
  max-height: calc(100% - 39px);
}
.active {
  position: absolute;
  box-shadow: rgb(255, 255, 255) 0 0 0 1px, rgb(31, 187, 125) 0 0 0 3px;
  background: rgba(31, 187, 125, 0.1);
}
</style>
