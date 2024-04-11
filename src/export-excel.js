import XLSX from 'xlsx'

// 支持列分组的数据导出，但是不支持数据部分的合并单元格
export default class ExcelExporter {
  constructor(options) {
    this.options = options
    this.worksheetNameList = []
  }

  createWorkbook () {
    this.wb = XLSX.utils.book_new()
    this.options.worksheets.forEach(ws => {
      this.addWorksheet(ws)
    })
  }

  addWorksheet(wsData) {
    // current row index, draw pointer
    this.index = 0
    // store column id and dataType
    this.columnList = []

    // format worksheet - merge cells...
    // 处理数据，比如 合并单元格
    /** data Object: - wsData.data
     * @property rows
     * @property columns
     * @property options
     * @property groupNames
     */
    this.mergeCells = []
    this.header = []
    this.headerDisplay = []
    this.createColumns(wsData.data)

    this.rows = []
    this.createRows(wsData.data)
    const ws = XLSX.utils.json_to_sheet([... this.headerDisplay, ... this.rows], { header: this.header, skipHeader: true })
    ws['!merges'] = this.mergeCells

    // Add a Worksheet to Workbook
    // 将 worksheet 放入 workbook 中，这里必须已经将 worksheet 的数据处理完了
    /** config Object: - wsData.config
     * @property name @type String
     */
    XLSX.utils.book_append_sheet(this.wb, ws, this.getWorksheetName(wsData.config))
  }

  getWorksheetName(config) {
    const wsName = config.name
    // check if worksheet name is duplicated
    const index = this.worksheetNameList.findIndex(item => item.name === wsName)
    if (index !== -1) {
      const worksheetNameCount = this.worksheetNameList[index].count ++
      return `${wsName} ${worksheetNameCount}`
    }

    this.worksheetNameList.push({
      name: wsName,
      count: 1
    })
    return wsName
  }

  addGroupNameInCols (columns, groupNames) {
    if (!groupNames || !groupNames.length) {
      return columns
    }
    groupNames.forEach((groupName, index) => {
      const groupId = `group_level_${index + 1}`
      const group = {
        name: groupName,
        id: groupId
      }
      columns.splice(index, 0, group)
    })
    return columns
  }

  createColumns (data) {
    const columns = this.addGroupNameInCols(data.columns, data.groupNames)
    let root = {
      children: columns,
      height: 0
    }
    root = this.prehandleColumns(root)
    // column height, width
    this.columnHeight = root.height - 1
    this.columnWidth = root.width

    this.handleColumns(root, this.index, 0, 0)
    const curRowNo = this.index + this.columnHeight
    this.index = curRowNo + 1
  }

  createRows(data) {
    const root = { children: data.rows }
    this.handleRows(root, 0)
  }

  getColumnDataType(column) {
    if (typeof column.customRender === 'function') {
      return 'custom'
    }
    return column.dataType || 'string'
  }

  prehandleColumns(parent) {
    const width = 1
    const height = 1
    let subWidth = 0
    let subHeight = 0
    const subs = parent.children || []
    if (!subs.length) {
      this.columnList.push({
        id: parent.prop || parent.dataIndex,
        dataType: this.getColumnDataType(parent),
        customFunction: parent.customRender
      })
    }
    subs.forEach(item => {
      if (!(item.prop || item.dataIndex || item.children?.length) || item.unexportable) {
        return;
      }
      item = this.prehandleColumns(item)
      subHeight = Math.max(subHeight, item.height)
      subWidth += item.width
    })

    parent.height = height + subHeight
    parent.width = Math.max(width, subWidth)
    return parent
  }

  // 这里处理后需要知道 每一列的开始列序号、结束列序号、开始行序号、结束行序号
  handleColumns(parent, startRow, startColumn, level) {
    const { children: subs } = parent
    const width = parent.width - 1
    const height = this.columnHeight - level - parent.height + 1
    const endRow = parent.children?.length > 0 ? startRow : startRow + height
    const endCol = startColumn + width
    
    if (level) {
      // isRoot = level === 0
      if (startRow !== endRow || startColumn !== endCol) {
        // 当发现需要合并单元格的时候，记录合并的行列信息
        this.mergeCells.push({
          s: {
            r: startRow - 1,
            c: startColumn
          },
          e: {
            r: endRow - 1,
            c: endCol
          }
        })
      }
      this.genHeaderCell(parent, startRow, startColumn, endRow, endCol)
    }
    if (!subs || !subs.length) {
      return
    }

    let subsColumn = startColumn
    level ++
    subs.forEach(item => {
      this.handleColumns(item, endRow + 1, subsColumn, level)
      subsColumn += item.width
    })
  }

  genHeaderCell(column, startRow, startColumn, endRow, endColumn) {
    if (column.unexportable) {
      return
    }

    const text = column.name || column.label || column.title

    const displayColumn = startRow > this.headerDisplay.length ? {} : {
      ... this.headerDisplay[startRow - 1]
    }

    if (column.children || !(column.prop || column.dataIndex)) {
      // 当列为 分组名 时，定位到具体的 header 的 id
      const cacheCol = this.columnList[startColumn]
      if (cacheCol?.id) {
        displayColumn[cacheCol.id] = text
      }
    } else {
      displayColumn[column.prop || column.dataIndex] = text

      this.header.push(column.prop || column.dataIndex)
    }

    if (startRow > this.headerDisplay.length) {
      this.headerDisplay.push(displayColumn)
    } else {
      this.headerDisplay[startRow - 1] = displayColumn
    }
  }

  handleRows(parent, level) {
    const subs = parent.children || []
    if (level) {
      this.genRow(parent, level)
    }
    subs.forEach(item => {
      if (item) {
        this.handleRows(item, level + 1)
      }
    })
  }

  formatValue(value) {
    if (value === undefined || value === null || value.toString() === "NaN") {
      return "\u2013"
    }
    return value
  }

  genRow(row, group) {
    if (row.rowType === 'blank') {
      return
    }
    const groupId = `group_level_${group}`
    const newRow = {}
    this.columnList.forEach(item => {
      const { id } = item
      let value = row[id]
      if (id === groupId) {
        // For current group level, Name column's value shown in group column, and Name column will show nothing
        value = row.name
        row.name = ''
      } else if (id.indexOf('group_level_') === 0 || row.rowType === 'blank') {
        // For other group levels, the value should be empty
        value = ''
      } else {
        // For non group column, keep original logic
        value = this.formatValue(value)
      }

      const { dataType, customFunction } = item
      let cellValue = ''
      if (dataType === 'custom' && typeof customFunction === 'function') {
        cellValue = customFunction({ text: value, record: row })
      } else if (dataType === 'number') {
        cellValue = value
      } else if (value) {
        cellValue = value
      }
      newRow[id] = cellValue
    })

    this.rows.push(newRow)
  }

  exportExcel (fileName) {
    this.createWorkbook()
    XLSX.writeFile(this.wb, fileName)
  }
}
