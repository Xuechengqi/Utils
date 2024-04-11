import ExcelExporter from '../src/export-excel.js'

import TableData from './data/export-excel/table-data.json'
import tableColumns from './data/export-excel/table-columns.json'

const exportExcel = data => {
        try {
            const exporter = new ExcelExporter({
                worksheets: [
                    {
                        config: {
                            name: 'sheet 1'
                        },
                        data: {
                            rows: data,
                            columns: tableColumns
                        }
                    }
                ]
            })
            exporter.exportExcel("上海证券交易所程序化交易投资者信息数据监控表.xlsx")
        } catch (error) {
            console.log('excel export: ', error)
            throw error
        }
    }

exportExcel(TableData)
