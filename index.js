const res = require('./src/resource')
const debug = require('debug')('excel')
const moment = require('moment')
const XLSX = require('xlsx')
const { readdirSync } = require('fs')
const { pathToFileURL } = require('url')
const Excel = require('exceljs')

const settings = JSON.parse(fs.readFileSync('./settings.json', 'utf8'))
const dirpath = process.env.LOAD_PATH || settings.load
const savepath = process.env.SAVE_PATH || settings.save

// helper to just get subdirectorys
const getDirectories = source => {
  let dirlist = readdirSync(source, { withFileTypes: true })
  return dirlist.filter(f => f.isDirectory()).map(f => source.pathname.substr(1) + '/' + f.name)
}

// generating the main sheet
const Main = () => {
  moment.updateLocale('en', { week: { dow: 1 } })
  let currentData = []
  const path = pathToFileURL(dirpath)
  getDirectories(path).forEach(person => {
    readdirSync(pathToFileURL(res.cleanpath(person))).forEach(file => {
      // make sure we only work on XLSX files
      if (file.split('.')[1] === 'xlsx') {
        // get name from the dir not file
        const name = res
          .cleanpath(person)
          .split('/')
          .pop()
        // name from file
        const workbook = XLSX.readFile(pathToFileURL(res.cleanpath(person) + '/' + file))
        const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]])
        sheet.forEach(row => {
          const date = moment(row[res.headers()[1]], 'w')
          // show output on console
          console.log(`user: ${name}\tweek: ${date.format('YYYY.w (MMM Do)')}`)
          const finalrow = res.fixrow(row, name)
          currentData.push(finalrow)
        })
      }
    })
  })
  // sort by week then name
  currentData.sort(res.sortentry)
  // debug(currentData)
  let sheet = XLSX.utils.json_to_sheet(currentData, { header: res.headers() })
  let wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, sheet, res.sheetname)
  // write new excel file
  try {
    XLSX.writeFile(wb, savepath + '/' + res.unstyledfile)
  } catch (e) {
    console.log(`ERROR: file '${savepath + '\\' + res.unstyledfile}' can not be written to`)
    process.exit(1)
  }
}

// styleing the document
const Style = () => {
  let workbook = new Excel.Workbook()
  workbook.xlsx
    .readFile(savepath + '/' + res.unstyledfile)
    .then(() => {
      let sheet = workbook.getWorksheet(1)
      // set an autofilter
      sheet.autoFilter = {
        from: 'A1',
        to: {
          row: sheet.rowCount,
          column: sheet.columnCount
        }
      }

      sheet.getRow(1).height = 18
      sheet.getRow(1).eachCell((cell, col) => {
        sheet.getColumn(col).width = res.widths[col]
      })

      // set the Header Styles
      sheet.getRow(1).eachCell(cell => {
        res.stylehead(cell)
      })
      // center the week numbers
      sheet.getColumn(2).eachCell((c, r) => {
        if (r !== 1) {
          c.alignment = { horizontal: 'center' }
        }
      })
      // apply formulas to the totals column
      sheet.getColumn(sheet.columnCount).eachCell((c, r) => {
        if (r !== 1) {
          const tint = c.value
          c.value = { formula: `=SUM(E${r}:K${r})`, result: tint }
        }
      })
      // color and style each row of data
      sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber !== 1) {
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            // center the daily hours
            if (colNumber > 4 && colNumber < 12) {
              cell.alignment = { horizontal: 'center' }
            }
            res.stylerows(cell, rowNumber)
          })
        }
      })
      // write styled excel file
      workbook.xlsx
        .writeFile(savepath + '/' + res.styledfilena)
        .then(() => {
          console.log(`master sheet saved: ${savepath + '\\' + res.styledfilena}`)
          debug('saved')
        })
        .catch(err => {
          console.log(`Error could not save file`)
          debug('could not save')
          debug(err)
        })
    })
    .catch(err => {
      console.log(`Error could not save file`)
      debug('Fail!')
      debug(err)
    })
}

Main()
Style()
