const debug = require('debug')('resource')

const unstyledfile = 'master.xlsx'
const styledfilena = 'master.xlsx'
const sheetname = 'hours'

const days = ['Mon', 'Tues', 'Wed', 'Thurs', 'Fri', 'Sat', 'Sun']
const headers = ['name', 'wk', 'date', 'class', ...days, 'total']
const colorrows = ['fff5f5fa', 'ffe5e5e5']
const colorhead = 'ffafafaf'
const daywidth = 7.75
const widths = [0, 20, 7, 9, 9, ...days.map(d => daywidth), 9]

const fix = input => {
  if (input === undefined) {
    return 0
  } else return input
}
const sortentry = (a, b) => {
  if (!a[headers[1]]) {
    debug(a)
    console.log(`Warning: missing week number: ${a[headers[0]]}`)
    a[headers[1]] = -1
  }
  if (!b[headers[1]]) {
    debug(b)
    console.log(`Warning: missing week number: ${b[headers[0]]}`)
    b[headers[1]] = -1
  }
  if (!a[headers[3]]) {
    console.log(`Warning: missing lunch/work label: ${a[headers[0]]}, week: ${a[headers[1]]}`)
    debug(a)
    a[headers[3]] = ' '
  }
  if (!b[headers[3]]) {
    console.log(`Warning: missing lunch/work label: ${b[headers[0]]}, week: ${b[headers[1]]}`)
    debug(b)
    b[headers[3]] = ' '
  }
  if (a[headers[1]] < b[headers[1]]) return -1
  if (a[headers[1]] > b[headers[1]]) return 1
  if (a[headers[0]].split(' ')[1].toLowerCase() < b[headers[0]].split(' ')[1].toLowerCase())
    return -1
  if (a[headers[0]].split(' ')[1].toLowerCase() > b[headers[0]].split(' ')[1].toLowerCase())
    return 1
  // if (a[headers[3]].toLowerCase() < b[headers[3]].toLowerCase()) return -1
  // if (a[headers[3]].toLowerCase() > b[headers[3]].toLowerCase()) return 1
  return 0
}
const fixrow = (old, name) => {
  let arr = Object.assign({}, old)
  arr['total'] = 0
  arr['name'] = name
  days.forEach(day => {
    arr[day] = fix(arr[day])
    arr['total'] += arr[day]
  })
  return arr
}
const stylehead = cell => {
  const text = cell.value
  cell.value = text.toUpperCase()
  cell.font = { size: 12, bold: true }
  cell.border = { bottom: { style: 'thick' } }
  cell.fill = {
    type: 'gradient',
    gradient: 'angle',
    degree: 0,
    stops: [
      { position: 0, color: { argb: colorhead } },
      { position: 1, color: { argb: colorhead } },
    ],
  }
  cell.alignment = { vertical: 'bottom', horizontal: 'center' }
}
const stylerows = (cell, rown) => {
  if (cell.value === 0) {
    cell.font = { color: { argb: colorrows[rown % 2] } }
  }
  cell.border = {
    left: { style: 'thin', color: { argb: colorrows[(rown + 1) % 2] } },
  }
  cell.fill = {
    type: 'gradient',
    gradient: 'angle',
    degree: 0,
    stops: [
      { position: 0, color: { argb: colorrows[rown % 2] } },
      { position: 1, color: { argb: colorrows[rown % 2] } },
    ],
  }
}

module.exports = {
  days,
  headers: () => {
    return headers.map(h => h)
  },
  sortentry,
  fixrow,
  stylehead,
  stylerows,
  unstyledfile,
  styledfilena,
  sheetname,
  widths,
  cleanpath: path => {
    return path.replace(/\%20/g, ' ')
  },
}
