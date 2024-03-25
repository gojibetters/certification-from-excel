import xlsx from "xlsx"
import jimp from "jimp"
import moment from "moment"

const fontNormal32 = await jimp.loadFont('fonts/tahoma32.fnt')
const fontNormal64 = await jimp.loadFont('fonts/tahoma64.fnt')
const fontBold64 = await jimp.loadFont('fonts/tahomabd64.fnt')
const formatDate = (date) => moment(new Date(date)).format("DD/MM/YYYY")

const table = xlsx.readFile('courseClass.xlsx', {
  cellDates: true
})
const ws = table.Sheets['Sheet1']
const jsonSheet = xlsx.utils.sheet_to_json(ws)

jsonSheet.forEach(async (row) => {

  const image = await jimp.read('images/certificate.jpg')

  // Get the information
  const name = row['Name'],
    courseName = row['Course Name'],
    workload = row['Workload'] + ' horas',
    beginDate = formatDate(row['Begin Date']),
    finishDate = formatDate(row['Finish Date'])

  const file = `images/certification - ${name}.jpg`

  // Set information on image
  image.print(fontBold64, 590, 475, name)
  image.print(fontNormal64, 625, 543, courseName)
  image.print(fontNormal64, 830, 610, "Aluno")
  image.print(fontNormal64, 860, 675, workload)
  image.print(fontNormal32, 430, 1020, beginDate)
  image.print(fontNormal32, 430, 1105, finishDate)
  image.print(fontNormal32, 1265, 1105, finishDate)

  // Save
  image.write(file)
})