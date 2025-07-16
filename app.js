import { EMAIL } from "./mail";

const CONFIG = {
  emailSender: EMAIL,
  months: {8:'Aug', 9:'Sep', 10:'Okt', 11:'Nov', 0:'Dec', 1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'Maj', 6:'Jun',},
  sheetOrder: [8,9,10,11,0,1,2,3,4,5,6],
  scaleRange: 'A2:E3',
  absenceColors: {1:'#a4c2f4',2:'#3c78d8',3:'#b6d7a8',4:'#6aa84f',5:'#fff2cc',6:'#ffff00',7:'#f6b26b',8:'#ff9900',9:'#ff0000',10:'#990000'},
  absenceTotalCell: 'V1',
}


function monthlyAbsenceCheck() {
  const todaysMails = getTodaysMail(CONFIG.emailSender);

  todaysMails.forEach(mail => {
    const allAbsences = getAbsence(mail);
    const allSortedAbsences = sortByClass(allAbsences);
    const highAbsenceStudents = allSortedAbsences.filter(entry => entry.absence >= 15.0);
    
    let school = allSortedAbsences[0].year[0];
    if (school == 'y' || school == 'Y') {
      school = 'Ydre';
    } else {
      school = 'Hestra';
    }
    
    const d = new Date();
    const month = d.getMonth();
    const year = d.getFullYear();
    const schoolYear = getCurrentSchoolYear(month, year);

    const fileName = 'Frånvaro-' + school + '-' + schoolYear;
    const files = DriveApp.getFilesByName(fileName);
    let spreadsheet;

    if (!files.hasNext()){
      spreadsheet = createSheet(fileName);
      createSummaryTableHeader(spreadsheet);
      Logger.log('File created for instance ' + school);
    } else {
      const file = files.next();
      spreadsheet = SpreadsheetApp.open(file);
      Logger.log('File found');
    }
    if (CONFIG.sheetOrder.includes(month)){
      const previousTotalAbsence = getPreviousTotal(spreadsheet);
      let totalAbsence = getTotalAbsence(highAbsenceStudents, previousTotalAbsence, month);
      setTotalAbsence(totalAbsence, spreadsheet);
      createMonthsTableHeader(month, spreadsheet);
      createMonthsTable(month, spreadsheet, highAbsenceStudents, totalAbsence);

      updateSummary(spreadsheet, totalAbsence);
    }
    Logger.log('Finished operations for: ' + school);
  });
}


function getTodaysMail(sender) {
  const today = new Date().toDateString();
  const threads = GmailApp.search('from:' + sender);
  
  return threads
    .map(t => t.getMessages()[0])
    .filter(m => m.getDate().toDateString() === today);
}


function getAbsence(message){
  const content = message.getPlainBody();
  const lines = content.split('\n');
  const results = [];

  for (let line of lines) {

    const values = line.split(',').map(v => v.trim());
    if (values.length === 3){
      const [name, year, absence] = values;
      results.push({
        name,
        year,
        absence: parseFloat(absence),
      });
    }
  }
  return results;
}


function sortByClass(absences){
  let newAbsence = [];
  absences.sort((a, b) => {
  const onlyLetters = /^[A-Za-z]+$/;
    const aIsOnlyLetters = onlyLetters.test(a.year);
    const bIsOnlyLetters = onlyLetters.test(b.year);

    if (aIsOnlyLetters && !bIsOnlyLetters) return -1;
    if (!aIsOnlyLetters && bIsOnlyLetters) return 1;

    return a.year.localeCompare(b.year);
  });

  absences.forEach(obj => {
    if (!/[4-9]/.test(obj.year[1])) {
    newAbsence.push(obj);
    }
  });
  return newAbsence;
}


function getPreviousTotal(spreadsheet){
  let sheet = spreadsheet.getSheetByName('Sammanställning');
  const rawPreviousTotal = sheet.getRange(CONFIG.absenceTotalCell).getValue();
  if (!rawPreviousTotal) return {};
  const cleanedPreviousTotal = rawPreviousTotal.replace(/^{|}$/g, '');
  let previousTotal = {};

  const entries = cleanedPreviousTotal.split(/},\s*/); 
  entries.forEach(entry => {
    const [name, dataStr] = entry.split("={");

    const dataPairs = dataStr.replace(/}$/, "").split(", ");
    const dataObj = {};

    dataPairs.forEach(pair => {
      let [key, value] = pair.split("=");
      if (key == 'total'){
        value = parseFloat(value);
      } 
      dataObj[key.trim()] = value;
    });

    previousTotal[name.trim()] = dataObj;
  });
  return previousTotal;
}


function getCurrentSchoolYear(monthInt, yearInt){
  if (monthInt > 7) {
    return (yearInt + '/' + (yearInt + 1));
  }else{
    return ((yearInt - 1) + '/' + yearInt);
  }
}


function createSheet(fileName){
  const spreadsheet = SpreadsheetApp.create(fileName);
  spreadsheet.renameActiveSheet('Sammanställning');
  let sheet = spreadsheet.getActiveSheet();
  let range = sheet.getRange(CONFIG.scaleRange);
  const absenceScaleColors = [[CONFIG.absenceColors[1],CONFIG.absenceColors[2],CONFIG.absenceColors[3],CONFIG.absenceColors[4],CONFIG.absenceColors[5]],
                              [CONFIG.absenceColors[6],CONFIG.absenceColors[7],CONFIG.absenceColors[8],CONFIG.absenceColors[9],CONFIG.absenceColors[10]]];
  const absenceScale = [[1,2,3,4,5],
                        [6,7,8,9,'+10']];

  range.setValues(absenceScale)
      .setFontWeight('bold') 
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment('center')
      .setBackgrounds(absenceScaleColors);


  for (let col = 1; col <= 15; col++) {
      sheet.setColumnWidth(col, 150);
    }

  CONFIG.sheetOrder.forEach(i => {

    let month = CONFIG.months[i];
    spreadsheet.insertSheet(month);
    sheet = spreadsheet.getSheetByName(month);
    range = sheet.getRange(CONFIG.scaleRange);
    range.setValues(absenceScale)
      .setFontWeight('bold') 
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment('center')
      .setBackgrounds(absenceScaleColors);

    for (let col = 1; col <= 15; col++) {
      sheet.setColumnWidth(col, 150);
    }
  });
  return spreadsheet;
}


function createMonthsTableHeader(month, spreadsheet){
  let sheet = spreadsheet.getSheetByName(CONFIG.months[month]);
  const row = 7;
  const startCol = 1;
  const titles = ['Namn', 'Klass', 'Fårnvaro', 'Orsak','Risk'];
  sheet.getRange(row, startCol, 1, titles.length)
    .setValues([titles])
    .setBorder(true, true, true, true, true, true)      
    .setFontWeight('bold');

}


function createSummaryTableHeader(spreadsheet){
  let sheet = spreadsheet.getSheetByName('Sammanställning');
  const row = 7;
  const startCol = 1;
  const titles = ['Namn', 'Klass', 'Senast över 15%'];
  sheet.getRange(row, startCol, 1, titles.length)
    .setValues([titles])
    .setBorder(true, true, true, true, true, true)      
    .setFontWeight('bold');
}


function createMonthsTable(month, spreadsheet, absences, totalAbsence){
  let data = absences.map(person => [person.name, person.year, person.absence]);
  let sheet = spreadsheet.getSheetByName(CONFIG.months[month]);
  let row = 8;
  const startCol = 1;

  sheet.getRange(row, startCol, data.length, data[0].length).setValues(data);
  
  for (let col = 1; col <= 15; col++) {
    sheet.setColumnWidth(col, 150);
  }

  let cellValue = sheet.getRange('A' + row).getValue();
  while (cellValue != ''){
    if (totalAbsence[cellValue]['total'] < 10) {
      sheet.getRange('A' + row + ':E' + row).setBackground(CONFIG.absenceColors[totalAbsence[cellValue]['total']]);
    } else {
      sheet.getRange('A' + row + ':E' + row).setBackground(CONFIG.absenceColors[10]);
    }
    row++;
    cellValue = sheet.getRange('A' + row).getValue();
  }
}


function getTotalAbsence(sortedAbsences, totalAbsence, month){
  sortedAbsences.forEach(pupil =>{
    if (pupil.name in totalAbsence){
      totalAbsence[pupil.name]['total'] = totalAbsence[pupil.name]['total'] + 1;
      totalAbsence[pupil.name]['lastAbcense'] = CONFIG.months[month];
    } else {
      totalAbsence[pupil.name] = {'year': pupil.year, 'lastAbcense': CONFIG.months[month], 'total': 1};
    };
  });
  return totalAbsence;
}


function setTotalAbsence(totalAbsence, spreadsheet){
  let sheet = spreadsheet.getSheetByName('Sammanställning');
  sheet.getRange(CONFIG.absenceTotalCell).setValue(totalAbsence);
}


function updateSummary(spreadsheet, totalAbsence){
  let keys = Object.keys(totalAbsence);
  let data = [];
  keys.forEach(pupil => {
    data.push([pupil, totalAbsence[pupil]['year'], totalAbsence[pupil]['lastAbcense']]);
  });

  data.sort((a, b) => {
  const onlyLetters = /^[A-Za-z]+$/;

  const aIsOnlyLetters = onlyLetters.test(a[1]);
  const bIsOnlyLetters = onlyLetters.test(b[1]);

  if (aIsOnlyLetters && !bIsOnlyLetters) return -1;
  if (!aIsOnlyLetters && bIsOnlyLetters) return 1;

  return a[1].localeCompare(b[1]);
  });
  
  let sheet = spreadsheet.getSheetByName('Sammanställning');
  let row = 8;
  const startCol = 1;

  sheet.getRange(row, startCol, data.length, data[0].length).setValues(data);

  let cellValue = sheet.getRange('A' + row).getValue();
  while (cellValue != ''){
    if (totalAbsence[cellValue]['total'] < 10) {
      sheet.getRange('A' + row + ':C' + row).setBackground(CONFIG.absenceColors[totalAbsence[cellValue]['total']]);
    } else {
      sheet.getRange('A' + row + ':C' + row).setBackground(CONFIG.absenceColors[10]);
    }
    row++;
    cellValue = sheet.getRange('A' + row).getValue();
  }
}