import { EMAIL } from "./mail";

const CONFIG = {
  emailSender: EMAIL,
  months: {7:'Aug',8:'Sep',9:'Okt',10:'Nov',11:'Dec',0:'Jan',1:'Feb',2:'Mar',3:'Apr',4:'Maj',5:'Jun',},
  sheetOrder: [7,8,9,10,11,0,1,2,3,4,5],
  scaleRange: 'A2:E3',
  absenceColors: {1:'#a4c2f4',2:'#3c78d8',3:'#b6d7a8',4:'#6aa84f',5:'#fff2cc',6:'#ffff00',7:'#f6b26b',8:'#ff9900',9:'#ff0000',10:'#990000'},
  absenceTotalCell: 'V1',
}


function monthlyAbsenceCheck() {
  const todaysMails = getTodaysMail(CONFIG.emailSender);

  todaysMails.forEach(function(mail){
    const allAbsences = getAbsence(mail);
    const allSortedAbsences = sortByClass(allAbsences);
    const allSortedFilteredAbsences = allSortedAbsences.filter(entry => entry.absence >= 15.0);
    
    let school = allSortedAbsences[0].year[0];
    if (school == "y" || school == "Y") {
      school = "Ydre";
    } else {
      school = "Hestra";
    }
    
    const d = new Date();
    const month = d.getMonth();
    const year = d.getFullYear();
    const schoolYear = getCurrentSchoolYear(month, year);

    const fileName = 'Frånvaro-' + school + '-' + schoolYear;
    const currentMonth = CONFIG.months[month];
    const files = DriveApp.getFilesByName(fileName);
    let spreadsheet;

    if (!files.hasNext()){
      spreadsheet = createSheet(fileName);
      Logger.log('File created for instance ' + school);
    } else {
      const file = files.next();
      spreadsheet = SpreadsheetApp.open(file);
      Logger.log('File found');
    }
    if (CONFIG.sheetOrder.includes(month + 1)){
      createMonthsTable(month + 1, spreadsheet, allSortedFilteredAbsences);
    }

    const previousTotalAbsence = getPreviousTotal(spreadsheet);
    let totalAbsence = getTotalAbsence(allSortedFilteredAbsences, previousTotalAbsence);

    setTotalAbsence(totalAbsence ,spreadsheet);
  });
}


function getTodaysMail(sender){
  let todaysMail = [];
  const d = new Date();
  let todaysDate = d.toString().split(' ').slice(0,3).join(' ');
  const threads = GmailApp.search("from:" + sender);
  threads.forEach(function(thread){
    const mailDate = thread.getMessages()[0].getDate().toString().split(' ').slice(0,3).join(' ');
    if (mailDate === todaysDate){
      todaysMail.push(thread);
    };
  }); 
  for (let i = 0; i < todaysMail.length; i++){
    todaysMail[i] = todaysMail[i].getMessages()[0];
  }
  return todaysMail;
}


function getAbsence(message){
  const content = message.getPlainBody();
  const lines = content.split('\n');
  const results = [];

  for (let line of lines) {

    const values = line.split(',').map(v => v.trim(' '));
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
    const isALetter = /^[A-Za-z]/.test(a.year);
    const isBLetter = /^[A-Za-z]/.test(b.year);

    if (isALetter && !isBLetter) return -1;
    if (!isALetter && isBLetter) return 1;

    if (a.year < b.year) return -1;
    if (a.year > b.year) return 1;
    return 0;
    });

    absences.forEach(function(obj) {
      if (!/[4-9]/.test(obj.year[1])) {
        newAbsence.push(obj);
      }
    });
  return newAbsence;
}


function getPreviousTotal(spreadsheet){
  let sheet = spreadsheet.getSheetByName('Sammanställning');
  const rawPreviousTotal = sheet.getRange(CONFIG.absenceTotalCell).getValue();
  const cleanedPreviousTotal = rawPreviousTotal.replace(/[{}]/g, "");
  const listPreviousTotal = cleanedPreviousTotal.split(', ');
  let previousTotal = {};

  listPreviousTotal.forEach(pair => {
    const [key, value] = pair.split('=');
    previousTotal[key] = parseFloat(value);  
  });
  return previousTotal;
}


function getCurrentSchoolYear(monthInt, yearInt){
  if (monthInt >= 7) {
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
    .setFontWeight("bold") 
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment("center")
    .setBackgrounds(absenceScaleColors);

  for (let col = 1; col <= 15; col++) {
    sheet.setColumnWidth(col, 150);
  }
  
  CONFIG.sheetOrder.forEach(function(i) {
    let month = CONFIG.months[i];
    spreadsheet.insertSheet(month);
  });
  return spreadsheet
}


function createMonthsTable(month, spreadsheet, absences){
  let data = absences.map(person => [person.name, person.year, person.absence]);
  let sheet = spreadsheet.getSheetByName(CONFIG.months[month]);
  const startRow = 5;
  const startCol = 1;
  sheet.getRange(startRow, startCol, data.length, data[0].length).setValues(data);
  for (let col = 1; col <= 15; col++) {
    sheet.setColumnWidth(col, 150);
  }
}


function getTotalAbsence(sortedAbsences, totalAbsence){
    sortedAbsences.forEach(function(pupil){
      if (pupil.name in totalAbsence){
        totalAbsence[pupil.name] = totalAbsence[pupil.name] + 1;
       } else {
        totalAbsence[pupil.name] = 1;
       };
    });
    return totalAbsence;
}


function setTotalAbsence(totalAbsence, spreadsheet){
  Logger.log(spreadsheet);
  let sheet = spreadsheet.getSheetByName('Sammanställning');
  Logger.log(sheet);

  sheet.getRange(CONFIG.absenceTotalCell).setValue(totalAbsence);
}