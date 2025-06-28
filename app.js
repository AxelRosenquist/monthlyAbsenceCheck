import { EMAIL } from "./mail";

function monthlyAbsenceCheck() {
  const sender = EMAIL;
  const threads = GmailApp.search("from:" + sender);
  let allAbsences = [];


  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      allAbsences = getAbsence(message);
    }
  }
  allAbsences = sortByClass(allAbsences);

  allAbsences.forEach(absence => Logger.log(absence));

  const d = new Date();
  const month = d.getMonth();
  const year = d.getFullYear();
  const months = {7:'Aug',8:'Sep',9:'Okt',10:'Nov',11:'Dec',0:'Jan',1:'Feb',2:'Mar',3:'Apr',4:'Maj',5:'Jun',};
  const schoolYear = getCurrentSchoolYear(month, year);
  const fileName = 'Frånvaro-' + schoolYear;
  const currentMonth = months[month];
  const files = DriveApp.getFilesByName(fileName);
  let spreadsheet;

  Logger.log(fileName)

  if (files.hasNext()){
    const file = files.next();
    spreadsheet = SpreadsheetApp.open(file);
    Logger.log('File found');
  } else {
    createSheet(fileName, months);
    Logger.log('File created');
  }
  // Do operations in document
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
      if (!/[4-9]/.test(obj.year[0])) {
        newAbsence.push(obj);
      }
    });
  return newAbsence;
}

function getCurrentSchoolYear(monthInt, yearInt){
  if (monthInt >= 7) {
    return (yearInt + '/' + (yearInt + 1));
  }else{
    return ((yearInt - 1) + '/' + yearInt);
  }
}

function createSheet(fileName, months){
  const sheetOrder = [7,8,9,10,11,0,1,2,3,4,5];
  const spreadsheet = SpreadsheetApp.create(fileName);
  spreadsheet.renameActiveSheet('Sammanställning');
  sheetOrder.forEach(function(i) {
    let month = months[i];
    spreadsheet.insertSheet(month);
    //Logger.log('Created sheet for: ' + month);
    });
}