import { EMAIL } from "./mail";

const CONFIG = {
  emailSender: EMAIL,
  months: {7:'Aug',8:'Sep',9:'Okt',10:'Nov',11:'Dec',0:'Jan',1:'Feb',2:'Mar',3:'Apr',4:'Maj',5:'Jun',},
  scaleRange: 'A2:E3',
  sheetOrder: [7,8,9,10,11,0,1,2,3,4,5],
}


// Main function
function monthlyAbsenceCheck() {
  const todaysMails = getTodaysMail(CONFIG.emailSender);

  todaysMails.forEach(function(mail){
    const allAbsences = getAbsence(mail);
    const allSortedAbsences = sortByClass(allAbsences);

    allSortedAbsences.forEach(absence => Logger.log(absence));

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

    if (files.hasNext()){
      const file = files.next();
      spreadsheet = SpreadsheetApp.open(file);
      Logger.log('File found');
    } else {
      createSheet(fileName);
      Logger.log('File created');
    }
  
    // Do operations in document
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
  CONFIG.sheetOrder.forEach(function(i) {
    let month = CONFIG.months[i];
    spreadsheet.insertSheet(month);
  });
    let sheet = spreadsheet.getActiveSheet();
    let range = sheet.getRange(CONFIG.scaleRange);
    Logger.log(typeof(range));
}