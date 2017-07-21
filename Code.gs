function pdfToText(pdfFile, options) {
  // Source: https://gist.github.com/mogsdad/e6795e438615d252584f

  Logger.log("Processing ticket " + pdfFile.getName());
  // Ensure Advanced Drive Service is enabled
  try {
    Drive.Files.list();
  }
  catch (e) {
    throw new Error( "To use pdfToText(), first enable 'Drive API' in Resources > Advanced Google Services." );
  }

  // Save the PDF in drive.
  var pdfName = pdfFile.getName();
  var resource = {
    title: pdfName,
    mimeType: MimeType.PDF,
  };
  var file = Drive.Files.insert(resource, pdfFile);
  
  // Save PDF as GDOC
  resource.title = pdfName.replace(/pdf$/, 'gdoc');
  var insertOpts = {
    ocr: true,
    ocrLanguage: options.ocrLanguage || 'en'
  }
  var gdocFile = Drive.Files.insert(resource, pdfFile, insertOpts);
  
  // Get text from GDOC  
  var gdocDoc = DocumentApp.openById(gdocFile.id);
  var text = gdocDoc.getBody().getText();
  
  // We're done using the GDOC. Delete it.
  Drive.Files.remove(gdocFile.id);
  
  return {
    'text': text,
    'link': file.alternateLink
  };
}

function exportTicketsToCalendar() {
  var today = new Date();
  var label = GmailApp.createLabel("Train ticket in calendar");
  var threads = GmailApp.search("from:(bilet.eic@intercity.pl) has:attachment newer_than:40d -label:train-ticket-in-calendar");
  
  Logger.log("Found " + threads.length + " unprocessed emails with tickets.");
  for (var i = 0; i < threads.length; i++) {
    var ticketPdf = threads[i].getMessages()[0].getAttachments()[0].getAs(MimeType.PDF);
    var ticketParsed = pdfToText(ticketPdf, { keepTextfile: false, textResult: true });
    var ticketTxt = ticketParsed.text.split('\n');
    Logger.log('Ticket txt:\n' + ticketTxt); 
    var link = ticketParsed.link;
    
    var startMonthDay = ticketTxt[0].split(' ')[1].split('.');
    var startTime = ticketTxt[1];
    var startMonth = startMonthDay[1];
    var startDay= startMonthDay[0];
    
    var endMonthDay = ticketTxt[4].split('.');
    var endTime = ticketTxt[5];
    var endMonth = endMonthDay[1];
    var endDay= endMonthDay[0];
    
    var from = ticketTxt[2];
    var to = ticketTxt[3];
    
    var startYear = today.getFullYear();
    var endYear = startYear;
    if (endMonth < startMonth) {
      endYear++;
    }
    
    var monthNames = {
      '01': 'January',
      '02': 'February',
      '03': 'March',
      '04': 'April',
      '05': 'May',
      '06': 'June',
      '07': 'July',
      '08': 'August',
      '09': 'September',
      '10': 'October',
      '11': 'November',
      '12': 'December',
    };
    
    var startDate = new Date(monthNames[startMonth] + ' ' + startDay + ', ' + startYear + ' ' + startTime + ':00 +0200');
    var endDate   = new Date(monthNames[endMonth] + ' ' + endDay + ', ' + endYear + ' ' + endTime   + ':00 +0200');
    Logger.log('start date: ' + startDate);
    Logger.log('end date: ' + endDate);
    var event = CalendarApp.createEvent('Train ' + from + ' ' + to, startDate, endDate);
    event.setColor(CalendarApp.EventColor.RED);
    event.setDescription(link);
    threads[i].addLabel(label);
  }
}
