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
  var label = GmailApp.createLabel("Train ticket in calendar");
  var threads = GmailApp.search("from:(bilet.eic@intercity.pl) has:attachment newer_than:40d -label:train-ticket-in-calendar");
  
  Logger.log("Found " + threads.length + " unprocessed emails with tickets.");
  for (var i = 0; i < threads.length; i++) {
    var ticketPdf = threads[i].getMessages()[0].getAttachments()[0].getAs(MimeType.PDF);
    var dateReceived = threads[i].getMessages()[0].getDate();
    var ticketParsed = pdfToText(ticketPdf, { keepTextfile: false, textResult: true });
    var ticketTxt = ticketParsed.text.split('\n');
    var link = ticketParsed.link;
    
    var startMonthDay = ticketTxt[4].split(' ')[0].split('.');
    var startTime = ticketTxt[5].split(' ')[0];
    var startMonth = startMonthDay[1];
    var startDay= startMonthDay[0];
    Logger.log('Departure time: ' + startTime + ', month: ' + startMonth + ', day: ' + startDay);
    Logger.log('Expected format: 14:30, month: 12, day: 24');
    
    var endMonthDay = ticketTxt[8].split(' ')[0].split('.');
    var endTime = ticketTxt[9].split(' ')[0];
    var endMonth = endMonthDay[1];
    var endDay= endMonthDay[0];
    Logger.log('Arrival time: ' + endTime + ', month: ' + endMonth + ', day: ' + endDay);
    
    var from = ticketTxt[6].split(' ')[0];
    var to = ticketTxt[7].split(' ')[0];
    Logger.log('From ' + from + ' to ' + to);
    
    var price = ticketTxt[11].split(': ')[1];
    Logger.log('Price: ' + price);
    
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
    
    var year = dateReceived.getFullYear();
    var timeZoneString = ':00 +0200';
    
    var parsedBackStartTime = Utilities.formatDate(
      new Date(monthNames[startMonth] + ' ' + startDay + ', ' + year + ' ' + startTime + timeZoneString),
      'Europe/Warsaw',
      'HH:mm');
    
    if (parsedBackStartTime != startTime) {
      Logger.log('Changing time zone becasue %s != %s.', startTime, parsedBackStartTime);
      timeZoneString = ':00 +0100';
    }
    Logger.log('timezone string: ' + timeZoneString);
    
    var startDate = new Date(monthNames[startMonth] + ' ' + startDay + ', ' + year + ' ' + startTime + timeZoneString);
    var endDate   = new Date(monthNames[endMonth] + ' ' + endDay + ', ' + year + ' ' + endTime   + timeZoneString);
    
    if (startDate.getTime() < dateReceived.getTime()) {
      var startDate = new Date(monthNames[startMonth] + ' ' + startDay + ', ' + (year + 1) + ' ' + startTime + timeZoneString);
    }
    if (endDate.getTime() < startDate.getTime()) {
      var endDate   = new Date(monthNames[endMonth] + ' ' + endDay + ', ' + (year + 1) + ' ' + endTime   + timeZoneString);
    }
    
    Logger.log('start date: ' + startDate);
    Logger.log('end date: ' + endDate);
    var event = CalendarApp.createEvent('Train ' + from + ' -> ' + to + ', ' + price, startDate, endDate);
    event.setColor(CalendarApp.EventColor.RED);
    event.setDescription(link);
    threads[i].addLabel(label);
  }
}
