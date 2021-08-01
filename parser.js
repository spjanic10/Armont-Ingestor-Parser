function onOpen() {
    const spreadsheet = SpreadsheetApp.getActive();
    let menuItems = [
        {name: 'Gather Alerts', functionName: 'gather'},
    ];
    spreadsheet.addMenu('Alerts', menuItems);
}

function gather() {
    let messages = getGmail();

    let curSheet = SpreadsheetApp.openById("sheetid");

    messages.forEach(message => {curSheet.appendRow(parseEmail(message))});
}

function getGmail() {
    const query = "from:webmaster@landstar.com AND is:unread";
    
    let threads = GmailApp.search(query);

    let messagesList = [];

var messages = GmailApp.getMessagesForThreads(threads);
for (var i = 0 ; i < messages.length; i++) {
  for (var j = 0; j < messages[i].length; j++) {
    if (messages[i][j].isUnread()){
    messagesList.push(messages[i][j].getPlainBody());
    messages[i][j].markRead();}
  }
}
  
    return messagesList;
}


function parseEmail(message){
    let result=[];
    let parsed=message.replaceAll('\n',' ').replaceAll('*','').replaceAll('\r','').replaceAll('  ',' ')
    // console.log(parsed)
  
    
    let loadNumber=parsed.replaceAll(' ','').match('Load#(.*)LoadDetails')
    result.push(loadNumber[1].replace('#',''))

    let agency=parsed.match('Agency(.*)Line Haul')
    result.push(agency[1].replace(' ',''))

    let lineHaul=parsed.match('Haul(.*)Trailer Type')
    result.push(lineHaul[1].replace(' ',''))
    
    let trailerType=parsed.match('Type(.*)Posted [0-9]')
    result.push(trailerType[1].replace(' ',''))

    let postDate=parsed.match("Posted (.*) Contact #")
    let postDateFinal=postDate[1].split(" ")
    result.push(postDateFinal[0])
    result.push(postDateFinal[1])


    let contactNumber=parsed.match('Contact # (.*?) </')
    result.push(contactNumber[1])
    
    let accessorials=parsed.match('Accessorials(.*?)Preloaded')
    result.push(accessorials[1].replaceAll(' ',''))

    
    let preloaded=parsed.match("Preloaded(.*?)Load")
    result.push(preloaded[1].replaceAll(' ',''))

    let loadCount=parsed.match("Count(.*?)Revenue")
    result.push(loadCount[1].replaceAll(' ',''))

    let revenue=parsed.match('Revenue(.*?)Team')
    result.push(revenue[1].replaceAll(' ',''))

    let team=parsed.match('Team(.*?)JIT')
    result.push(team[1].replaceAll(' ',''))

    let jit=parsed.match('JIT(.*?)Miles')
    result.push(jit[1].replaceAll(' ',''))

    
    let miles=parsed.match('Miles(.*?)Driver')
    result.push(miles[1].replaceAll(' ',''))

   
    let driverCollect=parsed.match('Collect(.*?)CSA')
    result.push(driverCollect[1].replaceAll(' ',''))

    let csa=parsed.match('CSA(.*?)Direct')
    result.push(csa[1].replaceAll(' ',''))

    let ratePerMile=parsed.match('Per Mile(.*?)Origin')
    result.push(ratePerMile[1].replaceAll(' ',''))

    let directShipper=parsed.match('Shipper(.*?)Rate')
    result.push(directShipper[1].replaceAll(' ',''))

    let stopsNumber=parsed.match('of Stops(.*?)Type')
    result.push(stopsNumber[1].replaceAll(' ',''))

// Sample Origin Payload
    // [ 'COLUMBUS,','OHIO','Pickup','07/19/2021','08:00','-','07/19/2021','15:00' ]

    let originArray=parsed.match('Origin (.*?) Pick')
    result.push(originArray[1])

    let originDate=parsed.match('Pickup (.*?) [0-9][0-9]:[0-9][0-9]')
    result.push(originDate[1])

    let destinationArray=parsed.match('Destination (.*?) Delivery')
    result.push(destinationArray[1])

    let destinationDate=parsed.match('Delivery (.*?) [0-9][0-9]:[0-9][0-9]')
    result.push(destinationDate[1])
    


    let comments=parsed.match("Comments (.*?)--")
    if (comments){
      result.push(comments[1])
    }
    

    

   

    return result;
}

