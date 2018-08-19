// SHEET IDS

// Robustified Index Master: 1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q
// Test Rig iQA: 1-5Vf4LbGOI29eVabBluk8WoHg5-8qJkzhgazLdLtVDE
// 1301 Data Storage: 1W8ECF6uqytFJJ927CH3Z5-Ki5sYR0mgv69UWHRt-wSk

// Alison Copywriting: 1O1t3I_BYILVjmiLXPeu_BGWIYU009XixjsdD1A8DOVM
// Test Rig iWriter: 1O1t3I_BYILVjmiLXPeu_BGWIYU009XixjsdD1A8DOVM



function onOpen2(){  

  var ui = SpreadsheetApp.getUi();
	ui.createMenu("Index")
    .addItem("1 Call Events", "pEv")
    .addItem("2 Write Event ID", "eventQA2")
	.addItem("3 Plan Copywriters", "pCopy")
    .addItem("4 Push to Copywriter", "qCopy")
    .addItem("5 Populate the Hopper", "uHop2")
    .addItem("6 Process the Hopper", "uHop")
    .addItem("7 Clear Sendwithus", "SWUclear")
    .addItem("8 Write to Sendwithus", "Email")
    .addItem("9 User Data", "data")
	.addToUi();
   
}



function pEv(){
 
  var iEventID = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID"); // locates Event ID sheet
  var ePs2 = iEventID.getRange("ad3:ad2991").getValues(); // creates an array of column 30 in Event ID sheet where the holding ID of events that have yet to be QA'd
  var ePfull = iEventID.getRange("a3:as2991").getValues(); // takes all Event ID and makes it an array
  var evQueue = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event QA"); // Locates the Event QA sheet
  
  for (var i=2; i < 2990; i++){
    
    if(ePs2[i] > 0){
     
      var print2 = evQueue.getRange(2,1).getValue(); // identifies where to print the details
      var xContact = ePfull[i][5]; // contact
      var xAccount = ePfull[i][6]; // account
      var xEvent = ePfull[i][13]; // event note
      var xLabel = ePfull[i][14]; // event label
      var xUEmail = ePfull[i][10]; // user email
      var xEURL = ePfull[i][15]; // event URL
      var xID = ePfull[i][0]; // event ID - maybe also i+1
      var xUser = ePfull[i][1]; // user 
      var xCompany = ePfull[i][2]; // user's compay
      var xSheet = ePfull[i][3]; // user account sheet link
      var xLink = ePfull[i][7]; // contact LinkedIn
      var xCEmail = ePfull[i][8]; // contact email 
      var xDate = ePfull[i][16]; // event Date
      var xRole = ePfull[i][12]; // contact Role
      var xLastB = ePfull[i][23]; // last bmail date
      
      evQueue.getRange(print2,1).setValue(xID);
      evQueue.getRange(print2,2).setValue(xUser);
      evQueue.getRange(print2,3).setValue(xCompany);
      evQueue.getRange(print2,4).setValue(xContact);
      evQueue.getRange(print2,5).setValue(xAccount);
      evQueue.getRange(print2,6).setValue(xLink);
      evQueue.getRange(print2,7).setValue(xRole);
      evQueue.getRange(print2,8).setValue(xEvent);
      evQueue.getRange(print2,9).setValue(xLabel);
      evQueue.getRange(print2,10).setValue(xEURL);
      evQueue.getRange(print2,11).setValue(xDate);
      evQueue.getRange(print2,17).setValue(xLastB);
      
    }
    
  }   
  
}
  
function pCopy(){
  
  var iEventID = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID"); // locates Event ID sheet
  var ePs = iEventID.getRange("aj3:aj2991").getValues(); // creates an array of column 36 in Event ID sheet where the holding ID of to be drafted events is kept
  var ePfull = iEventID.getRange("a3:bj2991").getValues();
  
  for (var i=2; i < 2990; i++){
    
    if(ePs[i] > 0){
     
      var pEvents = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Passed Events"); // Locates the Passed Events sheet
      var print = pEvents.getRange(1,1).getValue(); // identifies where to print the details
      var xUser = ePfull[i][1];
      var xCompany = ePfull[i][2];
      var xEvent = ePfull[i][13];
      var xLabel = ePfull[i][14];
     pEvents.getRange(print+2,2).setValue(i+1);
     pEvents.getRange(print+2,4).setValue(xUser);
     pEvents.getRange(print+2,5).setValue(xCompany);
     pEvents.getRange(print+2,6).setValue(xEvent);
     pEvents.getRange(print+2,7).setValue(xLabel);
     
     
    }
    
  } 

}



function qCopy() {
  
  var Karen = SpreadsheetApp.openById("11Vy-fUMSHGbkXf5hnlIHna1uCNHhK7uPT4rPSwxZf6Y").getSheetByName("Event Queue");
  var Ronnie = SpreadsheetApp.openById("1ixgXOMUT9PHOyKeV5K1XCEmGnSg-B2k0hjQtQyZn3FU").getSheetByName("Event Queue");
  var Alison = SpreadsheetApp.openById("1O1t3I_BYILVjmiLXPeu_BGWIYU009XixjsdD1A8DOVM").getSheetByName("Event Queue");
  var eventIDss = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID");
  var passEvent = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Passed Events");
  var cWriters = passEvent.getRange("b2:C102").getValues(); // copywriter IDs
  var ePfull = eventIDss.getRange("a3:bj2991").getValues(); // takes all Event ID and makes it an array
  
  for(var i = 0; i < 100; i++){
    
      var writer = cWriters[i][1];
    
    
    if(writer == "Ronnie"){
      
      var preID = cWriters[i][0]; // finds the ID of the event that has been allotted to this writer
      var eID = preID-1;
      Logger.log(eID);
      var xContact = ePfull[eID][5]; // contact
      var xAccount = ePfull[eID][6]; // account
      var xEvent = ePfull[eID][13]; // event note
      var xLabel = ePfull[eID][14]; // event label
      var xUEmail = ePfull[eID][10]; // user email
      var xEURL = ePfull[eID][15]; // event URL
      var xID = ePfull[eID][0]; // event ID - maybe also i+1
      var xUser = ePfull[eID][1]; // user 
      var xCompany = ePfull[eID][2]; // user's compay
      var xSheet = ePfull[eID][3]; // user account sheet link
      var xLink = ePfull[eID][7]; // contact LinkedIn
      var xCEmail = ePfull[eID][8]; // contact email 
      var xDate = ePfull[eID][16]; // event Date
      var xRole = ePfull[eID][12]; // contact Role
      var xUstory = ePfull[eID][11]; // user story
      var xSURL = ePfull[eID][22]; // Home static
       
      
      var ronps = Ronnie.getRange(1,1).getValue();
      var ronp = ronps*10;
      
      Ronnie.getRange(ronp+2,1).setValue(xID);
      Ronnie.getRange(ronp+2,2).setValue(xUser);
      Ronnie.getRange(ronp+2,3).setValue(xCompany);
      Ronnie.getRange(ronp+2,4).setValue(xContact);
      Ronnie.getRange(ronp+2,5).setValue(xAccount);
      Ronnie.getRange(ronp+2,6).setValue(xDate);
      Ronnie.getRange(ronp+4,2).setValue(xEvent);
      Ronnie.getRange(ronp+4,4).setValue(xEURL);
      Ronnie.getRange(ronp+4,6).setValue(xLink);
      Ronnie.getRange(ronp+4,5).setValue(xLabel);
      Ronnie.getRange(ronp+9,5).setValue(xRole);
      Ronnie.getRange(ronp+9,6).setValue(xUstory);
      Ronnie.getRange(ronp+9,4).setValue(xSURL);
      
      eventIDss.getRange(preID+2,37).setValue("Ronnie");
     
      
    } else if(writer == "Karen"){
        
      Logger.log(writer);
      var preID = cWriters[i][0]; // finds the ID of the event that has been allotted to this writer
      var eID = preID-1;
      Logger.log(eID);
      var xContact = ePfull[eID][5]; // contact
      var xAccount = ePfull[eID][6]; // account
      var xEvent = ePfull[eID][13]; // event note
      var xLabel = ePfull[eID][14]; // event label
      var xUEmail = ePfull[eID][10]; // user email
      var xEURL = ePfull[eID][15]; // event URL
      var xID = ePfull[eID][0]; // event ID - maybe also i+1
      var xUser = ePfull[eID][1]; // user 
      var xCompany = ePfull[eID][2]; // user's compay
      var xSheet = ePfull[eID][3]; // user account sheet link
      var xLink = ePfull[eID][7]; // contact LinkedIn
      var xCEmail = ePfull[eID][8]; // contact email 
      var xDate = ePfull[eID][16]; // event Date
      var xRole = ePfull[eID][12]; // contact Role
      var xUstory = ePfull[eID][11]; // user story
      var xSURL = ePfull[eID][22]; // user story
      
      Logger.log(xRole);
      Logger.log(xUstory);
      
      var karens = Karen.getRange(1,1).getValue();
      var karp = karens*10;
      
      Karen.getRange(karp+2,1).setValue(xID);
      Karen.getRange(karp+2,2).setValue(xUser);
      Karen.getRange(karp+2,3).setValue(xCompany);
      Karen.getRange(karp+2,4).setValue(xContact);
      Karen.getRange(karp+2,5).setValue(xAccount);
      Karen.getRange(karp+2,6).setValue(xDate);
      Karen.getRange(karp+4,2).setValue(xEvent);
      Karen.getRange(karp+4,4).setValue(xEURL);
      Karen.getRange(karp+4,6).setValue(xLink);
      Karen.getRange(karp+4,5).setValue(xLabel);
      Karen.getRange(karp+9,5).setValue(xRole);
      Karen.getRange(karp+9,6).setValue(xUstory); 
      Karen.getRange(karp+9,4).setValue(xSURL);     

      
      eventIDss.getRange(preID+2,37).setValue("Karen");  
      
    }  else if(writer == "Alison"){
        
      Logger.log(writer);
      var preID = cWriters[i][0]; // finds the ID of the event that has been allotted to this writer
      var eID = preID-1;
      Logger.log(eID);
      var xContact = ePfull[eID][5]; // contact
      var xAccount = ePfull[eID][6]; // account
      var xEvent = ePfull[eID][13]; // event note
      var xLabel = ePfull[eID][14]; // event label
      var xUEmail = ePfull[eID][10]; // user email
      var xEURL = ePfull[eID][15]; // event URL
      var xID = ePfull[eID][0]; // event ID - maybe also i+1
      var xUser = ePfull[eID][1]; // user 
      var xCompany = ePfull[eID][2]; // user's compay
      var xSheet = ePfull[eID][3]; // user account sheet link
      var xLink = ePfull[eID][7]; // contact LinkedIn
      var xCEmail = ePfull[eID][8]; // contact email 
      var xDate = ePfull[eID][16]; // event Date
      var xRole = ePfull[eID][12]; // contact Role
      var xUstory = ePfull[eID][11]; // user story
      var xSURL = ePfull[eID][22]; // user story
      
      
      var alisons = Alison.getRange(1,1).getValue();
      var ali = alisons*10;
      
      Alison.getRange(ali+2,1).setValue(xID);
      Alison.getRange(ali+2,2).setValue(xUser);
      Alison.getRange(ali+2,3).setValue(xCompany);
      Alison.getRange(ali+2,4).setValue(xContact);
      Alison.getRange(ali+2,5).setValue(xAccount);
      Alison.getRange(ali+2,6).setValue(xDate);
      Alison.getRange(ali+4,2).setValue(xEvent);
      Alison.getRange(ali+4,4).setValue(xEURL);
      Alison.getRange(ali+4,6).setValue(xLink);
      Alison.getRange(ali+4,5).setValue(xLabel);
      Alison.getRange(ali+9,5).setValue(xRole);
      Alison.getRange(ali+9,6).setValue(xUstory);
      Alison.getRange(ali+9,4).setValue(xSURL);
     
      
      eventIDss.getRange(preID+2,37).setValue("Alison");  
      
    }  
  
  
 }
  
  passEvent.getRange("b2:j250").clearContent();
  
}


function eventQA2() {
  
  var eventQAss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event QA");
  var eventIDss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event ID");
  var robo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Upload");
  var ss1 = SpreadsheetApp.openById("1W8ECF6uqytFJJ927CH3Z5-Ki5sYR0mgv69UWHRt-wSk").getSheetByName("Sheet2"); // locates 1301 Data Storage
  var data = ss1.getRange(2,1,2000,1).getValues(); // sets an array of the last 2000 events 
  var qArea = eventQAss.getRange("a3:m253").getValues();
  var tDate = new Date();
  
  
  var SQL = ss1.getRange("A1:A4000").getValues();
 
  
  for(var i = 0; i < 250; i++){
    
    if(qArea[i][12] !== ""){    // this is the QA label column
      
     
      
     eventIDss.getRange(qArea[i][0]+2,27).setValue(qArea[i][12]); // writes the event QA label to Event ID
     eventIDss.getRange(qArea[i][0]+2,26).setValue(qArea[i][11]); // writes the event QA notes to Event ID
     eventIDss.getRange(qArea[i][0]+2,14).setValue(qArea[i][7]); // writes the event QA notes to Event ID 
      
     var eM = eventIDss.getRange(qArea[i][0]+2,18).getValue();
     var qA = eventIDss.getRange(qArea[i][0]+2,27).getValue();
     var qAComm = eventIDss.getRange(qArea[i][0]+2,26).getValue();
     var desc = eventIDss.getRange(qArea[i][0]+2,31).getValue();
     var user = eventIDss.getRange(qArea[i][0]+2,2).getValue();
     var id = eventIDss.getRange(qArea[i][0]+2,1).getValue();
     var contact = eventIDss.getRange(qArea[i][0]+2,6).getValue();
      
     MailApp.sendEmail(eM,user, desc + "      " + qA + "     " + qAComm + "    Event ID " + id);
      
      
      // this clause adds the QA'd events to the information on the event already in 1301
      
    
    var label = qArea[i][8];  // the label column
    var analyst = eM;
    var QA = qArea[i][12];  // the label column
    var idg = robo.getRange(qArea[i][0]+1,24).getValue();
   //   Logger.log("does it get to prep?")
     //  Logger.log(user)
     //   Logger.log(contact)
     //   Logger.log(qArea[i][0]+1)
     //   Logger.log(idg)
        
        
      for (var z = 0; z<4000; z++){
        var blah = SQL[z];
        var lame = blah.toString();
      
        if(lame.indexOf("indexed2")<0 && lame.indexOf(user)>=0 &&  lame.indexOf(contact)>0 && lame.indexOf(idg)>0){ 
           var big = [blah,analyst,tDate,label,QA,"indexed2"];
            var newblah = big.join();
         // Logger.log("first a get")
             var t = ss1.getRange(z+1,1).getValue();
         // Logger.log(t)
            ss1.getRange(z+1,1).setValue(newblah);
         // Logger.log(z+1)
         // Logger.log(newblah)
         // Logger.log("action")
         
        }      
      }
    }
  }
  
 eventQAss.getRange("A3:s253").clearContent();
  
}


function uHop(){
  
  var eventIDss = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID");
  var hopper = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Hopper");
  var ePfull = eventIDss.getRange("a3:bj2991").getValues(); // takes all Event ID and makes it an array
  var qcSearch = hopper.getRange("a2:f400").getValues(); // array of Hopper
  var swu = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Sendwithus");
  
  for (var b = 0; b < 390; b++){
   
    var hop = qcSearch[b][2]; // b = 7
    
    if(hop =="Push to Email"){
      
      var nHead = qcSearch[b-4][4]
      var nGreet = qcSearch[b-3][4]
      var nBody = qcSearch[b-2][4]
      var eNum = qcSearch[b-7][0]
      var note = qcSearch[b-1][3]
      var ch1 = qcSearch[b-4][3]
      var ch2 = qcSearch[b-3][3]
      var ch3 = qcSearch[b-2][3]
      var enote = qcSearch[b-5][1]
      var newenote = qcSearch[b-6][2]
      var URL = qcSearch[b-5][3]
      var label = qcSearch[b-5][4]
      var eID = ePfull[eNum-1][0];
      var cEm = ePfull[eNum-1][40];
      
     
      
      if(ch1 !== "" || ch2 !== "" || ch3 !== "" || newenote !== ""){    // if any of the overrides (edit boxes) have been used then var ch is declared as "Change" 
                                                     // note logical operators || is OK, && is AND, ! is NOT
       var ch = "Change";
        
      MailApp.sendEmail(cEm + "," + "denis.anscomb@gmail.com", "Change made to bmail to " 
                        + ePfull[eNum-1][1] + " for event " + eNum, "Hi " + ePfull[eNum-1][36] 
                       + "  The copy linked to this event was edited:  " 
                      + ePfull[eNum-1][30] + "  " + "                 From:                              "+ ePfull[eNum-1][32] + "  "
                        + ePfull[eNum-1][33] + "  " + ePfull[eNum-1][34]  
                        + "         To:                      " + newenote + "    " + nHead + " . " +nGreet + "  " + nBody 
                        + "            Note:                            " + note);  // feedback email to writers
      
      }
      
      var qc = qcSearch[b-7][0];
      Logger.log(qc);
      Logger.log(b); 
      
      eventIDss.getRange(qc+2,49).setValue(nHead);
      eventIDss.getRange(qc+2,53).setValue(nGreet);
      eventIDss.getRange(qc+2,54).setValue(nBody);
      eventIDss.getRange(qc+2,38).setValue("email");
      eventIDss.getRange(qc+2,56).setValue(note);
      eventIDss.getRange(qc+2,52).setValue("Archive");
      eventIDss.getRange(qc+2,32).setValue(""); // gets rid of the 'Drafted" tag that Hopper uses to pull in events
      //eventIDss.getRange(qc+2,38).setValue("");
      eventIDss.getRange(qc+2,16).setValue(URL);
      if(newenote !== ""){eventIDss.getRange(qc+2,14).setValue(newenote);}
      
      hopper.getRange(b-4,3).clearContent();
      
      hopper.getRange(b-2,4).clearContent();
      hopper.getRange(b-1,4).clearContent();
      hopper.getRange(b,4).clearContent();
      hopper.getRange(b-2,3).clearContent();
      hopper.getRange(b-1,3).clearContent();
      hopper.getRange(b,3).clearContent();
      hopper.getRange(b-5,1).clearContent();
      hopper.getRange(b-5,2).clearContent();
      hopper.getRange(b-5,3).clearContent();
      hopper.getRange(b-5,4).clearContent();
      hopper.getRange(b-5,5).clearContent();
      hopper.getRange(b-5,6).clearContent();
      hopper.getRange(b-3,2).clearContent();
      hopper.getRange(b-3,4).clearContent();
      hopper.getRange(b-3,5).clearContent();
      hopper.getRange(b-3,6).clearContent();
      hopper.getRange(b+2,5).clearContent();
      hopper.getRange(b+2,3).clearContent();
      hopper.getRange(b+1,4).clearContent();
      

      
    }else if(hop =="Leave in Queue"){
      
      var nHead = qcSearch[b-4][4]
      var nGreet = qcSearch[b-3][4]
      var nBody = qcSearch[b-2][4]
      var note = qcSearch[b-1][3]
     
      
      var qc = qcSearch[b-7][0];
      //Logger.log(qc);
     // Logger.log(b);
      
      eventIDss.getRange(qc+2,49).setValue(nHead);
      eventIDss.getRange(qc+2,53).setValue(nGreet);
      eventIDss.getRange(qc+2,54).setValue(nBody);
      eventIDss.getRange(qc+2,56).setValue(note);
      
      hopper.getRange(b-4,3).clearContent();
      hopper.getRange(b-2,4).clearContent();
      hopper.getRange(b-1,4).clearContent();
      hopper.getRange(b,4).clearContent();
      hopper.getRange(b-2,3).clearContent();
      hopper.getRange(b-1,3).clearContent();
      hopper.getRange(b,3).clearContent();
      hopper.getRange(b-5,1).clearContent();
      hopper.getRange(b-5,2).clearContent();
      hopper.getRange(b-5,3).clearContent();
      hopper.getRange(b-5,4).clearContent();
      hopper.getRange(b-5,5).clearContent();
      hopper.getRange(b-5,6).clearContent();
      hopper.getRange(b-3,2).clearContent();
      hopper.getRange(b-3,4).clearContent();
      hopper.getRange(b-3,5).clearContent();
      hopper.getRange(b-3,6).clearContent();
      hopper.getRange(b+2,5).clearContent();
      hopper.getRange(b+2,3).clearContent();
      hopper.getRange(b+1,4).clearContent();
      
      
    } else if(hop == "Archive"){
      
      var nHead = hopper.getRange(b-2,5).getValue();
      var nGreet = hopper.getRange(b-1,5).getValue();
      var nBody = hopper.getRange(b,5).getValue();
      var note = hopper.getRange(b+1,4).getValue();
      
      var qc = qcSearch[b-7][0];
      Logger.log(qc);
      Logger.log(b); 
      
      eventIDss.getRange(qc+2,49).setValue(nHead);
      eventIDss.getRange(qc+2,53).setValue(nGreet);
      eventIDss.getRange(qc+2,54).setValue(nBody);
      eventIDss.getRange(qc+2,52).setValue("Archive");
      eventIDss.getRange(qc+2,32).setValue("");
      eventIDss.getRange(qc+2,38).setValue("");
      eventIDss.getRange(qc+2,56).setValue(note);
      
      hopper.getRange(b-4,3).clearContent();
      hopper.getRange(b-2,4).clearContent();
      hopper.getRange(b-1,4).clearContent();
      hopper.getRange(b,4).clearContent();
      hopper.getRange(b-2,3).clearContent();
      hopper.getRange(b-1,3).clearContent();
      hopper.getRange(b,3).clearContent();
      hopper.getRange(b-5,1).clearContent();
      hopper.getRange(b-5,2).clearContent();
      hopper.getRange(b-5,3).clearContent();
      hopper.getRange(b-5,4).clearContent();
      hopper.getRange(b-5,5).clearContent();
      hopper.getRange(b-5,6).clearContent();
      hopper.getRange(b-3,2).clearContent();
      hopper.getRange(b-3,4).clearContent();
      hopper.getRange(b-3,5).clearContent();
      hopper.getRange(b-3,6).clearContent();
      hopper.getRange(b+2,5).clearContent();
      hopper.getRange(b+2,3).clearContent();
      hopper.getRange(b+1,4).clearContent();
    
 
     
      
  }
  }
}
  
  
 function uHop2(){ 
   
  var eventIDss = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID");
  var hopper = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Hopper");
  var ePfull = eventIDss.getRange("a3:bf2991").getValues(); // takes all Event ID and makes it an array
  var qcSearch = hopper.getRange("a2:c400").getValues();
  var swu = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Sendwithus");
  
  
  for(var i = 0; i < 2988; i++){
    
      var drCop = ePfull[i][31]; // looking for any Events in drafted status. 
      var eIDpre = ePfull[i][0];
      var eID = eIDpre-1;
    
    if(drCop == "Drafted"){
      
      Logger.log(i);
      
      
      var xContact = ePfull[eID][5]; // contact
      var xAccount = ePfull[eID][6]; // account
      var xEvent = ePfull[eID][13]; // event note
      var xLabel = ePfull[eID][14]; // event label
      var xUEmail = ePfull[eID][10]; // user email
      var xEURL = ePfull[eID][15]; // event URL
      var xID = ePfull[eID][0]; // event ID - maybe also i+1
      var xUser = ePfull[eID][1]; // user 
      var xCompany = ePfull[eID][2]; // user's compay
      var xSheet = ePfull[eID][3]; // user account sheet link
      var xLink = ePfull[eID][7]; // contact LinkedIn
      var xCEmail = ePfull[eID][8]; // contact email 
      var xDate = ePfull[eID][16]; // event Date
      var xRole = ePfull[eID][12]; // contact Role
      var xHead = ePfull[eID][32]; // Head
      var xGreet = ePfull[eID][33]; // greet
      var xBody = ePfull[eID][34]; // body
      var xFHead = ePfull[eID][48]; // Final Header
      var xFBody = ePfull[eID][50]; // Final Body
      var xNote = ePfull[eID][55]; // QA Note
      
      
      var hPpre = hopper.getRange(1,1).getValue();
      var hp = hPpre*10;
      
      Logger.log(hp);
      
      hopper.getRange(hp+2,1).setValue(xID);
      hopper.getRange(hp+2,2).setValue(xUser);
      hopper.getRange(hp+2,3).setValue(xCompany);
      hopper.getRange(hp+2,4).setValue(xContact);
      hopper.getRange(hp+2,5).setValue(xAccount);
      hopper.getRange(hp+2,6).setValue(xDate);
      hopper.getRange(hp+4,2).setValue(xEvent);
      hopper.getRange(hp+4,4).setValue(xEURL);
      hopper.getRange(hp+4,6).setValue(xLink);
      hopper.getRange(hp+4,5).setValue(xLabel);
      hopper.getRange(hp+5,3).setValue(xHead);
      hopper.getRange(hp+9,5).setValue(xRole);
      hopper.getRange(hp+6,3).setValue(xGreet);
      hopper.getRange(hp+7,3).setValue(xBody);
     // hopper.getRange(hp+5,4).setValue(xFHead);
     // hopper.getRange(hp+7,4).setValue(xFBody);
      hopper.getRange(hp+8,4).setValue(xNote);
      

    }
  }
  
  }

function Email(){
  
  
  var eventIDss = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID");
  var ePfull = eventIDss.getRange("a3:bf2990").getValues(); // takes all Event ID and makes it an array
  var swu = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Sendwithus");
  var renew = swu.getRange("a3:a400");
  renew.clear();
  
  
  for(var m = 0; m < 2988; m++){
    
      var yahoo = ePfull[m][37];
      var eIDpre = ePfull[m][0];
      
    
    if(yahoo == "email"){
      
      
      var checker = ePfull[m][34];
      
      
      var point = swu.getRange(2,1).getValue();
      var stick = (point*24)+3;
      var n = m;
      
      var xContact = ePfull[n][5]; // contact
      var xAccount = ePfull[n][6]; // account
      var xEvent = ePfull[n][13]; // event note
      var xLabel = ePfull[n][14]; // event label
      var xUEmail = ePfull[n][10]; // user email
      var xEURL = ePfull[n][15]; // event URL
      var xID = ePfull[n][0]; // event ID - maybe also i+1
      var xUser = ePfull[n][1]; // user 
      var xCompany = ePfull[n][2]; // user's compay
      var xSheet = ePfull[n][3]; // user account sheet link
      var xLink = ePfull[n][7]; // contact LinkedIn
      var xCEmail = ePfull[n][8]; // contact email 
      var xDate = ePfull[n][16]; // event Date
      var xRole = ePfull[n][12]; // contact Role
      var xHead = ePfull[n][48]; // Head
      var xGreet = ePfull[n][49]; // greet
      var xBody = ePfull[n][50]; // body
      var xDesc = ePfull[n][30]; // description for news
      
      Logger.log(xHead);
      Logger.log(xBody);
      Logger.log(xGreet);
      
       swu.getRange(stick,3).setValue(xUser);
       swu.getRange(stick,1).setValue(xID);
       swu.getRange(stick+1,6).setValue(xUEmail);
       swu.getRange(stick+2,3).setValue(xDesc);
       swu.getRange(stick+3,3).setValue(xCEmail);
       swu.getRange(stick+4,4).setValue(xCompany);
       swu.getRange(stick+5,4).setValue(xContact);
       swu.getRange(stick+9,3).setValue(xHead);
       swu.getRange(stick+11,3).setValue(xBody);
       eventIDss.getRange(xID+2,38).setValue(""); // gets rid of the email tag
  
}

  }
}

function SWUclear() {
  
 var cSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sendwithus");
  
  for (var i = 0; i < 40; i++){
    
   var j = i*24 
   cSS.getRange(j+3,1).clear();
   cSS.getRange(j+3,3).clear();
   cSS.getRange(j+4,6).clear();
   cSS.getRange(j+5,3).clear();
   cSS.getRange(j+6,3).clear();
   cSS.getRange(j+7,4).clear();
   cSS.getRange(j+8,4).clear();
   cSS.getRange(j+12,3).clear();
   cSS.getRange(j+14,3).clear();
    
  }
  
}

function data() {
  
  var eventIDss = SpreadsheetApp.openById("1sEjzhq96me6aaQLIBqY6Wfgy9D6VrKhtHL9eUoqyT2Q").getSheetByName("Event ID");
  var ePfull = eventIDss.getRange("a3:bf2991").getValues(); // takes all Event ID and makes it an array
  var users = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User List"); 
  var user = users.getRange("j2:j36").getValues(); // makes an array of the user list
  
  for (var f = 0; f < 34; f++){
    
    if(user[f][0] !==""){
      
      var cust = user[f][0]; // sets cust to be the user name for each of the users
      var count = 0;
      var cPass = 0;
  
   for(var d = 0; d < 2980; d++){
     
     if(ePfull[d][1] == cust){
          
      var count = count + 1; 
      var eventQAL = ePfull[d][29]; // should be the QA label but check
      var ev1 = ePfull[d][28]; // should be the QA label but check 
   
       if (ev1 == "PASS"){
         var cPass = cPass + 1;}
       
       users.getRange(f+2,11).setValue(count);
       users.getRange(f+2,12).setValue(cPass);
  
      
       }
      } 
     }
   }
  }


function backfill(){ // temp function to add historical QA'd events to the 1301 Data Storage sheet


  var eventIDss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Event ID");
  var robo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Upload");
  var ss1 = SpreadsheetApp.openById("1W8ECF6uqytFJJ927CH3Z5-Ki5sYR0mgv69UWHRt-wSk").getSheetByName("Sheet2"); // locates 1301 Data Storage
  var data = ss1.getRange(2,1,12000,1).getValues(); // sets an array of the last 12000 events 
  var data2 = robo.getRange(1620,1,250,26).getValues();
  var data3 = eventIDss.getRange(1621,1,250,30).getValues();
  
  var SQL = ss1.getRange("A1:A12000").getValues();
  
  var tDate = new Date();
  
  // for the area in question
  
  for(var i = 0; i < 250; i++){
    if(data2[i][23] !=""){
      
    
    var label = data3[i][14];
      var check = data3[i][0];
      var QA = data3[i][26];
    
    var analyst = data2[i][17];
    var id = data2[i][23];
    var user = data2[i][1];
    var contact = data2[i][5];
    
 //   Logger.log(label)
 //   Logger.log(analyst)
 //   Logger.log(id)
 //   Logger.log(user)
 //   Logger.log(contact)
   
    
    // this is the adder:
     for (var z = 0; z<12000; z++){
       var blah = SQL[z];
       var lame = blah.toString();
      // Logger.log(lame.indexOf(id))
       
       if(lame.indexOf("indexed2")<0 && lame.indexOf(user)>=0 &&  lame.indexOf(contact)>=0 && lame.indexOf(id)>=0){ 
          var big = [blah,analyst,tDate,label,QA,"indexed2"];
          var newblah = big.join();
           ss1.getRange(z+1,1).setValue(newblah);
       
        // Logger.log(label)
         //  Logger.log(newblah)
       //  Logger.log(QA)
      //   Logger.log(id)
      //   Logger.log(user)
     //    Logger.log(contact)
    //     Logger.log(z+1)
         
    //      Logger.log(newblah)
        // Logger.log(z+1)
          
        }      
      
  }
    }
  }
}

function what(){
  
  Logger.log("help")
}

