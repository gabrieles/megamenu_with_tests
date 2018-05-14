//set the global vars
var sheetID = "17DXfn0eVBfqz4pFeahMMPeCjlDtU-s2-5Ik7ck_fl74"; //needed as you cannot use getActiveSheet() while the sheet is not in use (as in a standalone application like this one)
var scriptURL = "https://script.google.com/macros/s/AKfycbyYvxA2LNTnP3r5L4yJW_8BBYvbSb0rwK4YfdAvkv8qsHrWeY8/exec";  //the URL of this web app
var settingsName = "Settings"; //the name of the sheet where all the settings are stored - handy to avoid editing code all the time
var responsibilitiesName = "Responsibilities"; //the name of the sheet where you assign an owner to each and every menu item
var flatListName = "Global menu items"; //the name of the sheet where to print a flat list of all menu items, and their parents. Useful for creating a sitemap in Visio
var sheetHTML = 'Status'

// ******************************************************************************************************
// Function to display the HTML as a webApp
// ******************************************************************************************************
function doGet(e) {
  
  //pass a parameter via the URL as ?action=XXX
  var action = e.parameter.action;
  
  switch(action) {
    case "measure":
      var template = HtmlService.createTemplateFromFile('megamenu_measure');
      break;
    
    case "submit":
      var template = HtmlService.createTemplateFromFile('submit');
      break;
    
    case "showHTML":
      var template = HtmlService.createTemplateFromFile('showHTML');
      break;
      
    default:
      var template = HtmlService.createTemplateFromFile('megamenu');  
  }
      

  var htmlOutput = template.evaluate()
                   .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                   .setTitle('ICR Megamenu');

  return htmlOutput;
};



// ******************************************************************************************************
// Function to add a custom menu every time the worksheet is opened
// ******************************************************************************************************
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
    .addSubMenu(ui.createMenu('On Settings sheet')    
      .addItem('Update list of sheets', 'updateSettingsList')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('On Global menu items sheet')
          .addItem('Update list of items', 'populateList')
    )
    .addToUi();
}



// ******************************************************************************************************
// Function to populate list of pages and where they sit in the menu structure
// ******************************************************************************************************

function populateList() {
  
  var ss = SpreadsheetApp.openById(sheetID);
  var sSettings = ss.getSheetByName(settingsName);
  var outSheet = ss.getSheetByName(flatListName);
   
  var sheetName = '';
  var cellVal = '';
  var lvl = 0;
  var sheetItems = [];
  var sheetItemsLevels = [];
  var sheetItemsParents = [];
  var sheetItemsFormat = [];
  var colPar = '';
  var subPar = '';
  
  
  //clear old data
  outSheet.getRange(2, 1, outSheet.getDataRange().getLastRow(), 3).clear() 
  
  
  //loop over the sheets, and populate the sheetItems array
  for(var h=0; h<ss.getSheets().length; h++){   
    var sheet = ss.getSheets()[h]; // look at every sheet in spreadsheet   
     
    sheetName = sheet.getName();
    
    //check that the sheet is part of the menu structure
    cellVal = sSettings.getRange(h+2, 2, 1,1).getValue();

    if (cellVal == "Primary" || cellVal == "Secondary") {
      sheetItems.push(sheetName);
      sheetItemsLevels.push(0);
      sheetItemsParents.push(cellVal);
      sheetItemsFormat.push(0)
      
      //prepare to loop through all the items
      var lastRow = sheet.getLastRow(); 
      var lastCol = sheet.getLastColumn();
      
      //loop over the columns in the sheet
      for ( var j=1; j<lastCol+1; j++ ){
        
        //set column name as parent for standard items, and add column title to the array withteh worksheet name as its parent
        colPar = sheet.getRange(1, j, 1,1).getValue();
        sheetItems.push(colPar);
        sheetItemsLevels.push(1);
        sheetItemsParents.push(sheetName);
        sheetItemsFormat.push(1)
        
        //loop over the rows
        for ( var i=2; i<lastRow+1; i++ ){
	      cellVal = sheet.getRange(i, j, 1,1).getValue();
          if (cellVal != '') {
            
            //if the item is in italics, it is part of a submenu
            if ( sheet.getRange(i, j, 1,1).getFontStyle() == 'italic' ) {
              sheetItems.push(cellVal);
                       
              //the first italicised item becomes the new submenu parent, and its parent is the column title. All other items have the sumenu parent as a parent.
              if ( subPar == '' ) { 
                sheetItemsParents.push(colPar);
                sheetItemsLevels.push(2);
                sheetItemsFormat.push(3)
                subPar = cellVal; 
              } else {
                sheetItemsParents.push(subPar);
                sheetItemsLevels.push(3);
                sheetItemsFormat.push(4)
              }  
            } else {
              subPar = '';
              sheetItems.push(cellVal);
              sheetItemsLevels.push(2);
              sheetItemsParents.push(colPar);
              sheetItemsFormat.push(2)
            }
            
          } else {
            subPar = '';
          }
	    }
      }
    }  
  }
  
  //print all items in the first column
  
  for(h=0; h<sheetItems.length; h++){  
    outSheet.getRange(h+2,1).setValue(sheetItems[h]);
    outSheet.getRange(h+2,2).setValue(sheetItemsLevels[h]);
    outSheet.getRange(h+2,3).setValue(sheetItemsParents[h]);
    //set format
    setFormat(outSheet, h+2, 1, sheetItemsFormat[h])
  }
}



// ******************************************************************************************************
// Function to set the format of list items on the Global Menu items page
// ******************************************************************************************************
function setFormat(oSheet, oRow, oCol, formatID) {
  
  switch(formatID) {
    case 0:
      oSheet.getRange(oRow,oCol).setFontColor("#fff");
      oSheet.getRange(oRow,oCol).setFontWeight("bold");
      oSheet.getRange(oRow,oCol).setFontStyle("normal");
      oSheet.getRange(oRow,oCol).setBackground("#A71930");
      break;
    
    case 1:
      oSheet.getRange(oRow,oCol).setFontColor("#000");
      oSheet.getRange(oRow,oCol).setFontWeight("bold");
      oSheet.getRange(oRow,oCol).setFontStyle("normal");
      oSheet.getRange(oRow,oCol).setBackground("#fff");
      break;
    
    //case 2: is default  
      
    case 3:
      oSheet.getRange(oRow,oCol).setFontColor("#666");
      oSheet.getRange(oRow,oCol).setFontWeight("bold");
      oSheet.getRange(oRow,oCol).setFontStyle("italic");
      oSheet.getRange(oRow,oCol).setBackground("#fff");
      break;
      
    case 4:
      oSheet.getRange(oRow,oCol).setFontColor("#000");
      oSheet.getRange(oRow,oCol).setFontWeight("normal");
      oSheet.getRange(oRow,oCol).setFontStyle("italic");
      oSheet.getRange(oRow,oCol).setBackground("#fff");
      break;
      
    default:
      oSheet.getRange(oRow,oCol).setFontColor("#000");
      oSheet.getRange(oRow,oCol).setFontWeight("normal");
      oSheet.getRange(oRow,oCol).setFontStyle("normal");
      oSheet.getRange(oRow,oCol).setBackground("#fff");
      break;
  }
  
}



// ******************************************************************************************************
// Function to update the list of sheets in the first column of the Settings page
// ******************************************************************************************************
function updateSettingsList() {
  var ss = SpreadsheetApp.openById(sheetID);
  var sheets = ss.getSheets();
  var sSettings = ss.getSheetByName(settingsName);
  for(h=0; h<sheets.length; h++){  
    sSettings.getRange(h+2,1,1,1).setValue(sheets[h].getName()); 
  }
}




// ******************************************************************************************************
// Function to print out the content of an HTML file into another (used to load the CSS and JS)
// ******************************************************************************************************
function getContent(filename) {
  var pageContent = HtmlService.createTemplateFromFile(filename).getRawContent();
  return pageContent;
}




// ******************************************************************************************************
// Function to print out the content of the sheets with the menu items
// ******************************************************************************************************
function printMenu(action) {
  
  var sheetName = '';
  var codeP = ''; //code for the primary menu
  var codeS = ''; //code for the secondary menu 
  var codeB = ''; //code for the "edit options" menu
  var menuAction = ''; //string to store what to do with each sheet
  var addCode = ''; //additional code to be printed
  var addImgAction = '';
  var addStyle = '';

  
  
  //get settings
  
  var ss = SpreadsheetApp.openById(sheetID);
  var sSettings = ss.getSheetByName(settingsName);
  var sheets = ss.getSheets();
  var primaryColor   = sSettings.getRange(4, 7, 1,1).getValue();
  var secondaryColor = sSettings.getRange(5, 7, 1,1).getValue();
  var maxColMenu     = sSettings.getRange(6, 7, 1,1).getValue(); //max number of columns in the submenu before starting a new row
  var logoImgURL     = sSettings.getRange(7, 7, 1,1).getValue();
  var searchImgURL   = sSettings.getRange(8, 7, 1,1).getValue();
  var measureImgURL  = sSettings.getRange(9, 7, 1,1).getValue();
  var formImgURL     = sSettings.getRange(10, 7, 1,1).getValue();
  var htmlImgURL     = sSettings.getRange(11, 7, 1,1).getValue();
  var pPadding       = sSettings.getRange(12, 7, 1,1).getValue(); //padding on the primary menu - handy to change it if the length of the text changes and this causes the menu to grow too wide. (Default 0 30px)
  var addCSS         = sSettings.getRange(13, 7, 1,1).getValue(); //additional CSS to inject. Just in case

   
  
  
  //prepare code to be added in the menu items based on the action to be performed
  switch(action) {
    
    //if you are testing the menu
    case 'measure': 
      addCode = 'onclick="storeClick(this)"'; 
      codeB = '';
      addImgAction = '<a class="img-action" alt="Submit a new task to be tested" title="Submit a new task to be tested" href="' + scriptURL + '?action=submit" target="_blank"><img width="24" height="24" src="' + formImgURL + '"/></a>';
      break;
      
    //if you are submitting a new task
    case 'submit':
      addCode = '';
      codeB = '';
      break;
    
    default:
      addCode = '';
      codeB = '<a class="control red action" href="' + ss.getUrl() + '" target="_blank">Edit Options</a>';  
      addImgAction = '<a class="img-action" alt="Test this menu" title="Test this menu" href="' + scriptURL + '?action=measure" target="_blank"><img width="24" height="24" src="' + measureImgURL + '"/></a>';
      addImgAction += '<a class="img-action" alt="See the HTML code for this menu" title="See the HTML code for this menu" href="' + scriptURL + '?action=showHTML" target="_blank" style="margin-left: 32px;"><img width="24" height="24" src="' + htmlImgURL + '"/></a>';
  }
   
  
  
  //print menus

  //loop over all the sheets, and add their content to the respective menus as needed
  for(var h=0; h<ss.getSheets().length; h++){   
    var sheet = ss.getSheets()[h]; // look at every sheet in spreadsheet   
    var sheetName = sheet.getName(); //get the sheet name
    
    //grab the action from the "settings" sheet, and add the code to the respective menu
    menuAction = sSettings.getRange(h+2, 2, 1,1).getValue();
    
    switch (menuAction) {
      case "Primary":
        codeP += '<li role="menuitem" class="primary" ' + addCode + '> <button ontouchstart="makeActive(this);">' + sheetName + '</button><div class="mega-menu" aria-hidden="true" role="menu">';
        codeP += createMenuItems(sheet, action, maxColMenu);
        codeP += '</div></li>';  
        break;
      case "Secondary":
        codeS += '<li role="menuitem" class="secondary"> <button ontouchstart="makeActive(this);">' + sheetName + '</button><div class="mega-menu" aria-hidden="true" role="menu">';
        codeS += createMenuItems(sheet, action, maxColMenu);
        codeS += '</div></li>';
      default:
        //do nothing
    }  
  }
  
  //open and close menus as needed
  if(codeP) { codeP = '<div id="primary"   class="menu-wrapper" role="navigation">                    <ul class="nav" role="menubar" id="menu-wrapper">' + codeP + '</ul></div>'; }
  if(codeS) { codeS = '<div id="secondary" class="menu-wrapper" role="navigation">' + addImgAction + '<ul class="nav" role="menubar" id="menu-wrapper">' + codeS + '</ul></div>'; }
  
  
  
  //add searchbox
  var codeSearch = '<div class="menu-wrapper" id="search-wrapper"><input placeholder="Search" type="search" name="search" value="" size="32" maxlength="128" class="input-search"></div>';
  
  
  
  //add inline CSS 
  var cWidth = ['99%','49%','32%','24%','19%','15.5%','11.5%','10%']; //menu group width: store them as an array and choose the right one based on what is set in the settings sheet
  addStyle += '<style>';
    addStyle += '.input-search{background-image: url(\'' + searchImgURL + '\');}';
    addStyle += '.nav > li > button:focus, .nav > li:hover > button, .nav li.active > button{ background-color: ' + primaryColor + ';}';
    addStyle += '#secondary .nav > li > button:focus, #secondary .nav > li:hover > button, level2:hover ul, .nav-column .level2 button:hover {color: ' + primaryColor + ';}';
    addStyle += 'li.level2:hover > button.btn:after { border-left-color: ' + primaryColor + '; }'
    addStyle += '.nav-column button:hover{ color: ' + secondaryColor +';}';
    if(maxColMenu ) { addStyle += 'body .nav-column { width: '+ cWidth[maxColMenu-1] +';}'; }
    if (pPadding){ addStyle += 'body .nav > li > button { padding: ' + pPadding + '; }'; }
    if (addCSS){ addStyle += addCSS ;}
  addStyle += '</style>';
  
  
  
  var codeMenu = addStyle + codeS + codeSearch + codeP + codeB; 
  return codeMenu; 
  
}



// ******************************************************************************************************
// Utility function to return the content of a sheet as a menu
// ******************************************************************************************************

function createMenuItems(sh, ac, mcm) {
  var outC = '';
  var styleCode = '';
  var addCode = '';
  var aCode = '';
  
  var lastRow = sh.getLastRow(); 
  var lastCol = sh.getLastColumn(); 
  
  var cellVal = '' ;
  var cellColor = '';
  var lvl = 1;
  var isIt = false;
  

  if (ac == 'measure'){ aCode = ' onclick="storeClick(this)"'; } 
  
  //loop over the columns - each one is a menu group
  for ( var j=1; j<lastCol+1; j++ ){
 
    if (j == mcm+1) { outC += '<div class="separator"></div>'; }
    
    //Create the menu group container
    outC += '<div class="nav-column">';
    
    //create the menu group heading
    cellVal = sh.getRange(1, j, 1,1).getValue();
    cellColor = sh.getRange(1, j, 1,1).getFontColor();
    styleCode = '';
    if ( cellColor!= '#000000' ) { styleCode += ' style="color: '+ cellColor +'"';}   
    outC += '<h3><button' + styleCode + '' + aCode + '>' + cellVal + '</button></h3>';
        
    //start the menu
    outC+= '<ul>';
    
    // "lvlN" keeps track of the menu item "level" and it printed as its CSS class. Default is lvl1
    //  lvl1 is "root level", 
    //  lvl2 is the parent of a submenu, 
    //  lvl3 is the child of a submenu.
    lvl = 1;
    
    for ( var i=2; i<lastRow+1; i++ ){
          
      cellVal = sh.getRange(i, j, 1,1).getValue();
      if (cellVal != '') {
        
        isIt = sh.getRange(i, j, 1,1).getFontStyle() == 'italic' ? true : false;
        
        addCode = '</li>'; //default value to be used to simply close an item. To be replaced if needed

        if (lvl > 1 )  {
          //if it is not in italic, but was in a submenu, close it (and the parent item)
          if (!isIt) {  
            //stop submenu
            lvl = 1;
            outC += '</ul></li>';
          } else {
            //continue submenu
            lvl  = 3; 
          }
        }

        
        //if the current cell is in italic, but the previous is not, start submenu
        if (isIt && lvl == 1 ) {
            lvl = 2;
            //instead of closing the li, open a new UL
            addCode = '<ul>';        
        } 
        
      
        styleCode = '';
        cellColor = sh.getRange(i, j, 1,1).getFontColor();
        if ( cellColor!= '#000000' ) { styleCode += ' style="color: '+ cellColor +'"';} 
        outC += '<li role="menuitem" class="level' + lvl + '"><button class="btn" ' + styleCode + '' + aCode + '>' + cellVal + '</button>';
        outC += addCode;
        
      } else {
         //if you find a gap after being in a submenu, close it
         if (lvl>1) {
           lvl = 1;
           outC += '</ul></li>';
         }
      }  
        
    }
    
    //if you are at the end of a menu group and still in a submenu, close it
    if (lvl>1) {outC += '</ul></li>';}
    
    
    //close the menu and its container
    outC += '</ul></div>';        
    
  }
  return outC;
}




// ******************************************************************************************************
// Function to print out the HTML necessary to generate the menu
// ******************************************************************************************************
function printMenuHTML(action) {
  
  var sheetName = '';
  var codeP = '';   //code for the primary menu
  var codeS = '';   //code for the secondary menu 
  var codeOut = ''; //code for output
  var menuAction = ''; //string to store what to do with each sheet
  var addCode = ''; //additional code to be printed
  var addImgAction = '';
  var addStyle = '';
  
  //set the sheet to work on
  var ss = SpreadsheetApp.openById(sheetID);
  var sh = ss.getSheetByName(sheetHTML);
  
  var lastRow = sh.getLastRow(); 
 
  var nextItemLvl = 0
  var itemTxt = ''
  var itemURL = ''
  var itemMenu = ''
  var itemLvl = 0
  var closePLvl = new Array('</div> </div> </div> </div> </li>', '</ul> </div>', '</ul> </li>')
  var closeSLvl = new Array(' ', '</ul> </div>', '</ul> </li>')
  
  codeP += '<div id="navbar" class="navbar-collapse collapse" aria-expanded="false">'
  codeP += '<ul class="icr-nav nav navbar-nav">'
  
  //loop over all the rows
  for ( var i=2; i<lastRow+1; i++ ){
    itemTxt     = sh.getRange(i, 1, 1,1).getValue();
    itemURL     = sh.getRange(i, 9, 1,1).getValue();
    itemMenu    = sh.getRange(i, 10, 1,1).getValue();
    itemLvl     = sh.getRange(i, 2, 1,1).getValue();
    nextItemLvl = sh.getRange(i+1, 2, 1,1).getValue();
    
    
    //remove base URL to make the menu more portable
    itemURL = itemURL.replace('https://testchlspweb01:8080', '')
    itemURL = itemURL.replace('http://sharepoint', '')
    //itemURL = itemURL.replace('[PRODUCTION_URL]', '')
    
    
    switch (itemMenu) {
      
      //primary menu
      case 'p':  
     
        switch (itemLvl) {
            
          case 0:        
            codeP += '<li class="dropdown">'
            codeP += '<a href="' + itemURL + '" class="dropdown-toggle" role="button" aria-haspopup="true" aria-expanded="false">' + itemTxt + '</a>'
            codeP += '<div class="dropdown-menu hidden-xs hidden-sm">'
            codeP += '<div class="container">'
            codeP += '<div class="container-md bg-nav">'
            codeP += '<div class="row">'
            if (nextItemLvl == 0) { codeP += closePLvl[0] }
            break;
            
          case 1:
            codeP += '<div class="col-xs-5ths">'
            codeP += '<h4 class="dd-header">' + itemTxt.toUpperCase() + '</h4>'
            codeP += '<ul class="mega-menu-list">'
            if (nextItemLvl == 1) { codeP += closePLvl[1] }
            if (nextItemLvl == 0) { codeP += closePLvl[1] + closePLvl[0] }
            break;
            
          case 2:  
            switch(nextItemLvl) {
              case 3:
                codeP += '<li class="dropdown-submenu">'
                codeP += '<a href="#" class="dropdown-toggle" role="button" aria-haspopup="true" aria-expanded="false">' + itemTxt + '</a>'  
                codeP += '<ul class="dropdown-menu">'  
                break;  
              case 2:
                codeP += '<a href="' + itemURL + '"><li>' + itemTxt + '</li></a>' 
                break;
              case 1:
                codeP += '<a href="' + itemURL + '"><li>' + itemTxt + '</li></a>' + closePLvl[1]  
                break;    
              default:
                codeP += '<a href="' + itemURL + '"><li>' + itemTxt + '</li></a>' + closePLvl[1] + closePLvl[0]    
            }
            break;
            
          case 3:  
            switch(nextItemLvl) {
              case 3:
                codeP += '<a href="' + itemURL + '"><li>' + itemTxt + '</li></a>' 
                break;
              case 2:
                codeP += '<a href="' + itemURL + '"><li>' + itemTxt + '</li></a> ' + closePLvl[2] 
                break;
              case 1:
                codeP += '<a href="' + itemURL + '"><li>' + itemTxt + '</li></a> ' + closePLvl[2] + closePLvl[1]  
                break;    
              default:
                codeP += '<a href="' + itemURL + '"><li>' + itemTxt + '</li></a> ' + closePLvl[2] + closePLvl[1] + closePLvl[0] 
            }
            break;
          default: 
            //do nothing  
        }
        break;
        
      ///end primary
        
        
        
      //secondary menu  
      case 's':
        switch (itemLvl) {
            
          case 0:  
            //do nothing
            break;
            
          case 1:    
            codeS += '<div class="my-icr-col mega-menu-list icr-nav" style="padding-top: 0;">'
            codeS += '<h4 class="dd-header">' + itemTxt.toUpperCase() + '</h4>'
            codeS += '<ul class="content-list pt-20 navbar-nav">'
            if (nextItemLvl == 1 || nextItemLvl == 0) { codeS += closeSLvl[1] }
            break;
            
          case 2:
            switch(nextItemLvl) {
              case 3:
                codeS += '<li class="dropdown-submenu dropdown" style="float:none">'
                codeS += '<a href="#" class="dropdown-toggle" role="button" aria-haspopup="true" aria-expanded="false">' + itemTxt + '</a>'  
                codeS += '<ul class="dropdown-menu" style="top: 0; left: 100%; padding: 0; background-color: #a71930; color: #fff;">'  
                break;
                
              case 2:
                codeS += '<li style="float:none"><a href="' + itemURL + '">' + itemTxt + '</a></li>' 
                break;
                
              default:
                codeS += '<li style="float:none"><a href="' + itemURL + '">' + itemTxt + '</a></li>' 
                codeS += closeSLvl[1]  
            } 
            break;
            
          case 3:  
            switch(nextItemLvl) {
              case 3:
                codeS += '<a href="' + itemURL + '"><li>' + itemTxt + '</li></a>' 
                break;
              case 2:
                codeS += '<a href="' + itemURL + '"><li>' + itemTxt + '</li></a> ' + closeSLvl[2] 
                break;
              default:
                codeS += '<a href="' + itemURL + '"><li>' + itemTxt + '</li></a> ' + closeSLvl[2] + closeSLvl[1]  
                break;    
            }
            break;
            
          default:  
          //do nothing  
        }
        break;
      
      //end secondary  
      
        
      default:
      //do nothing   
    }
      
  }
  
  //close the primary menu
  codeP += '</ul> </div>'     
  
  //prepare for output: place the primary menu in a textarea, and encode the text to display
  codeOut += '<div id="p-wrapper" class="form-item">'
  codeOut += '<h2>Primary menu</h2> <textarea id="primaryMenu" name="primaryMenu">' + htmlEscape( codeP ) + '</textarea>'
  codeOut += '</div>'
  
  //prepare for output: place the primary menu in a textarea, and encode the text to display
  codeOut += '<div id="s-wrapper" class="form-item">'
  codeOut += '<h2>Secondary menu</h2> <textarea id="secondaryMenu" name="secondaryMenu">' + htmlEscape( codeS ) + '</textarea>'
  codeOut += '</div>'
  
  return codeOut; 
  
}


// ******************************************************************************************************
// Function to escape HTML (convert it to text that can be printed as text)
// ******************************************************************************************************
function htmlEscape(str) {
    return str
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

// ******************************************************************************************************
// Function to convert escaped HTML into HTML
// ******************************************************************************************************
function htmlUnescape(str){
    return str
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&');
}



// ******************************************************************************************************
// Function to print out the tasks
// ******************************************************************************************************
function printTasks() {
  
  var codeTask = ''; 
  var sheetName = '';
  var headingName = '';
  var cellVal = '' ;
  var j = 0; 
  var nextID = '';
  var solVal = '';
  var taskList = [];
  var colIndex = 0;
  var maxTasks = 0;
 
  var ss = SpreadsheetApp.openById(sheetID);
  var sheet = ss.getSheetByName("Tasks");
 
  
  var lastCol = sheet.getLastColumn();  
 
  var doWeLimit = sheet.getRange(1, 1, 1,1).getValue();
 
  
  //create array of all task column identifiers, including only tasks that have been marked as "Included"
  for (j=2; j<lastCol; j++ ) { 
    if ( sheet.getRange(3, j, 1,1).getValue() == 'Included' ) {
      taskList.push(j); 
      maxTasks++;
    }  
  } 

  //reset the number of max tasks to display if needed (and if it is lower than the number of included tasks)
  if( doWeLimit != 'Display all') { 
    var mTasks = Number(/\d+/.exec(doWeLimit));
    if ( mTasks < maxTasks ) { maxTasks = mTasks; }
  }
  
  //create list of elements to display
  taskList = shuffle(taskList);           //shuffle the list of tasks
  taskList = taskList.slice(0, maxTasks); //truncate the list of tasks to the desired number
  taskList.push('thankyou');              //the l;ast item to display is the "thank you" box
  
  
  
  //print welcome message
  codeTask += '<div id="welcome" class="boxMessage">'; 
    codeTask += '<p>' + ss.getSheetByName(settingsName).getRange(13, 7, 1,1).getValue().replace(/\n/gm,"</p><p>") + '</p>';   
    codeTask += '<button class="control green" onclick="var el= this.parentNode; hide(el); show(\'task' + taskList[0] + '\');">Hide this message and show the first task</button>';
  codeTask += '</div>';
    
  
 
  //print list of tasks
  codeTask += '<div id="task-wrapper">';
  
 
  for (j=0; j<maxTasks; j++ ){
   
    colIndex = taskList[j];
     
    cellVal = sheet.getRange(1, colIndex, 1,1).getValue();
   
   
   
    //print a task only if there is a task description (or else the feedback column will create trouble)
    if (cellVal.length > 1) {
      solVal = sheet.getRange(2, colIndex, 1,1).getValue();
     
      codeTask += '<div class="task hidden" id="task' + colIndex + '">';
   
        //print the task description and start button
        codeTask += '<div id="taskText' + colIndex + '" class="boxMessage">';  
          codeTask += '<p id="p' + colIndex + '">' + cellVal + '</p>'; //task description
          codeTask += '<button id="start' + colIndex + '" class="control red" onclick="startTask(' + colIndex + ',\'' + solVal + '\') ">Start</button>'; 
        codeTask += '</div>';
     
     
        //print task reminder and the "I give up" button (start as hidden)
        codeTask += '<div id="taskBtn' + colIndex + '" class="hidden taskbtn-wrapper">';  
          codeTask += '<p>Task: '+ cellVal + '</p>'; // add reminder text above the "I give up" button
          codeTask += '<button id="out' + colIndex + '" class="control red giveup" onclick="stopTask(); setSolution(\'I give up\'); ">I give up!</button>';
        codeTask += '</div>';
     
     
        //print the input fields to store the results
        codeTask += '<input class="hidden" type="text" id="storeStart' + colIndex + '" name="storeStart' + colIndex + '" value="" />';
        codeTask += '<input class="hidden" type="text" id="storeStop' + colIndex + '" name="storeStop' + colIndex + '" value="" />';
        codeTask += '<input class="hidden" type="text" id="storeClicks' + colIndex + '" name="storeClicks' + colIndex + '" value="" />';
        codeTask += '<input class="hidden" type="textarea" id="storeMouse' + colIndex + '" name="storeMouse' + colIndex + '" />';
   
      codeTask += '</div>';  
    }
   
  }

  codeTask += '</div>';

  
  
  
  //print feedback textarea, thank you message, and "Submit" button
  codeTask += '<div id="thankyou" class="hidden boxMessage thankyou">';
    codeTask += '<textarea id="feedback" name="feedback" value="" placeholder="Any feedback? Type it here!"></textarea>';
    codeTask += '<p class="message">Click on "Submit", and you are done!</p>';  
    codeTask += '<button class="control red" onclick="var outV = outArray(' + lastCol + '); google.script.run.printResult(outV, \'Tasks\'); hide(\'thankyou\'); show(\'restart\')">Submit</button>';
  codeTask += '</div>';
  
  
  //print restart code
  codeTask += '<div id="restart"  class="hidden" >';
  codeTask += '<a  class="control red" href="' + scriptURL + '?action=measure" style="display:inline-block;">Do it again?</a>';
  codeTask += '<p class="message">The list of tasks is randomly generated every time</p>';
  codeTask += '</div>';
  
  
  return codeTask; 
}





// ******************************************************************************************************
// Function to print a new row of items (pass as the array "outArr") in the spreadsheet (identified by name, passed as "targetSheet")
// ******************************************************************************************************
function printResult(outArr, targetSheet){
  var ss = SpreadsheetApp.openById(sheetID);
  var sheet = ss.getSheetByName(targetSheet);
 
  //lock to avoid concurrent writes 
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
  
  // get column where to print the data
  var rowNum = sheet.getLastRow()+1; 
    
  //print the timestamp
  var timestap = new Date();
  sheet.getRange(rowNum, 1, 1, 1).setValue(timestap);
  
  //print all collected values.
  //it is more efficient to print a single array that to set each value individually
  sheet.getRange(rowNum,2,1,outArr.length).setValues([outArr]);
  
  lock.releaseLock();
  
}




// ******************************************************************************************************
// Function to print all the tasks in the spreadsheet
// ******************************************************************************************************
function submitTask(outArr){
  var ss = SpreadsheetApp.openById(sheetID);
  var sheet = ss.getSheetByName("Tasks");
  
  //lock to avoid concurrent writes 
  var lock = LockService.getPublicLock();
  
  // get column where to print the data
  var lastCol = sheet.getLastColumn(); 
  sheet.insertColumnAfter(lastCol-1); //add a column before the last (where "feedback" is stored)
  
  //print the task description and solution, and set the status as "Submitted"
  var oArr = [];
  sheet.getRange(1,lastCol,1,1).setValue(outArr[0]);
  sheet.getRange(2,lastCol,1,1).setValue(outArr[1]);
  sheet.getRange(3,lastCol,1,1).setValue('Submitted');
    
  //add comment to the first cell
  var noteText = 'Audience : ' + outArr[2];
  if (outArr[3]) {noteText += ' Note : ' + outArr[3];}
  sheet.getRange(1,lastCol,1,1).setNote(noteText);
  
  lock.releaseLock();
}





// ******************************************************************************************************
// Function to print all the menu items as a collapsible list of options when creating a new task
// ******************************************************************************************************
function printOptions(){
  var ss = SpreadsheetApp.openById(sheetID);
  var sSettings = ss.getSheetByName(settingsName);
  var sheets = ss.getSheets();
  var sheetName ='';
  var cellVal = '';
  var outC = '<div id="taskSelect" class="select"><ul>';
  
  //loop over the sheets, and the HTML to be used by the JSTree plugin
  for(var h=0; h<ss.getSheets().length; h++){   
    var sheet = ss.getSheets()[h]; // look at every sheet in spreadsheet   
    
    //check that the sheet is part of the menu structure
    cellVal = sSettings.getRange(h+2, 2, 1,1).getValue();
    if (cellVal == "Primary" || cellVal == "Secondary") {
      var lastRow = sheet.getLastRow(); 
      var lastCol = sheet.getLastColumn();
      sheetName = sheet.getName();
      
      outC += '<li class="lvl1">' + sheetName + '<ul>';
    
      //loop over the columns in the sheet
      for ( var j=1; j<lastCol+1; j++ ){
        outC += '<li class="lvl2">' + sheet.getRange(1, j, 1,1).getValue() + '<ul>';
        
        //loop over the rows
        for ( var i=2; i<lastRow+1; i++ ){
	      cellVal = sheet.getRange(i, j, 1,1).getValue();
          if (cellVal != '') {
  	        outC += '<li class="lvl3">' + cellVal + '</li>';
	      }
	    }
        
        outC += '</ul></li>';
      }
    
      outC += '</ul></li>';
    }  
  }
  
  outC += '</ul></div>';
  
  
  
  //initialise jsTree
  outC += '<script> $(function() { $("#taskSelect").jstree( {';
    outC += '"core" : { "multiple" : false, "themes" : { "dots" : false} },';
    outC += '"plugins" : ["search","wholerow"]';
  outC += '}); }); </script>';

  
  //add CSS
  var primaryColor = sSettings.getRange(4, 7, 1,1).getValue();
  
  outC += '<style>';
    outC += 'body .jstree-default .jstree-wholerow-clicked{ background: ' + primaryColor +'; }';
  outC += '</style>';
  
  return outC;
}




// ******************************************************************************************************
// Function to print all the audience groups as a collapsible list of options when creating a new task
// ******************************************************************************************************
function printAudience(){
  var ss = SpreadsheetApp.openById(sheetID);
  var sSettings = ss.getSheetByName(settingsName);
  var cellVal = '';
  
  //print list (start with "everyone")
  var outC = '<div id="audienceSelect" class="select"><ul>';
  outC += '<li class="lvl1">Everyone<ul>';
  
  var lastRow = sSettings.getLastRow();
  //loop over the rows
  for ( var i=2; i<lastRow+1; i++ ){
    cellVal = sSettings.getRange(i, 4, 1,1).getValue();
    if (cellVal != '') {
      outC += '<li class="lvl2">' + cellVal + '</li>';
    }
  }
  outC += '</ul></li></ul></div>';
  
  
  
  //initialise jsTree
  outC += '<script> $(function() { $("#audienceSelect").jstree( {';
    outC += '"core" : { "multiple" : true, "themes" : { "dots" : false} },';
    outC += '"plugins" : ["wholerow","checkbox"]';
  outC += '}); }); </script>';

  
  return outC;
}




// ******************************************************************************************************
// Function to randomise an array from https://stackoverflow.com/questions/2450954/how-to-randomize-shuffle-a-javascript-array
// use it like: arr = shuffle(arr);
// ******************************************************************************************************
function shuffle(array) {
  var currentIndex = array.length, temporaryValue, randomIndex;

  // While there remain elements to shuffle...
  while (0 !== currentIndex) {

    // Pick a remaining element...
    randomIndex = Math.floor(Math.random() * currentIndex);
    currentIndex -= 1;

    // And swap it with the current element.
    temporaryValue = array[currentIndex];
    array[currentIndex] = array[randomIndex];
    array[randomIndex] = temporaryValue;
  }

  return array;
}