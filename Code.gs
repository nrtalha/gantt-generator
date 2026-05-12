function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Gantt Generator')
      .addItem('Open Generator', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Gantt Generator')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==========================================
// GROQ API KEY
const GROQ_API_KEY = "gsk_YOUR_API_KEY_HERE";
// ==========================================

function calculateBusinessEndDate(startDate, duration, unit) {
  let daysToAdd = 0;
  if (unit === 'Week') daysToAdd = Math.round(duration * 5);
  else if (unit === 'Day') daysToAdd = Math.round(duration);
  else if (unit === 'Month') daysToAdd = Math.round(duration * 20); 

  if (daysToAdd <= 0) return new Date(startDate);
  
  let date = new Date(startDate.getTime());
  let added = 0;
  let targetAdd = daysToAdd - 1; 
  
  while (added < targetAdd) {
    date.setDate(date.getDate() + 1);
    if (date.getDay() !== 0 && date.getDay() !== 6) {
      added++;
    }
  }
  return date;
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  
  if (sheet.getRange(2, 1).getValue() !== "Task Name") return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  if (row > 2 && (col === 2 || col === 4)) {
    redrawAllBars(sheet);
  }
}

function redrawAllBars(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 2 || lastCol <= 4) return;

  const timelineRange = sheet.getRange(3, 5, lastRow - 2, lastCol - 4);
  timelineRange.setBackground("#ffffff");

  const unitStr = sheet.getRange(2, 5).getValue() || "";
  const unit = unitStr.charAt(0) === 'W' ? 'Week' : (unitStr.charAt(0) === 'M' ? 'Month' : 'Day');
  
  const baseDate = new Date(sheet.getRange(3, 2).getValue());
  if (!(baseDate instanceof Date) || isNaN(baseDate.getTime())) return;
  
  const data = sheet.getRange(3, 2, lastRow - 2, 3).getValues(); 

  for (let i = 0; i < data.length; i++) {
    let taskDate = data[i][0];
    let duration = parseFloat(data[i][2]); 
    let hexColor = "#1a73e8"; 

    if (!(taskDate instanceof Date) || isNaN(duration) || duration <= 0) continue;

    
    let endDate = calculateBusinessEndDate(taskDate, duration, unit);
    sheet.getRange(3 + i, 3).setValue(endDate);

    let timeDiff = taskDate.getTime() - baseDate.getTime();
    let indexOffset = 0;
    
    if (unit === 'Week') indexOffset = Math.round(timeDiff / (1000 * 60 * 60 * 24 * 7));
    else if (unit === 'Day') indexOffset = Math.round(timeDiff / (1000 * 60 * 60 * 24)); // Still visually spans calendar days in chart
    else if (unit === 'Month') indexOffset = (taskDate.getFullYear() - baseDate.getFullYear()) * 12 + (taskDate.getMonth() - baseDate.getMonth());
    
    if (indexOffset < 0) indexOffset = 0; 
    let startCol = 5 + indexOffset; 

    let drawDur = Math.ceil(duration);
    if (startCol + drawDur - 1 > lastCol) drawDur = lastCol - startCol + 1;
    if (drawDur > 0 && startCol <= lastCol) {
      sheet.getRange(3 + i, startCol, 1, drawDur).setBackground(hexColor);
    }
  }
}


function generateGantt(projectData, duration, unit, startDateStr) {
  if (!GROQ_API_KEY || !GROQ_API_KEY.startsWith("gsk_")) return "Error: API Key missing or invalid in Code.gs";
  if (!projectData) return "Error: Project Data is empty.";
  if (!duration) return "Error: Duration limit is missing.";

  const url = "https://api.groq.com/openai/v1/chat/completions";
  
  const prompt = `You are a strict Data Parser generating a Gantt chart JSON.
  Data: ${projectData}
  
  CRITICAL JSON RULES: 
  1. Extract exact week numbers for 'startPeriod' based on the text (W0 = 0, W16 = 16).
  2. For 'duration', output ONLY decimals. If it's a fraction of a ${unit} (e.g., 2 days in weeks), calculate the decimal (e.g., 0.29). NO mathematical formulas.
  3. All keys and string values MUST be wrapped in standard double quotes for valid JSON syntax.
  4. Replace slashes (/), backslashes, and internal quotes within the task names with spaces (e.g., rename "UI/UX" to "UI UX") to prevent escaping errors.
  
  Output strictly as a JSON object containing an array. No conversational text. 
  Schema: {"tasks": [{"taskName": "string", "duration": number, "startPeriod": number}]}`;

  const payload = {
    "model": "llama-3.3-70b-versatile",
    "messages": [{"role": "user", "content": prompt}],
    "temperature": 0.0, 
    "max_tokens": 2048,
    "response_format": {"type": "json_object"} 
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {"Authorization": "Bearer " + GROQ_API_KEY.trim()}, 
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) return "Groq API Error: " + response.getContentText();

    let rawContent = JSON.parse(response.getContentText()).choices[0].message.content;
    
    let jsonData;
    try {
      jsonData = JSON.parse(rawContent);
    } catch(err) {
      return "JSON Syntax Error from LLM: " + err.message + "\n\nRaw Text: " + rawContent.substring(0, 200);
    }

    let tasks = jsonData.tasks;
    if (!tasks || !Array.isArray(tasks)) return "Error: LLM did not return the tasks array correctly.";

    const dateParts = startDateStr.split('-');
    // Adjust global start date to a Monday if it happens to land on a weekend
    let currentDate = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
    if (currentDate.getDay() === 0) currentDate.setDate(currentDate.getDate() + 1); // Sunday -> Monday
    if (currentDate.getDay() === 6) currentDate.setDate(currentDate.getDate() + 2); // Saturday -> Monday

    let colDates = [];
    
    for(let i = 0; i < duration; i++) {
      colDates.push(new Date(currentDate));
      if(unit === 'Week') currentDate.setDate(currentDate.getDate() + 7);
      if(unit === 'Day') currentDate.setDate(currentDate.getDate() + 1); // Note: Columns show raw calendar for spacing
      if(unit === 'Month') currentDate.setMonth(currentDate.getMonth() + 1);
    }

    let monthRow = ["", "", "", ""]; 
    let weekRow = ["Task Name", "Start Date", "End Date", "Duration"]; 
    
    let currentMonthTracker = "";
    colDates.forEach((d, index) => {
      let m = d.toLocaleString('default', { month: 'short', year: 'numeric' });
      if(m !== currentMonthTracker) {
        monthRow.push(m);
        currentMonthTracker = m;
      } else {
        monthRow.push(""); 
      }
      let unitPrefix = unit.substring(0, 1).toUpperCase();
      weekRow.push(unitPrefix + index); 
    });

    let dataMatrix = [monthRow, weekRow];
    let colorMatrix = [
      Array(monthRow.length).fill("#f1f3f4"), 
      Array(weekRow.length).fill("#e8eaed")   
    ];

    tasks.forEach((task) => {
       let periodIndex = task.startPeriod;
       if (periodIndex < 0) periodIndex = 0;
       if (periodIndex >= colDates.length) periodIndex = colDates.length - 1;
       
       let exactTaskDate = colDates[periodIndex];
       
       // Calculate End Date strictly considering Business Days
       let endDate = calculateBusinessEndDate(exactTaskDate, task.duration, unit);

       let rowData = [task.taskName, exactTaskDate, endDate, task.duration];
       let rowColors = ["#ffffff", "#ffffff", "#ffffff", "#ffffff"]; 

       for (let i = 0; i < duration; i++) {
         rowData.push(""); 
         if (i >= task.startPeriod && i < task.startPeriod + Math.ceil(task.duration)) {
           rowColors.push("#1a73e8"); 
         } else {
           rowColors.push("#ffffff"); 
         }
       }
       dataMatrix.push(rowData);
       colorMatrix.push(rowColors);
    });

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.clear(); 
    sheet.getDataRange().clearDataValidations(); 
    
    const numRows = dataMatrix.length;
    const numCols = dataMatrix[0].length;
    
    const range = sheet.getRange(1, 1, numRows, numCols);
    range.setValues(dataMatrix);          
    range.setBackgrounds(colorMatrix);    

    sheet.getRange(1, 1, 2, numCols).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(1, 1, numRows, 4).setHorizontalAlignment("center"); 
    sheet.getRange(1, 1, numRows, 1).setHorizontalAlignment("left"); 
    range.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
    
    sheet.getRange(3, 2, numRows - 2, 2).setNumberFormat("mmm d, yyyy");

    sheet.setColumnWidth(1, 220); 
    sheet.setColumnWidth(2, 110); 
    sheet.setColumnWidth(3, 110); 
    sheet.setColumnWidth(4, 70);  
    sheet.setColumnWidths(5, numCols - 4, 35); 

    return "Success! W0 Business-Day Gantt Generated.";

  } catch (e) {
    return "Script Crash: " + e.message;
  }
}
