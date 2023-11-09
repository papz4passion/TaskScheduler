const XLSX = require('xlsx');
const { DateTime } = require('luxon');
const { writeFileSync } = require('fs');
const moment = require('moment');

function parseExcelDate(serial) {
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  return new Date(utc_value * 1000);
}

function parseSheet(filePath, sheetName) {
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error(`Sheet ${sheetName} not found.`);
  }
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const headers = json.shift();
  return json.map((row) => {
    return headers.reduce((acc, header, index) => {
      acc[header] = row[index] || null;
      return acc;
    }, {});
  }).filter(task => task[headers[0]] != null);
}

function assignTasksToDevelopers(tasks, developers, startDate, endDate, developerLeaves) {
  const dateRange = [];

  let currentDate = moment(parseExcelDate(startDate));
  endDate = moment(parseExcelDate(endDate));
  while (currentDate.isSameOrBefore(endDate)) {
    dateRange.push(currentDate.format('YYYY-MM-DD'));
    currentDate.add(1, 'days');
  }

  const developerTaskMap = developers.map(dev => ({
    name: dev,
    tasks: dateRange.reduce((acc, date) => {
      acc[date] = [];
      return acc;
    }, {})
  }));

  for (let date of dateRange) {
    for (let dev of developerTaskMap) {
      if (developerLeaves[dev.name] && developerLeaves[dev.name].includes(date)) {
        continue; // Skip if developer is on leave
      }

      let hoursAssigned = 0;
      while (hoursAssigned < 8 && tasks.length > 0) {
        const task = tasks.shift(); // Get the task with the highest priority
        if (hoursAssigned + task.Estimate <= 8) {
          dev.tasks[date].push(task);
          hoursAssigned += task.Estimate;
        } else {
          tasks.unshift(task); // Put the task back if it doesn't fit in the current day
          break;
        }
      }
    }
  }

  return developerTaskMap;
}

function main(filePath) {
  const tasksSheetJson = parseSheet(filePath, 'Sheet1');
  const developersSheetJson = parseSheet(filePath, 'Sheet2');

  const tasks = tasksSheetJson
    .filter(task => task.Task && !isNaN(task.Estimate) && task.TaskOrder)
    .sort((a, b) => a.TaskOrder - b.TaskOrder); // Sort tasks by priority

  const developers = developersSheetJson.map(dev => dev.Developer);
  const startDate = developersSheetJson[0].SprintStartDate;
  const endDate = developersSheetJson[0].SprintEndDate;
  const developerLeaves = developersSheetJson.reduce((acc, dev) => {
    acc[dev.Developer] = dev.Leaves ? dev.Leaves.split(',').map(date => date.trim()) : [];
    return acc;
  }, {});

  const assignedTasks = assignTasksToDevelopers(tasks, developers, startDate, endDate, developerLeaves);

  // Convert the developerTaskMap to a worksheet
  const output = XLSX.utils.book_new();
  assignedTasks.forEach(dev => {
    const ws = XLSX.utils.json_to_sheet(Object.entries(dev.tasks).map(([date, tasks]) => ({
      Date: date,
      Tasks: tasks.map(task => task.Task).join(', '),
      Hours: tasks.reduce((sum, task) => sum + task.Estimate, 0)
    })), { header: ["Date", "Tasks", "Hours"] });
    XLSX.utils.book_append_sheet(output, ws, dev.name);
  });

  XLSX.writeFile(output, 'output.xlsx');
}

// Get the command line argument for the file path
const filePath = process.argv[2];
main(filePath);
