function askGDriveID() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt("Please enter Google Drive ID to create the folders on");
  return result.getResponseText()
}

function addMissingLines(matrix){
  // There're sometimes several folders listed in one cell in the last column.
  // We have to add several row similar to the row, where such cell exists, 
  // except that the last cell in every row contains only one entry

  var new_matrix = []
  for (var x = 0; x < matrix.length; x++){
    var y = matrix[x].length - 1
    if (matrix[x][y].includes("\n")){
      var last_cell = matrix[x].pop()
      var new_cells = last_cell.split("\n")
      for (var j = 0; j < new_cells.length; j++){
        var new_row = []
        for (var i = 0; i < y; i++){
          new_row.push(matrix[x][i])
        }
        new_row.push(new_cells[j])
        new_matrix.push(new_row)  
      }
    } else {
      new_matrix.push(matrix[x])
    }
  }

  return new_matrix
}

function fillInTheGaps(matrix){
  // Somehow I haven't found a way to create a tree data structure, so we have to fill the gaps 
  // in the table, so that every row shows full path to the folder mentioned in the last 
  // cell of the row.

  for (var x = 1; x < matrix.length; x++){
    for (var y = matrix[0].length-1; y >=0 ; y--){
      if (! matrix[x][y] && matrix[x][y+1]){
        matrix[x][y] = matrix[x-1][y]
      }
    }
  }
  return matrix
}

function checkFolders(base_folder, folder_name) {
  if (base_folder){ //Checks if is specified main folder
    var folders = DriveApp.getFolderById(base_folder).getFolders(); //Getting folder list from specified location
  } else { //If not specified folder we take all Drive folders
   var folders = DriveApp.getFolders(); //Getting all folders in Drive
  }

  var folder_exist = false, folder; //By default folder don`t exists till we check and confirm that it exists

  while(folders.hasNext()){ //Looping through all folders in specified location while we have folders which to check
    folder = folders.next(); //Taking next folder to check if it is folder that we are looking for
    if(folder.getName() == folder_name){ //Checking it`s name
     folder_exist = folder.getId(); //If folder is found we assign its ID to our return variable
    } //If folder don`t exists we have already assigned false to return value
  }

  if (!folder_exist) { //Checks if folder don`t exists than we need to create it
    if(base_folder){ //Checking where we should create folder
      DriveApp.getFolderById(base_folder).createFolder(folder_name); //If base folder is specified than we create there folder
    } else {
      DriveApp.createFolder(folder_name); //If base folder not specified than we create folder in Drive root folder
    }
   folder_exist = checkFolders(base_folder, folder_name); //Now run self again to check if folder was created and if it is created than we get folder ID
  }
  return folder_exist; //Returning false if our folder is not created or folder ID if folder exists and is created
}

function makeFolder(parent_id,values,row,column){
  var new_folder_id
  // If the folder named "values[row][column]" hasn't been created then create it
  new_folder_id = checkFolders(parent_id,values[row][column])
  // If values[row][column+1] is defined call makeFolder(new_folder_id,values,row,column+1)
  if (values[row][column+1]) {
    makeFolder(new_folder_id,values,row,column+1)
  }
}

function makeFolderOld(parent_id,values,row,column){
  var new_folder_id
  // If the folder named "values[row][column]" hasn't been created then create it
  if (row>0 && values[row][column] == values[row-1][column]) {
    new_folder_id = DriveApp.getFoldersByName(values[row][column]).next().getId()
  } else {
    new_folder_id = DriveApp.getFolderById(parent_id).createFolder(values[row][column]).getId()
  }
  // If values[row][column+1] is defined call makeFolder(new_folder_id,values,row,column+1)
  if (values[row][column+1]) {
    makeFolder(new_folder_id,values,row,column+1)
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Google Drive')
    .addItem('Create folders', 'menuCreateFolders')
    .addToUi();
}

function menuCreateFolders() {
  
  var values = SpreadsheetApp.getActiveSheet().getDataRange().getValues()
  var gdrive_id = askGDriveID()

  values = fillInTheGaps(values)
  values = addMissingLines(values)

  for (var i = 0; i < values.length; i++){
    makeFolder(gdrive_id,values,i,0)
  }
}