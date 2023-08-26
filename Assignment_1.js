/*
PLEASE -->first : download this package : npm install xlsx 
       -->second : download the NodeJs_course_1 in this form .xlsx
       --third : create a new excel file to save the result in it , like :TestJS.xlsx
       --forth : put all the files in one folder and open it in VS Code , and start testing !                           
*/
// Require Modules.
const fs = require("fs");
const xlsx= require("xlsx");

//Read the origin file 
const ReadOriginal = xlsx.readFile("NodeJs_Test.xlsx" , {cellText: true});

//test some operations :
console.log(ReadOriginal.SheetNames); //The output : ['Sheet1'] .

//Save the sheet/tab we want to work on it :
var WSheet=ReadOriginal.Sheets['Sheet1'] ;
//console.log(WSheet); // shows info. about the sells [type ,value ,...] 

/*
perform CRUD operations on it --->>>
 1)The "Read" operation , like : show tha data stored in the excel file 
*/
var JsonData = xlsx.utils.sheet_to_json(WSheet) ;
//console.log(JsonData); //print the datafor each student 

/*
 2)The "Create" operation , like : adding new row , ...
*/
//because the original excel sheet , has 50 students number in it , so when we create new lines , we start from id=51
const new_row_1={
    "Student Number":51 ,
    "Student Email": 'email@email.com',
    "Phone Number" :'0000000000' ,
    "University" :'JUST' ,
    "Specialization" :'CS' ,
    "Codewars": 'https://www.codewars.com' ,
    "Github" : 'https://github.com'
};

const new_row_2={
    "Student Number":52 ,
    "Student Email": 'test@test.com',
    "Phone Number" :'999999999' ,
    "University" :'JU' ,
    "Specialization" :'CIS' ,
    "Codewars": 'https://www.codewars.com' ,
    "Github" : 'https://github.com'
};

JsonData.push(new_row_1);
JsonData.push(new_row_2);

//this method for modifying existing data on the student numbers from 1 to 50
/* 3) The "Update" operation , like : Modifying data in rows , ...
update the last row (41)
-FIRST METHOD for "Update" multiple data in the same row ---->>>> */
const row_No_41 = JsonData.find(row => row["Student Number"] === 41);
if (row_No_41) {
    Object.assign(row_No_41, {
        "Student Email": "example@example.com",
        "Phone Number": "1111111111",
        "University": "YU",
        "Specialization": "CPE",
        "Codewars": "https://www.codewars.com",
        "Github": "https://github.com"
    });
}

//-SECOND METHOD for "Update" single data---->>>>
var updated_row_No = JsonData[40]; // (40) because numbering starts from 0 
updated_row_No["Phone Number"] = '07999999999';


/* 4) The "Delete" operation 
-FIRST METHOD for "Delete" , it's removes the whole row with also the "Student Number"  ---->>>> */
var RowNumber = 40; 
JsonData.splice(RowNumber, 1);

//-SECOND METHOD for "Delete"  certain row without "Student Number" col. ---->>>> */
const row_to_delete = 52;
const index = JsonData.findIndex(row => row["Student Number"] === row_to_delete);
if (index !== -1) {
    const studentNo = JsonData[index]["Student Number"]; 
    JsonData.splice(index, 1); 
    const deletedRow_52 = { "Student Number": studentNo };
    JsonData.push(deletedRow_52);
}

// Convert JsonData to .xlsx again 
var ExcelData = xlsx.utils.json_to_sheet(JsonData);
ReadOriginal.Sheets['Sheet1'] = ExcelData ;

// Write the new data into new .xlsx file
xlsx.writeFile(ReadOriginal , 'TestJS.xlsx');
console.log("All The CRUD operations Done Succesfully !");
