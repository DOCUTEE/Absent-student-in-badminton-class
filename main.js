const axios = require('axios');
const XLSX = require('xlsx');

async function downloadXLSX(url) {
    const response = await axios.get(url, { responseType: 'arraybuffer' });
    return response.data;
}

async function convertXLSXToJSON(data, sheetIndex) {
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[sheetIndex];
    const worksheet = workbook.Sheets[sheetName];
    // console.log(sheetName);
    const json = XLSX.utils.sheet_to_json(worksheet);
    return json;
}
async function getAbsentStudentList() {
    const url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQTcHsRqB6m-hEIK_HXg6dMxiAl0TqThlhzbEDz2IOqHttOJLfAGnZXCM0XsR-7wEfaAoCSJzUeun3N/pub?output=xlsx';
    const week = 1;
    const strWeek = "Tuần " + week.toString();
    console.log(strWeek);
    const MSSV = "MSSV";
    const HoVaTen = "Họ và tên";
    const strStudentExist = "Có mặt";
    let absentList = {
        "Week": strWeek,
        "teamList": [

        ]
    }
    // console.log(absentList);
    try {
        const xlsxData = await downloadXLSX(url);
        for (let i = 0; i < 5; i++) {
            const jsonData = await convertXLSXToJSON(xlsxData, i);
            let team = {
                "Name": "Nhóm " + (i + 1).toString(),
                "List": []
            };
            jsonData.forEach(student => {
                if (student[strWeek] != strStudentExist) {
                    let studentInfo = {
                        "MSSV": student[MSSV],
                        "Name": student[HoVaTen],
                        "Reason": student[strWeek]
                    };
                    team["List"].push(studentInfo);
                }
            });
            absentList["teamList"].push(team);
        }
    } catch (error) {
        console.error('Error downloading or converting the file:', error);
    }
    return absentList;
}

async function exportAbsentStudentList() {
    const absentJSON = await getAbsentStudentList();
    //   console.log(JSON.stringify(absentJSON,null,2));
    // let absentListDiv = document.getElementById('absentStudentList');
    // absentListDiv.innerHTML = '';
    // Object.entries(absentJSON["teamList"]).forEach(team => {
    //     // let teamUL = document.createElement('ul');
    //     // console.log(team);
    //     if (team['List'] != null)
    //     Object.entries(team['List']).forEach(student => {
    //         // let studentLI = document.createElement('li');
    //         // studentLI.innerHTML = student["MSSV"] + " | " + student["Name"] + " | " + student["Reason"];
    //         // teamUL.appendChild(studentLI);
    //         console.log(student["MSSV"] + " | " + student["Name"] + " | " + student["Reason"]);
    //     });
    //     // absentListDiv.appendChild(teamUL);
    // })
    console.log(Object.entries(absentJSON["teamList"]));
}

exportAbsentStudentList();