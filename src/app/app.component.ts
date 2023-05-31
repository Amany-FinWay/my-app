import { Component } from '@angular/core';
import { AgChartOptions } from 'ag-charts-community';
import * as XLSX from 'xlsx';
const { utils } = XLSX;
@Component({
  selector: 'app-root',
  template: `<ag-charts-angular
  [options]="options"
></ag-charts-angular>

<ag-charts-angular
  [options]="options"
></ag-charts-angular>
<ag-charts-angular
  [options]="options"
></ag-charts-angular>
<ag-charts-angular
  [options]="options"
></ag-charts-angular>
<ag-charts-angular
  [options]="options"
></ag-charts-angular>
<ag-charts-angular
  [options]="options"
></ag-charts-angular>
<ag-charts-angular
  [options]="options"
></ag-charts-angular>
<div style=" margin: auto; width: 50%;">
  <button (click)="exportExcel()">Export to Excel</button>
  <button (click)="exportExcel2()">Export to Excel2</button>
  <button (click)="exportToExcel3()">Export to Excel</button>


  <table id="excel-table">
    <tr>
      <th>Id</th>
      <th>Name</th>
      <th>Username</th>
      <th>Email</th>
    </tr>
    <tr *ngFor="let item of userList">
      <td>{{item.id}}</td>
      <td>{{item.name}}</td>
      <td>{{item.username}}</td>
      <td>{{item.email}}</td>
    </tr>
  </table>
</div>
`
,
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'my-app';

  options: AgChartOptions;
  userList = [

    {

    "id": 1,

    "name": "Leanne Graham",

    "username": "Bret",

    "email": "Sincere@april.biz"

    },

    {

    "id": 2,

    "name": "Ervin Howell",

    "username": "Antonette",

    "email": "Shanna@melissa.tv"

    },

    {

    "id": 3,

    "name": "Clementine Bauch",

    "username": "Samantha",

    "email": "Nathan@yesenia.net"

    },

    {

    "id": 4,

    "name": "Patricia Lebsack",

    "username": "Karianne",

    "email": "Julianne.OConner@kory.org"

    },

    {

    "id": 5,

    "name": "Chelsey Dietrich",

    "username": "Kamren",

    "email": "Lucio_Hettinger@annie.ca"

    }

    ];

    data = [
      {
        name: 'John',
        age: '30',
        hobbies: [
          { name: 'Reading', type: 'Indoor' },
          { name: 'Hiking', type: 'Outdoor' },
        ]
      },
      {
        name: 'Jane',
        age: '25',
        hobbies: [
          { name: 'Swimming', type: 'Outdoor' },
          { name: 'Swimming', type: 'Outdoor' }
        ]
      },
      {
        name: 'Haidy',
        age: '25',
        hobbies: [
          { name: 'Swimming', type: 'Outdoor' },
          { name: 'Swimming', type: 'Outdoor' },
          { name: 'Swimming', type: 'Outdoor' },
        ]
      }
    ];

    fileName= 'ExcelSheet.xlsx';

  constructor() {

    this.options = {
      title: {
        text: 'CIB',
        fontSize: 35
      },
      data: [
        { lable: 'Closed Tickets', value: 56 },
        { lable: 'Opend Tickets', value: 22 },
        { lable: 'Suspended Tickets', value: 6 },
        { lable: 'Cancelled Tickets', value: 8 }
      ],
      series: [
        {
          type: 'pie',
          angleKey: 'value',
          calloutLabelKey: 'label',
          sectorLabelKey: 'value',
          sectorLabel: {
            color: 'white',
            fontWeight: 'bold',
            fontSize: 25
          },
        },
      ],
    };
  }

  exportExcel2(){
//-------------------------------------------works--------------------------------------------------------------------

    const worksheet = XLSX.utils.aoa_to_sheet([[]]);

    const allDate: any[] = [];
    this.data.forEach(x => {
      x.hobbies.forEach((y, index) => {
        allDate.push({
          name: index == 0 ? x.name : '',
          age: index == 0 ? x.age : '',
          hobbyName: y.name,
          hobbyType: y.type
        });
      });
    });

    const arrayOfArrays = allDate.map(obj => Object.values(obj));
    arrayOfArrays.unshift(['Name', 'Age', 'Hobby Name', 'Hobby Type']);

  for (let i = 2; i < arrayOfArrays.length; i++) {
    if (arrayOfArrays[i][0] !== '' || arrayOfArrays[i][1] !== '') {
      arrayOfArrays.splice(i, 0, []);
      i++;
    }
  }

const workbook = XLSX.utils.book_new();

XLSX.utils.sheet_add_aoa(worksheet, arrayOfArrays, { origin: 0 }); // 0

XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

XLSX.writeFile(workbook, 'path/to/your/new/file.xlsx');

    //const worksheet = XLSX.utils.json_to_sheet(allDate);

    //XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    //XLSX.writeFile(workbook, 'data.xlsx');
 //---------------------------------------------------------

// const workbook = XLSX.utils.book_new();
// const worksheet = XLSX.utils.aoa_to_sheet([[]]);

// const roworksheet: any[][] = [  ['Header 1', 'Header 2', 'Header 3'],
//   ['Data 1', 'Data 2', 'Data 3'],
//   ['Data 4', 'Data 5', 'Data 6'],
//   ['Data 4', 'Data 5', 'Data 6'],
// ];

//sconsole.log(roworksheet);

// for (let i = roworksheet.length - 1; i >= 0; i--) {
//   roworksheet.splice(i + 1, 0, []);
// }

// XLSX.utils.sheet_add_aoa(worksheet, roworksheet, { origin: 0 });

// XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// XLSX.writeFile(workbook, 'path/to/your/new/file.xlsx');



//-----------------------------------------------------------
  }


  exportExcel() {

    // Create a new workbook
    const workbook = XLSX.utils.book_new();

    // Create a new worksheet
    const worksheet = XLSX.utils.json_to_sheet(this.userList);

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    // Set the column widths
    const columnWidths = [
      {wch: 10},
      {wch: 20},
      {wch: 30},
      {wch: 40},
      {wch: 50},
    ];
    worksheet['!cols'] = columnWidths;

    // Set the cell styles
    const headerStyle = {
      fill: {fgColor: {rgb: '0000FF'}},
      font: {bold: false}
    };
    const dataStyle = {
      fill: {fgColor: {rgb: 'FFFFFF'}},
      font: {bold: false}
    };
    const headerRange = XLSX.utils.decode_range(worksheet['!ref']!);
    for (let i = headerRange.s.r; i <= headerRange.e.r; i++) {
      for (let j = headerRange.s.c; j <= headerRange.e.c; j++) {
        const cellAddress = XLSX.utils.encode_cell({r: i, c: j});
        const cell = worksheet[cellAddress];
        if (cell && cell.t === 's' && i === headerRange.s.r) {
          cell.s = headerStyle;
        } else {
          cell.s = dataStyle;
        }
      }
    }
    // Export the workbook
    XLSX.writeFile(workbook, 'data.xlsx');
  }

  styleRow(sheet: XLSX.WorkSheet, rowIndex: number, color: string) {
    const range = XLSX.utils.decode_range(sheet['!ref']!);
    if (rowIndex < range.s.r || rowIndex > range.e.r) {
      return;
    }
    for (let i = range.s.c; i <= range.e.c; i++) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: i });
      const cell = sheet[cellAddress];
      if (cell) {
        const style = cell.s || {};
        style.fill = { type: 'pattern', patternType: 'solid', fgColor: { rgb: color } };
        sheet[cellAddress].s = style;
      }
    }
  }

  exportToExcel3(){
    const data = [
      ['Name', 'Age', 'Gender'],
      ['John', 30, 'Male'],
      ['Jane', 25, 'Female'],
      ['Bob', 40, 'Male']
    ];

 // Create a workbook object
 const workbook = XLSX.utils.book_new();

 // Create a worksheet object using the array of arrays
 const worksheet = XLSX.utils.aoa_to_sheet(data);

 // Set the fill color of the first cell
 const cell = worksheet['A1'];
 cell.s = { fill: { bgColor: { rgb: 'FFFF0000' } } };

 // Add the worksheet to the workbook
 XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

 // Convert the workbook to an ArrayBuffer
 const buffer = XLSX.write(workbook, { type: 'array' });

 // Create a Blob object from the ArrayBuffer
 const blob = new Blob([buffer], { type: 'application/octet-stream' });

 // Generate a download link and click it
 const link = document.createElement('a');
 link.href = window.URL.createObjectURL(blob);
 link.download = 'data.xlsx';
 link.click();
  }

}
