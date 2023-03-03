import { Component } from '@angular/core';

import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
  selector: 'app-sheet',
  templateUrl: './sheet.component.html',
})
export class SheetJSComponent {
  // Initialize data with a default value
  data: AOA = [
    [1, 2],
    [3, 4],
  ];
  // Set options for writing Excel files
  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  // Set the default file name
  fileName: string = 'SheetJS.xlsx';

  // This method is called when a file is selected in the input element
  onFileChange(evt: any) {
    /* wire up file reader */
    // Get the file input element
    const target: DataTransfer = <DataTransfer>evt.target;
    // Make sure only one file is selected
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    // Create a FileReader object to read the file
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      // Get the binary data from the FileReader object
      const bstr: string = e.target.result;
      // Parse the binary data into a Workbook object using the XLSX library
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      /* grab first sheet */
      // Get the name of the first worksheet in the workbook
      const wsname: string = wb.SheetNames[0];
      // Get the Worksheet object for the first worksheet
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      // Convert the Worksheet object to a JavaScript object using the XLSX library
      // and assign it to the `data` property of the component
      this.data = <AOA>XLSX.utils.sheet_to_json(ws, { range: 1, header: 1 });
      // Log the resulting data to the console for debugging purposes
      console.log(this.data);
    };
    // Read the selected file as a binary string using the FileReader object
    reader.readAsBinaryString(target.files[0]);
  }
}
