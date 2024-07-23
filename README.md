# website
This is a basic website for beginners
import { sp } from "@pnp/sp";
import * as XLSX from "xlsx";

private async onButtonClick() {
  // Create a file input element
  const fileInput = document.createElement('input');
  fileInput.type = 'file';

  // Trigger the file dialog
  fileInput.click();

  // Handle file selection
  fileInput.onchange = async (event) => {
    const files = (event.target as HTMLInputElement).files;
    if (files && files.length > 0) {
      const file = files[0];
      const data = await this.readExcelFile(file);

      if (data) {
        await this.updateSharePointList(data);
        console.log('SharePoint list updated successfully');
      }
    }
  };
}

private readExcelFile(file: File): Promise<any[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e: any) => {
      const binaryStr = e.target.result;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(worksheet);
      resolve(json);
    };
    reader.onerror = (error) => reject(error);
    reader.readAsBinaryString(file);
  });
}

private async updateSharePointList(data: any[]) {
  for (const item of data) {
    await sp.web.lists.getByTitle("YourListTitle").items.add(item);
  }
}
