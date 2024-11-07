import * as React from "react";
import { PrimaryButton } from "@fluentui/react";
import * as ExcelJS from "exceljs";
import { _context, retrieveRecords } from "../services/DataverseService";

export interface IButtonProps {
  // Define your props here
}

export const Button: React.FC<IButtonProps> = (props) => {
  return (
    <div style={{ display: "flex", justifyContent: "end", alignItems: "center", height: "100%", width: "100%" }}>
      <PrimaryButton
        text="Export to excel"
        onClick={onClick}
        style={{
          display: "flex",
          justifyContent: "center",
          alignItems: "center",
          width: "100%",
          textAlign: "center",
        }}
      ></PrimaryButton>
    </div>
  );
};

const onClick = async () => {
  const data = _context.parameters.sampleDataSet;
  const records = composeExcelArray(data);

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sample Sheet");
  const excelRows = [data.columns.map((column) => column.displayName)];

  records.forEach((record) => {
    const row: Array<string> = [];
    data.columns.forEach((column) => {
      switch (column.dataType) {
        case "SingleLine.Text":
          row.push((record.getValue(column.name) as string) || "");
          break;
        case "Lookup.Simple":
          row.push((record.getValue(column.name) as ComponentFramework.LookupValue)?.name || "");
          break;
        case "Number":
          row.push((record.getValue(column.name) as number).toString() || "");
          break;
        case "DateAndTime.DateOnly":
          row.push((record.getValue(column.name) as Date)?.toString() || "");
          break;
        case "DateAndTime.DateAndTime":
          row.push((record.getValue(column.name) as Date)?.toString() || "");
          break;
        case "OptopnSet":
          row.push((record.getValue(column.name) as number).toString() || "");
          break;
        case "TwoOptions":
          row.push((record.getValue(column.name) as boolean) ? "Yes" : "No");
          break;
        case "Currency":
          row.push((record.getValue(column.name) as number)?.toString() || "");
          break;
      }
    });
    excelRows.push(row);
  });

  excelRows.forEach((row) => {
    worksheet.addRow(row);
  });

  const buffer = await workbook.xlsx.writeBuffer();

  // Create a Blob from the buffer
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  // Create a temporary link element to trigger the download
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "sample.xlsx"; // Set the desired file name
  document.body.appendChild(a); // Append the link to the body
  a.click(); // Trigger the download
  document.body.removeChild(a); // Clean up the link element
  URL.revokeObjectURL(url);
};

const composeExcelArray = (dataSet: ComponentFramework.PropertyTypes.DataSet): Array<ComponentFramework.PropertyHelper.DataSetApi.EntityRecord> => {
  const data: Array<ComponentFramework.PropertyHelper.DataSetApi.EntityRecord> = [];

  dataSet.sortedRecordIds.forEach((recordId) => {
    const record = dataSet.records[recordId];
    data.push(record);
  });

  return data;
};
