# Excel Importer for VB.Net Windows Forms Application

This application allows you to import Excel files and sheets into a VB.Net Windows Forms Application. You can then import the data from the selected sheet into an SQL Server database.

## Requirements
- Microsoft Visual Studio (or any compatible IDE) with VB.Net support.
- .NET Framework (version X.X or later).
- An SQL Server instance accessible from your application.

## Features
- Import Excel files (.xls, .xlsx) into the application.
- Select specific sheets from the imported Excel files.
- Display the data from the selected sheet in a DataGridView on the form.
- Import the displayed data into an SQL Server database.

## Getting Started
1. Clone or download the repository to your local machine.
2. Open the solution file (`ExcelImporter.sln`) in Microsoft Visual Studio.
3. Build the solution to restore NuGet packages and compile the project.
4. Configure your SQL Server connection string in the application code.
5. Run the application to start importing Excel files and sheets.

## Usage
1. Launch the application.
2. Click the Browse button to select an Excel file.
3. Choose the desired sheet from the Excel file using the dropdown menu.
4. The data from the selected sheet will be displayed in the DataGridView.
5. Click the Import button to import the displayed data into the SQL Server database.

## Troubleshooting
If you encounter any issues or errors while using the application, please refer to the following troubleshooting steps:
- Check SQL Server Connection: Ensure that your SQL Server instance is running and accessible from the application.
- Verify Excel File: Make sure the selected Excel file contains valid data and is not corrupted.
- Review Application Code: Double-check the application code to ensure that the SQL Server connection string is correctly configured.
- Error Logging: Add error handling and logging to the application code to capture detailed error messages and facilitate debugging.

## Contributing
Contributions to this project are welcome! If you encounter any bugs, have feature requests, or want to contribute improvements, feel free to open an issue or submit a pull request on GitHub.

## License
This project is licensed under the MIT License.
