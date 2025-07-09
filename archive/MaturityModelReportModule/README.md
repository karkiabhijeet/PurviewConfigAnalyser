# Maturity Model Report Module

## Overview
The Maturity Model Report Module is a PowerShell module designed to evaluate maturity models based on specified configurations in JSON and XML formats. It generates comprehensive reports in both Excel and JSON formats, clearly indicating which parts of the maturity model are met and which are not.

## Features
- Read and parse JSON configuration files.
- Read and parse XML configuration files for maturity models.
- Evaluate the maturity model against the provided configurations.
- Generate reports in Excel format.
- Generate reports in JSON format.

## Installation
To install the Maturity Model Report Module, clone the repository and import the module in your PowerShell session:

```powershell
Import-Module "path\to\MaturityModelReportModule\src\MaturityModelReport.psm1"
```

## Usage
1. **Read JSON Configuration**: Use the `Read-JsonFile` function to load your JSON configuration.
   ```powershell
   $jsonData = Read-JsonFile -Path "path\to\your\input.json"
   ```

2. **Read XML Configuration**: Use the `Read-XmlConfig` function to load your XML maturity model.
   ```powershell
   $xmlData = Read-XmlConfig -Path "path\to\your\maturity-model.xml"
   ```

3. **Evaluate Maturity Model**: Use the `Evaluate-MaturityModel` function to assess the maturity model against the loaded configurations.
   ```powershell
   $evaluationResults = Evaluate-MaturityModel -JsonData $jsonData -XmlData $xmlData
   ```

4. **Export Reports**:
   - To export the evaluation results to an Excel file:
     ```powershell
     Export-ReportToExcel -Results $evaluationResults -OutputPath "path\to\output.xlsx"
     ```
   - To export the evaluation results to a JSON file:
     ```powershell
     Export-ReportToJson -Results $evaluationResults -OutputPath "path\to\output.json"
     ```

## Examples
- Sample input JSON file: [examples/sample-input.json](examples/sample-input.json)
- Sample maturity model XML file: [examples/sample-maturity-model.xml](examples/sample-maturity-model.xml)

## Testing
Unit tests are provided to ensure the functionality of the module. You can run the tests using the following command:
```powershell
Invoke-Pester -Path "path\to\MaturityModelReportModule\tests"
```

## Contributing
Contributions are welcome! Please submit a pull request or open an issue for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for details.