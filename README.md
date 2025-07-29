# Excel VBA Test Case to Feature File Converter

This VBA macro automates the process of converting test cases from an Excel spreadsheet into formatted feature files suitable for automation frameworks.

## Features

- **Data Extraction**: Extracts test case data from specified Excel columns
- **Feature File Generation**: Creates individual `.feature` files for each test case
- **CSV Export**: Optional CSV export functionality
- **Error Handling**: Comprehensive error handling and validation
- **Customizable**: Easy column mapping customization
- **User-Friendly**: Interactive directory selection and progress feedback

## Excel Data Structure Requirements

Your Excel sheet should have the following structure:

| Column A | Column B | Column C | Column D |
|----------|----------|----------|----------|
| Test Case ID | Scenario Name | Steps | Expected Result |
| TC001 | Login Test | Enter username and password | User should be logged in successfully |
| TC002 | Search Test | Enter search term and click search | Search results should be displayed |

**Important Notes:**
- Headers should be in Row 1
- Test case data should start from Row 2
- Empty rows will be skipped automatically

## Installation Instructions

### Step 1: Open Excel VBA Editor
1. Open your Excel workbook
2. Press `Alt + F11` to open the Visual Basic Editor
3. In the Project Explorer (left panel), right-click on your workbook name
4. Select **Insert** → **Module**

### Step 2: Copy the VBA Code
1. In the new module window, copy and paste the entire code from `TestCaseToFeatureConverter.bas`
2. Press `Ctrl + S` to save the module
3. Close the VBA Editor (`Alt + Q`)

### Step 3: Enable Macros (if needed)
1. Go to **File** → **Options** → **Trust Center**
2. Click **Trust Center Settings**
3. Select **Macro Settings**
4. Choose **Enable all macros** (or **Disable all macros with notification**)
5. Click **OK**

## Usage Instructions

### Method 1: Using VBA Editor
1. Open the VBA Editor (`Alt + F11`)
2. In the Project Explorer, double-click the module you created
3. Place your cursor inside any of the main functions:
   - `ConvertTestCasesToFeatureFiles()`
   - `ExportTestCasesToCSV()`
   - `ValidateSheetStructure()`
   - `ShowInstructions()`
4. Press `F5` to run the selected function

### Method 2: Using Excel Ribbon (Recommended)
1. Go to **File** → **Options** → **Customize Ribbon**
2. In the right panel, check **Developer** tab
3. Click **OK**
4. Go to the **Developer** tab
5. Click **Macros**
6. Select the desired macro and click **Run**

### Method 3: Using Quick Access Toolbar
1. Go to **File** → **Options** → **Quick Access Toolbar**
2. In the left panel, select **Macros**
3. Select the desired macro and click **Add** → **OK**
4. The macro will now appear in your Quick Access Toolbar

## Available Functions

### 1. `ConvertTestCasesToFeatureFiles()`
**Main function** - Converts Excel test cases to individual feature files.

**Process:**
1. Prompts for output directory selection
2. Reads test case data from Excel
3. Generates feature content in Gherkin format
4. Creates individual `.feature` files for each test case
5. Shows completion message

**Output Format:**
```gherkin
Feature: [ScenarioName]

  Scenario: [ScenarioName] - [TestCaseID]
    Given the test setup is ready
    When I perform the following steps:
      [Steps]
    Then I expect:
      [ExpectedResult]
```

### 2. `ExportTestCasesToCSV()`
**Optional function** - Exports test cases to CSV format.

**Process:**
1. Prompts for output directory selection
2. Creates a single CSV file with all test cases
3. Includes headers and properly escaped data

### 3. `ValidateSheetStructure()`
**Validation function** - Checks if your Excel sheet has the correct structure.

**Checks:**
- Presence of data in expected columns
- Warns about missing or empty columns
- Provides feedback on sheet structure

### 4. `ShowInstructions()`
**Help function** - Displays usage instructions and setup requirements.

## Customization Options

### Changing Column Mappings
To modify which columns contain your data, edit these constants at the top of the module:

```vba
Private Const COL_TEST_CASE_ID As Integer = 1      ' Column A
Private Const COL_SCENARIO_NAME As Integer = 2     ' Column B
Private Const COL_STEPS As Integer = 3             ' Column C
Private Const COL_EXPECTED_RESULT As Integer = 4   ' Column D
```

**Example:** If your Test Case ID is in Column C instead of Column A:
```vba
Private Const COL_TEST_CASE_ID As Integer = 3      ' Column C
```

### Changing Data Start Row
If your data doesn't start from Row 2, modify:
```vba
Private Const START_ROW As Integer = 2             ' Data starts from row 2
```

### Modifying Feature File Format
To change the Gherkin format, edit the `GenerateFeatureContent()` function:

```vba
Private Function GenerateFeatureContent(testCaseID As String, scenarioName As String, steps As String, expectedResult As String) As String
    ' Customize the format here
    content = "Feature: " & scenarioName & vbCrLf & vbCrLf
    ' ... rest of the formatting
End Function
```

## Error Handling

The macro includes comprehensive error handling for:
- Missing or empty data
- Invalid file paths
- File system errors
- Excel structure issues

**Common Issues and Solutions:**

1. **"No test case data found"**
   - Ensure your data starts from Row 2
   - Check that you have data in Column A

2. **"No output directory selected"**
   - When prompted, select any file in your desired output folder
   - The macro will extract the folder path automatically

3. **"An error occurred"**
   - Check that you have write permissions to the output directory
   - Ensure the output directory exists

## File Naming Convention

Generated feature files are named using:
1. **Test Case ID** (if available)
2. **Scenario Name** (if Test Case ID is empty)
3. **Timestamp** (if both are empty)

Invalid characters are automatically replaced with underscores.

## Example Output

**Input Excel Data:**
| Test Case ID | Scenario Name | Steps | Expected Result |
|--------------|---------------|-------|-----------------|
| TC001 | Login Test | Enter username "admin" and password "123" | User should be logged in successfully |

**Generated Feature File (`TC001.feature`):**
```gherkin
Feature: Login Test

  Scenario: Login Test - TC001
    Given the test setup is ready
    When I perform the following steps:
      Enter username "admin" and password "123"
    Then I expect:
      User should be logged in successfully
```

## Troubleshooting

### Macro Not Running
1. Ensure macros are enabled in Excel
2. Check that the VBA code was copied correctly
3. Verify the module name matches the code

### No Files Generated
1. Check that you selected a valid output directory
2. Ensure you have write permissions
3. Verify your Excel data structure

### Incorrect Data Extraction
1. Run `ValidateSheetStructure()` to check your data format
2. Verify column mappings in the constants
3. Check that data starts from the correct row

## Security Notes

- The macro only reads from your active worksheet
- It creates new files but doesn't modify existing ones
- No data is sent to external sources
- All processing is done locally on your machine

## Support

For issues or questions:
1. Run `ShowInstructions()` for basic help
2. Use `ValidateSheetStructure()` to check your data format
3. Check the error messages for specific guidance

## Version Information

- **Version**: 1.0
- **Compatibility**: Excel 2010 and later
- **VBA Version**: Compatible with all modern Excel VBA environments
