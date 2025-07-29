# Sample Excel Data Structure

This document shows the exact structure your Excel sheet should have for the VBA macro to work correctly.

## Required Structure

| A | B | C | D |
|---|---|---|---|
| **Test Case ID** | **Scenario Name** | **Steps** | **Expected Result** |
| TC001 | Login Test | Enter username "admin" and password "123" | User should be logged in successfully |
| TC002 | Search Functionality | Enter "test product" in search box and click search | Search results should display relevant products |
| TC003 | Add to Cart | Click "Add to Cart" button for item "Product A" | Item should be added to shopping cart |
| TC004 | Checkout Process | Complete checkout form and click "Place Order" | Order should be confirmed with order number |
| TC005 | User Registration | Fill registration form with valid data and submit | New user account should be created successfully |

## Important Notes

### Row Structure
- **Row 1**: Headers (Test Case ID, Scenario Name, Steps, Expected Result)
- **Row 2 onwards**: Actual test case data
- Empty rows will be automatically skipped

### Column Requirements
- **Column A**: Test Case ID (unique identifier)
- **Column B**: Scenario Name (descriptive name for the test)
- **Column C**: Steps (detailed test steps)
- **Column D**: Expected Result (what should happen after the test)

### Data Guidelines
1. **Test Case ID**: Should be unique and descriptive (e.g., TC001, LOGIN_001)
2. **Scenario Name**: Clear, descriptive name of what is being tested
3. **Steps**: Detailed step-by-step instructions
4. **Expected Result**: Clear description of expected outcome

## Example Output Files

### Generated Feature File: `TC001.feature`
```gherkin
Feature: Login Test

  Scenario: Login Test - TC001
    Given the test setup is ready
    When I perform the following steps:
      Enter username "admin" and password "123"
    Then I expect:
      User should be logged in successfully
```

### Generated Feature File: `TC002.feature`
```gherkin
Feature: Search Functionality

  Scenario: Search Functionality - TC002
    Given the test setup is ready
    When I perform the following steps:
      Enter "test product" in search box and click search
    Then I expect:
      Search results should display relevant products
```

## Common Mistakes to Avoid

1. **Missing Headers**: Ensure Row 1 contains the exact headers
2. **Wrong Column Order**: Data must be in columns A, B, C, D in that order
3. **Empty First Row**: Don't leave Row 1 empty
4. **Mixed Data Types**: Keep all data as text (don't use formulas in data cells)
5. **Special Characters**: Avoid using characters like `\/:*?"<>|` in Test Case IDs

## Customization

If your data is in different columns, you can modify the constants in the VBA code:

```vba
Private Const COL_TEST_CASE_ID As Integer = 1      ' Column A
Private Const COL_SCENARIO_NAME As Integer = 2     ' Column B
Private Const COL_STEPS As Integer = 3             ' Column C
Private Const COL_EXPECTED_RESULT As Integer = 4   ' Column D
```

For example, if your Test Case ID is in Column C:
```vba
Private Const COL_TEST_CASE_ID As Integer = 3      ' Column C
```