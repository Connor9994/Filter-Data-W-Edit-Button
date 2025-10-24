# Filter-Data-W-Edit-Button

![GitHub stars](https://img.shields.io/github/stars/Connor9994/Filter-Data-W-Edit-Button?style=social) ![GitHub forks](https://img.shields.io/github/forks/Connor9994/Filter-Data-W-Edit-Button?style=social) ![GitHub issues](https://img.shields.io/github/issues/Connor9994/Filter-Data-W-Edit-Button) 

SEE "Instructions" Folder for PDF examples (Blurred due to the PHI-nature of the info)

# TANYR - Excel Data Processing Tool

## Overview

TANYR is a PowerShell-based Windows Forms application designed to automate data processing tasks in Microsoft Excel. This tool provides a graphical interface for creating and applying data transformation rules, then organizing processed data into separate tabs based on defined criteria.

## Features

### 🗂️ Sheet Selection
- Connect to active Excel workbooks
- Select specific data sheets for processing
- Persistent storage of previously selected sheets

![1](https://github.com/Connor9994/Filter-Data-W-Edit-Button/blob/main/Pictures/1.png)

### ⚙️ Rule Management
- **Create Custom Rules**: Define data transformation rules with multiple criteria
- **Match Rules**: Filter data based on specific values in designated columns
- **Alphabetical Rules**: Apply rules based on alphabetical ranges (A-Z)
- **Rule Prioritization**: Reorder rules with drag-and-drop functionality
- **Rule Editing**: Modify existing rules with an intuitive interface

![2](https://github.com/Connor9994/Filter-Data-W-Edit-Button/blob/main/Pictures/2.png)

### 🔄 Apply Rules
- Process Excel data according to defined rules
- Automatically assign names based on rule conditions
- Create backup of original files before processing
- Real-time progress tracking

### 📑 Tab Sorting & Organization
- Split processed data into separate Excel tabs
- Select specific columns to copy to new tabs
- Customize tab creation order and content
- Export organized data to new Excel files

![3](https://github.com/Connor9994/Filter-Data-W-Edit-Button/blob/main/Pictures/3.png)

## How It Works

### Rule Types

1. **Value Matching Rules**
   - Match specific values in designated columns
   - Supports multiple matching criteria
   - Case-sensitive value comparison

2. **Alphabetical Range Rules**
   - Apply rules based on first letter ranges (e.g., A-M, N-Z)
   - Customizable alphabetical boundaries
   - Flexible column selection

### Data Flow

1. **Input**: Active Excel workbook with data sheet
2. **Processing**: Apply user-defined rules to transform data
3. **Output**: Reorganized data in new tabs based on rule assignments
4. **Export**: New Excel file with structured, categorized data

## Prerequisites

- **Microsoft Excel** must be installed and accessible
- **PowerShell** with Windows Forms support
- Appropriate permissions to read/write Excel files

## Usage

1. **Launch the Application**
   ```powershell
   .\TANYR.ps1
   ```

2. **Select Data Source**
   - Click "Select Sheet" to choose an Excel worksheet
   - Ensure Excel is open with the target workbook

3. **Configure Rules**
   - Use "Manage Rules" to create transformation rules
   - Define matching criteria and alphabetical ranges
   - Set rule execution order

4. **Process Data**
   - Click "Apply Rules" to transform the data
   - Select the target column for name assignments

5. **Organize Output**
   - Use "Sort Tabs" to split data into separate sheets
   - Choose columns to include in each tab
   - Export final organized workbook

## Rule Configuration

### Creating a Rule
1. **Name**: Unique identifier for the rule
2. **Assigned Name**: Value to assign when rule conditions are met
3. **Match Conditions**: Optional value matching in specific columns
4. **Alphabetical Range**: Optional letter range filtering
5. **Execution Order**: Rules are processed from top to bottom

### Rule Priority
- Rules are evaluated in defined order
- First matching rule determines the assigned value
- Use up/down arrows to adjust priority

## Technical Details

- **Backup System**: Automatically creates backups before processing
- **Error Handling**: Comprehensive try-catch blocks for stability
- **Memory Management**: Proper COM object cleanup and garbage collection
- **Persistent Settings**: User preferences saved between sessions
- **Performance**: Optimized for large datasets with efficient processing

## Limitations

- Requires Microsoft Excel to be installed
- Designed for Windows environments
- Processes one worksheet at a time
- Rules are applied sequentially (first match wins)

## Support

For issues or questions regarding TANYR, ensure:
- Excel is properly installed and accessible
- Input files are in supported Excel formats
- Sufficient permissions for file operations
- Adequate system resources for large datasets

---

*Note: This tool was developed for automated Excel data processing and organization, particularly useful for repetitive data categorization tasks.*
