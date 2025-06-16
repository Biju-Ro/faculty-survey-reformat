# Faculty Sustainability Survey Processor

Transforms Faculty Sustainability Survey data from wide format to long format, creating clean datasets for analysis in Tableau, Excel, and other visualization tools.

## Setup

### Prerequisites
- [Node.js](https://nodejs.org/) (version 14+)

### Installation
```bash
git clone [your-repo-url]
cd faculty-survey-processor
npm install
```

## Usage

### Basic Usage
```bash
# Process survey file (creates Excel with multiple sheets)
npm start "your-survey-file.xlsx"

# Create separate CSV files instead
npm start "your-survey-file.xlsx" csv
```

### Examples
```bash
npm start "2024 Sustainability Survey.xlsx"
npm start "survey-data.csv" csv
```

## Output

Creates clean long-format data with only selected survey items:
- **Excel**: Single file with 4 sheets (SDGs, Keywords, Content Topics, Competencies)
- **CSV**: 4 separate files for each category

Each row represents one course with one selected item, eliminating empty columns and making data analysis-ready.

## License

MIT