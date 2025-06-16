const XLSX = require('xlsx');
const fs = require('fs');

function reformatFacultySurveySelectedOnly(filename) {
    console.log(`Processing file: ${filename}`);
    
    // Read the file
    const workbook = XLSX.readFile(filename);
    const firstSheetName = Object.keys(workbook.Sheets)[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // Convert to array format to work with column indexes
    const rawData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1, 
        raw: false,
        defval: ''
    });
    
    console.log(`Loaded ${rawData.length} rows with ${rawData[0].length} columns`);
    
    // Define the exact items and their categories
    const categories = {
        sdgs: {
            name: 'SDGs',
            items: [
                { id: 'SDG_1', name: 'SDG 1: No Poverty', keyword: 'no poverty' },
                { id: 'SDG_2', name: 'SDG 2: Zero Hunger', keyword: 'zero hunger' },
                { id: 'SDG_3', name: 'SDG 3: Good Health and Well-being', keyword: 'good health' },
                { id: 'SDG_4', name: 'SDG 4: Quality Education', keyword: 'quality education' },
                { id: 'SDG_5', name: 'SDG 5: Gender Equality', keyword: 'gender equality' },
                { id: 'SDG_6', name: 'SDG 6: Clean Water and Sanitation', keyword: 'clean water' },
                { id: 'SDG_7', name: 'SDG 7: Affordable and Clean Energy', keyword: 'affordable energy' },
                { id: 'SDG_8', name: 'SDG 8: Decent Work and Economic Growth', keyword: 'decent work' },
                { id: 'SDG_9', name: 'SDG 9: Industry, Innovation and Infrastructure', keyword: 'industry innovation' },
                { id: 'SDG_10', name: 'SDG 10: Reduced Inequalities', keyword: 'reduced inequalities' },
                { id: 'SDG_11', name: 'SDG 11: Sustainable Cities and Communities', keyword: 'sustainable cities' },
                { id: 'SDG_12', name: 'SDG 12: Responsible Consumption and Production', keyword: 'responsible consumption' },
                { id: 'SDG_13', name: 'SDG 13: Climate Action', keyword: 'climate action' },
                { id: 'SDG_14', name: 'SDG 14: Life Below Water', keyword: 'life below water' },
                { id: 'SDG_15', name: 'SDG 15: Life on Land', keyword: 'life on land' },
                { id: 'SDG_16', name: 'SDG 16: Peace, Justice and Strong Institutions', keyword: 'peace justice' },
                { id: 'SDG_17', name: 'SDG 17: Partnerships for the Goals', keyword: 'partnership' },
                { id: 'SDG_18', name: 'SDG 18: Other', keyword: 'other' }
            ]
        },
        keywords: {
            name: 'Keywords',
            items: [
                { id: 'KW_1', name: 'Interconnection', keyword: 'interconnection' },
                { id: 'KW_2', name: 'Ethics', keyword: 'ethics' },
                { id: 'KW_3', name: 'Justice', keyword: 'justice' },
                { id: 'KW_4', name: 'Preservation for Future Generations', keyword: 'preservation' },
                { id: 'KW_5', name: 'Equity', keyword: 'equity' },
                { id: 'KW_6', name: 'Other', keyword: 'other' }
            ]
        },
        contentTopics: {
            name: 'Content Topics',
            items: [
                { id: 'CT_1', name: 'Environmental Policy', keyword: 'environmental policy' },
                { id: 'CT_2', name: 'Social Innovation', keyword: 'social innovation' },
                { id: 'CT_3', name: 'Economic Sustainability', keyword: 'economic sustainability' },
                { id: 'CT_4', name: 'Corporate Sustainability', keyword: 'corporate sustainability' },
                { id: 'CT_5', name: 'Climate Change', keyword: 'climate change' },
                { id: 'CT_6', name: 'Resource Management', keyword: 'resource management' },
                { id: 'CT_7', name: 'Sustainable Development', keyword: 'sustainable development' },
                { id: 'CT_8', name: 'Environmental Justice', keyword: 'environmental justice' },
                { id: 'CT_9', name: 'Other', keyword: 'other' }
            ]
        },
        competencies: {
            name: 'Competencies',
            items: [
                { id: 'COMP_1', name: 'Systems Thinking', keyword: 'systems thinking' },
                { id: 'COMP_2', name: 'Anticipatory Competency', keyword: 'anticipatory' },
                { id: 'COMP_3', name: 'Normative Competency', keyword: 'normative' },
                { id: 'COMP_4', name: 'Strategic Competency', keyword: 'strategic' },
                { id: 'COMP_5', name: 'Interpersonal Competency', keyword: 'interpersonal' },
                { id: 'COMP_6', name: 'Other', keyword: 'other' }
            ]
        }
    };
    
    // Define course configurations based on actual column positions
    const courseConfigs = [
        {
            name: 'Course 1',
            courseCol: 2,
            sustainabilityValues: 3,
            sustainabilityValuesTime: 4,
            sustainabilityKnowledge: 5,
            sustainabilityKnowledgeTime: 6,
            sustainabilitySkills: 7,
            sustainabilitySkillsTime: 8,
            globalGoalsStart: 9,
            globalGoalsEnd: 26,
            keywordsStart: 27,
            keywordsEnd: 32,
            contentTopicsStart: 33,
            contentTopicsEnd: 41,
            competenciesStart: 42,
            competenciesEnd: 47,
            selfEvaluation: 49,
            prefix: 'Q2'
        },
        {
            name: 'Course 2',
            courseCol: 51,
            sustainabilityValues: 52,
            sustainabilityValuesTime: 53,
            sustainabilityKnowledge: 54,
            sustainabilityKnowledgeTime: 55,
            sustainabilitySkills: 56,
            sustainabilitySkillsTime: 57,
            globalGoalsStart: 58,
            globalGoalsEnd: 75,
            keywordsStart: 76,
            keywordsEnd: 81,
            contentTopicsStart: 82,
            contentTopicsEnd: 90,
            competenciesStart: 91,
            competenciesEnd: 96,
            selfEvaluation: 98,
            prefix: 'Q3'
        },
        {
            name: 'Course 3',
            courseCol: 100,
            sustainabilityValues: 101,
            sustainabilityValuesTime: 102,
            sustainabilityKnowledge: 103,
            sustainabilityKnowledgeTime: 104,
            sustainabilitySkills: 105,
            sustainabilitySkillsTime: 106,
            globalGoalsStart: 107,
            globalGoalsEnd: 124,
            keywordsStart: 125,
            keywordsEnd: 130,
            contentTopicsStart: 131,
            contentTopicsEnd: 139,
            competenciesStart: 140,
            competenciesEnd: 145,
            selfEvaluation: 147,
            prefix: 'Q4'
        },
        {
            name: 'Course 4',
            courseCol: 149,
            sustainabilityValues: 150,
            sustainabilityValuesTime: 151,
            sustainabilityKnowledge: 152,
            sustainabilityKnowledgeTime: 153,
            sustainabilitySkills: 154,
            sustainabilitySkillsTime: 155,
            globalGoalsStart: 156,
            globalGoalsEnd: 173,
            keywordsStart: 174,
            keywordsEnd: 179,
            contentTopicsStart: 180,
            contentTopicsEnd: 188,
            competenciesStart: 189,
            competenciesEnd: 194,
            selfEvaluation: 196,
            prefix: 'Q5'
        },
        {
            name: 'Course 5',
            courseCol: 198,
            sustainabilityValues: 199,
            sustainabilityValuesTime: 200,
            sustainabilityKnowledge: 201,
            sustainabilityKnowledgeTime: 202,
            sustainabilitySkills: 203,
            sustainabilitySkillsTime: 204,
            globalGoalsStart: 205,
            globalGoalsEnd: 222,
            keywordsStart: 223,
            keywordsEnd: 228,
            contentTopicsStart: 229,
            contentTopicsEnd: 237,
            competenciesStart: 238,
            competenciesEnd: 243,
            selfEvaluation: 245,
            prefix: 'Q6'
        },
        {
            name: 'Course 6',
            courseCol: 247,
            sustainabilityValues: 248,
            sustainabilityValuesTime: 249,
            sustainabilityKnowledge: 250,
            sustainabilityKnowledgeTime: 251,
            sustainabilitySkills: 252,
            sustainabilitySkillsTime: 253,
            globalGoalsStart: 254,
            globalGoalsEnd: 271,
            keywordsStart: 272,
            keywordsEnd: 277,
            contentTopicsStart: 278,
            contentTopicsEnd: 286,
            competenciesStart: 287,
            competenciesEnd: 292,
            selfEvaluation: 294,
            prefix: 'Q102'
        }
    ];
    
    const emailCol = 299;
    const facultyNameCol = 300;
    
    // Helper functions
    function cleanValue(value) {
        if (!value || value.toString().trim() === '' || value.toString().trim() === ' ') {
            return 'N/A';
        }
        return value.toString().trim();
    }
    
    function hasValue(value) {
        return value && value.toString().trim() !== '' && value.toString().trim() !== ' ';
    }
    
    // Create separate datasets for each category - ONLY SELECTED ITEMS
    const datasets = {
        sdgs: [],
        keywords: [],
        contentTopics: [],
        competencies: []
    };
    
    // Process each row (skip header row)
    for (let rowIndex = 1; rowIndex < rawData.length; rowIndex++) {
        const row = rawData[rowIndex];
        let email = cleanValue(row[emailCol]);
        let facultyName = cleanValue(row[facultyNameCol]);
        
        // Make anonymous if no email or faculty name
        if (email === 'N/A' && facultyName === 'N/A') {
            email = 'Anonymous';
            facultyName = 'Anonymous';
        } else if (email === 'N/A') {
            email = 'Anonymous';
        } else if (facultyName === 'N/A') {
            facultyName = 'Anonymous';
        }
        
        // Process each course
        courseConfigs.forEach(config => {
            const courseTitle = row[config.courseCol];
            
            // Skip if no course title
            if (!hasValue(courseTitle)) {
                return;
            }
            
            // Create base record for this course
            const baseRecord = {
                ResponseId: row[0] || `Row_${rowIndex}`,
                FacultyName: facultyName,
                FacultyEmail: email,
                CourseSlot: config.name,
                CourseTitle: courseTitle.toString().trim(),
                IncludesSustainabilityValues: cleanValue(row[config.sustainabilityValues]),
                SustainabilityValuesTime: cleanValue(row[config.sustainabilityValuesTime]),
                IncludesSustainabilityKnowledge: cleanValue(row[config.sustainabilityKnowledge]),
                SustainabilityKnowledgeTime: cleanValue(row[config.sustainabilityKnowledgeTime]),
                IncludesSustainabilitySkills: cleanValue(row[config.sustainabilitySkills]),
                SustainabilitySkillsTime: cleanValue(row[config.sustainabilitySkillsTime]),
                SelfEvaluation: cleanValue(row[config.selfEvaluation])
            };
            
            // Process SDGs - Only create rows for SELECTED items
            const selectedSDGs = [];
            for (let i = config.globalGoalsStart; i <= config.globalGoalsEnd; i++) {
                const goalValue = row[i];
                if (hasValue(goalValue)) {
                    selectedSDGs.push(goalValue.toString().trim());
                }
            }
            
            // Create rows only for selected SDGs
            categories.sdgs.items.forEach(item => {
                const isSelected = selectedSDGs.some(selected => 
                    selected.toLowerCase().includes(item.keyword) || 
                    item.keyword.includes(selected.toLowerCase())
                );
                
                if (isSelected) {  // ONLY add row if selected
                    const sdgRecord = {
                        ...baseRecord,
                        ItemId: item.id,
                        ItemName: item.name,
                        ItemCategory: 'SDG',
                        OriginalSelection: selectedSDGs.find(selected => 
                            selected.toLowerCase().includes(item.keyword) || 
                            item.keyword.includes(selected.toLowerCase())
                        ),
                        TotalSelectedInCategory: selectedSDGs.length
                    };
                    
                    datasets.sdgs.push(sdgRecord);
                }
            });
            
            // Process Keywords - Only create rows for SELECTED items
            const selectedKeywords = [];
            for (let i = config.keywordsStart; i <= config.keywordsEnd; i++) {
                const keywordValue = row[i];
                if (hasValue(keywordValue)) {
                    selectedKeywords.push(keywordValue.toString().trim());
                }
            }
            
            // Create rows only for selected keywords
            categories.keywords.items.forEach(item => {
                const isSelected = selectedKeywords.some(selected => 
                    selected.toLowerCase().includes(item.keyword) || 
                    item.keyword.includes(selected.toLowerCase())
                );
                
                if (isSelected) {  // ONLY add row if selected
                    const keywordRecord = {
                        ...baseRecord,
                        ItemId: item.id,
                        ItemName: item.name,
                        ItemCategory: 'Keyword',
                        OriginalSelection: selectedKeywords.find(selected => 
                            selected.toLowerCase().includes(item.keyword) || 
                            item.keyword.includes(selected.toLowerCase())
                        ),
                        TotalSelectedInCategory: selectedKeywords.length
                    };
                    
                    datasets.keywords.push(keywordRecord);
                }
            });
            
            // Process Content Topics - Only create rows for SELECTED items
            const selectedContentTopics = [];
            for (let i = config.contentTopicsStart; i <= config.contentTopicsEnd; i++) {
                const topicValue = row[i];
                if (hasValue(topicValue)) {
                    selectedContentTopics.push(topicValue.toString().trim());
                }
            }
            
            // Create rows only for selected content topics
            categories.contentTopics.items.forEach(item => {
                const isSelected = selectedContentTopics.some(selected => 
                    selected.toLowerCase().includes(item.keyword) || 
                    item.keyword.includes(selected.toLowerCase())
                );
                
                if (isSelected) {  // ONLY add row if selected
                    const contentTopicRecord = {
                        ...baseRecord,
                        ItemId: item.id,
                        ItemName: item.name,
                        ItemCategory: 'Content Topic',
                        OriginalSelection: selectedContentTopics.find(selected => 
                            selected.toLowerCase().includes(item.keyword) || 
                            item.keyword.includes(selected.toLowerCase())
                        ),
                        TotalSelectedInCategory: selectedContentTopics.length
                    };
                    
                    datasets.contentTopics.push(contentTopicRecord);
                }
            });
            
            // Process Competencies - Only create rows for SELECTED items
            const selectedCompetencies = [];
            for (let i = config.competenciesStart; i <= config.competenciesEnd; i++) {
                const competencyValue = row[i];
                if (hasValue(competencyValue)) {
                    selectedCompetencies.push(competencyValue.toString().trim());
                }
            }
            
            // Create rows only for selected competencies
            categories.competencies.items.forEach(item => {
                const isSelected = selectedCompetencies.some(selected => 
                    selected.toLowerCase().includes(item.keyword) || 
                    item.keyword.includes(selected.toLowerCase())
                );
                
                if (isSelected) {  // ONLY add row if selected
                    const competencyRecord = {
                        ...baseRecord,
                        ItemId: item.id,
                        ItemName: item.name,
                        ItemCategory: 'Competency',
                        OriginalSelection: selectedCompetencies.find(selected => 
                            selected.toLowerCase().includes(item.keyword) || 
                            item.keyword.includes(selected.toLowerCase())
                        ),
                        TotalSelectedInCategory: selectedCompetencies.length
                    };
                    
                    datasets.competencies.push(competencyRecord);
                }
            });
        });
    }
    
    console.log(`Generated selected-only datasets:`);
    console.log(`  SDGs: ${datasets.sdgs.length} selected items`);
    console.log(`  Keywords: ${datasets.keywords.length} selected items`);
    console.log(`  Content Topics: ${datasets.contentTopics.length} selected items`);
    console.log(`  Competencies: ${datasets.competencies.length} selected items`);
    
    return datasets;
}

// Convert dataset to CSV
function convertToCSV(data) {
    if (!data || data.length === 0) return '';
    
    const headers = Object.keys(data[0]);
    const csvRows = [headers.join(',')];
    
    data.forEach(row => {
        const values = headers.map(header => {
            const value = row[header] || '';
            const stringValue = value.toString();
            // Escape commas and quotes
            if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
                return `"${stringValue.replace(/"/g, '""')}"`;
            }
            return stringValue;
        });
        csvRows.push(values.join(','));
    });
    
    return csvRows.join('\n');
}

// Create Excel workbook with multiple sheets
function createExcelWorkbook(datasets) {
    const workbook = XLSX.utils.book_new();
    
    Object.keys(datasets).forEach(key => {
        const data = datasets[key];
        if (data && data.length > 0) {
            const worksheet = XLSX.utils.json_to_sheet(data);
            const sheetName = key.charAt(0).toUpperCase() + key.slice(1);
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        }
    });
    
    return workbook;
}

// Analyze datasets
function analyzeDatasets(datasets) {
    const analysis = {};
    
    Object.keys(datasets).forEach(key => {
        const data = datasets[key];
        const categoryName = key.charAt(0).toUpperCase() + key.slice(1);
        
        // Get unique courses in this dataset
        const uniqueCourses = [...new Set(data.map(d => `${d.FacultyEmail}_${d.CourseTitle}`))];
        const totalCourses = uniqueCourses.length;
        
        // Item popularity (count how many times each item was selected)
        const itemCounts = {};
        data.forEach(record => {
            if (!itemCounts[record.ItemName]) {
                itemCounts[record.ItemName] = 0;
            }
            itemCounts[record.ItemName]++;
        });
        
        const mostPopularItems = Object.entries(itemCounts)
            .map(([name, count]) => ({ name, count }))
            .sort((a, b) => b.count - a.count)
            .slice(0, 10);
        
        // Faculty engagement
        const facultyEngagement = {};
        data.forEach(record => {
            if (!facultyEngagement[record.FacultyEmail]) {
                facultyEngagement[record.FacultyEmail] = {
                    name: record.FacultyName,
                    courses: new Set(),
                    totalSelections: 0
                };
            }
            facultyEngagement[record.FacultyEmail].courses.add(record.CourseTitle);
            facultyEngagement[record.FacultyEmail].totalSelections++;
        });
        
        const topFaculty = Object.entries(facultyEngagement)
            .map(([email, data]) => ({
                email,
                name: data.name,
                courses: data.courses.size,
                totalSelections: data.totalSelections
            }))
            .filter(f => f.email !== 'Anonymous')
            .sort((a, b) => b.totalSelections - a.totalSelections)
            .slice(0, 5);
        
        analysis[key] = {
            name: categoryName,
            totalSelections: data.length,
            totalCourses: totalCourses,
            uniqueFaculty: new Set(data.filter(d => d.FacultyEmail !== 'Anonymous').map(d => d.FacultyEmail)).size,
            averageSelectionsPerCourse: totalCourses > 0 ? (data.length / totalCourses).toFixed(1) : 0,
            mostPopularItems: mostPopularItems,
            topFaculty: topFaculty,
            sustainabilityIntegration: {
                values: { yes: 0, no: 0, na: 0 },
                knowledge: { yes: 0, no: 0, na: 0 },
                skills: { yes: 0, no: 0, na: 0 }
            }
        };
        
        // Count sustainability integration (per unique course)
        const uniqueCourseRecords = data.filter((record, index, self) => 
            index === self.findIndex(r => r.ResponseId === record.ResponseId && r.CourseTitle === record.CourseTitle)
        );
        
        uniqueCourseRecords.forEach(record => {
            const values = record.IncludesSustainabilityValues.toLowerCase();
            if (values.includes('yes')) analysis[key].sustainabilityIntegration.values.yes++;
            else if (values.includes('no')) analysis[key].sustainabilityIntegration.values.no++;
            else analysis[key].sustainabilityIntegration.values.na++;
            
            const knowledge = record.IncludesSustainabilityKnowledge.toLowerCase();
            if (knowledge.includes('yes')) analysis[key].sustainabilityIntegration.knowledge.yes++;
            else if (knowledge.includes('no')) analysis[key].sustainabilityIntegration.knowledge.no++;
            else analysis[key].sustainabilityIntegration.knowledge.na++;
            
            const skills = record.IncludesSustainabilitySkills.toLowerCase();
            if (skills.includes('yes')) analysis[key].sustainabilityIntegration.skills.yes++;
            else if (skills.includes('no')) analysis[key].sustainabilityIntegration.skills.no++;
            else analysis[key].sustainabilityIntegration.skills.na++;
        });
    });
    
    return analysis;
}

// Main processing function
function processFileSelectedOnly(inputFile, outputFormat = 'excel') {
    try {
        console.log('=== Selected-Only Faculty Survey Reformatter ===');
        console.log('Each row represents one Course with one SELECTED item');
        
        // Check if input file exists
        if (!fs.existsSync(inputFile)) {
            throw new Error(`File not found: ${inputFile}`);
        }
        
        // Process the data
        const datasets = reformatFacultySurveySelectedOnly(inputFile);
        
        // Generate analysis
        console.log('\n=== Analysis by Category ===');
        const analysis = analyzeDatasets(datasets);
        
        Object.keys(analysis).forEach(key => {
            const categoryAnalysis = analysis[key];
            console.log(`\n${categoryAnalysis.name}:`);
            console.log(`  Total Selections: ${categoryAnalysis.totalSelections}`);
            console.log(`  Courses with Selections: ${categoryAnalysis.totalCourses}`);
            console.log(`  Faculty Members: ${categoryAnalysis.uniqueFaculty}`);
            console.log(`  Avg Selections per Course: ${categoryAnalysis.averageSelectionsPerCourse}`);
            console.log(`  Most Popular Items:`);
            categoryAnalysis.mostPopularItems.forEach((item, index) => {
                console.log(`    ${index + 1}. ${item.name}: ${item.count} courses`);
            });
            if (categoryAnalysis.topFaculty.length > 0) {
                console.log(`  Most Active Faculty:`);
                categoryAnalysis.topFaculty.forEach((faculty, index) => {
                    console.log(`    ${index + 1}. ${faculty.name}: ${faculty.totalSelections} selections across ${faculty.courses} courses`);
                });
            }
        });
        
        const baseFilename = inputFile.replace(/\.[^/.]+$/, '');
        let outputPaths = [];
        
        if (outputFormat === 'excel') {
            // Create Excel file with multiple sheets
            const workbook = createExcelWorkbook(datasets);
            const excelPath = `${baseFilename}_selected_only.xlsx`;
            XLSX.writeFile(workbook, excelPath);
            outputPaths.push(excelPath);
            console.log(`\n✅ Success! Selected-only Excel file saved to: ${excelPath}`);
        } else {
            // Create separate CSV files
            Object.keys(datasets).forEach(key => {
                const csvOutput = convertToCSV(datasets[key]);
                const csvPath = `${baseFilename}_${key}_selected.csv`;
                fs.writeFileSync(csvPath, csvOutput, 'utf8');
                outputPaths.push(csvPath);
                console.log(`✅ ${key.charAt(0).toUpperCase() + key.slice(1)} CSV saved to: ${csvPath}`);
            });
        }
        
        // Show sample structure
        console.log('\n=== Sample Selected-Only Structure ===');
        const sampleRecord = datasets.sdgs[0];
        if (sampleRecord) {
            console.log('Each row represents one Course with one SELECTED item:');
            ['ResponseId', 'FacultyName', 'CourseTitle', 'ItemId', 'ItemName', 'OriginalSelection'].forEach(col => {
                console.log(`  ${col}: ${sampleRecord[col]}`);
            });
            console.log('\nExample: If AS101 selected SDG1 and SDG4, you get:');
            console.log('  Row 1: AS101 | SDG_1 | SDG 1: No Poverty');
            console.log('  Row 2: AS101 | SDG_4 | SDG 4: Quality Education');
            console.log('  (No rows for SDG2, SDG3, etc.)');
        }
        
        return { datasets, analysis, outputPaths };
        
    } catch (error) {
        console.error('❌ Error:', error.message);
        process.exit(1);
    }
}

// Command line usage
if (require.main === module) {
    const args = process.argv.slice(2);
    
    if (args.length === 0) {
        console.log('Usage: node selected-only-reformatter.js <input-file> [format]');
        console.log('');
        console.log('Arguments:');
        console.log('  input-file    Input Excel/CSV file to process');
        console.log('  format        Output format: "excel" (default) or "csv"');
        console.log('');
        console.log('Examples:');
        console.log('  node selected-only-reformatter.js "survey.xlsx"           # Creates Excel with 4 sheets');
        console.log('  node selected-only-reformatter.js "survey.csv" csv        # Creates 4 separate CSV files');
        console.log('');
        console.log('Output Structure:');
        console.log('  ONLY rows for SELECTED items - no Y/N columns needed');
        console.log('  If course selected 3 SDGs → 3 rows in SDG sheet');
        console.log('  If course selected 2 keywords → 2 rows in Keywords sheet');
        console.log('  If course selected 0 topics → 0 rows in Topics sheet');
        console.log('');
        console.log('Key Columns:');
        console.log('  - ItemId: Unique identifier (e.g., SDG_1, KW_3, CT_5)');
        console.log('  - ItemName: Full name (e.g., "SDG 1: No Poverty")');
        console.log('  - OriginalSelection: The exact text that was selected in survey');
        console.log('  - TotalSelectedInCategory: How many total items selected in this category');
        console.log('');
        console.log('Perfect for:');
        console.log('  - Clean, compact data with no empty rows');
        console.log('  - Counting item popularity: COUNT(ItemName)');
        console.log('  - Faculty engagement analysis');
        console.log('  - Course-item relationship mapping');
        process.exit(1);
    }
    
    const format = args[1] || 'excel';
    if (format !== 'excel' && format !== 'csv') {
        console.error('Error: Format must be "excel" or "csv"');
        process.exit(1);
    }
    
    processFileSelectedOnly(args[0], format);
}

module.exports = { 
    reformatFacultySurveySelectedOnly, 
    convertToCSV, 
    createExcelWorkbook,
    analyzeDatasets, 
    processFileSelectedOnly 
};