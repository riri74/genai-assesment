# Excel Template Auto-Fill with AI Matching

This script automates filling placeholders in an Excel template using data from multiple CSV files. It uses AI to match placeholder labels in the template with the best corresponding data fields.

---

## How It Works

1. **Load and Aggregate Data**  
   The script reads several CSV files containing cost and operational data. It sums up relevant numbers grouped by fields like roles or categories.

2. **Identify Placeholders**  
   It scans the Excel template for cells starting with `◦`, which are placeholders to be filled.

3. **Match Placeholders to Data Fields**  
   For each placeholder, it asks an AI (via Groq API) to find the closest matching data field from the aggregated data.

4. **Validate Matches**  
   The script checks if the AI’s match makes sense (e.g., nursing data shouldn’t match bed days). If not, it tries a fallback string similarity match.

5. **Fill the Template**  
   When a good match is found, the corresponding value is written next to the placeholder cell in the Excel sheet.

6. **Save Results**  
   The updated Excel file is saved as a new output file.

---

## How to Execute

1. Make sure you have Python 3 installed.

2. Set your Groq API key as an environment variable:

   ```bash
   export GROQ_API_KEY="your_api_key_here"
   python3 main.py

---

## Summary

This tool speeds up filling Excel templates by intelligently mapping placeholder labels to data fields using AI and validation checks, reducing manual effort and errors.
