## 📊 Selenium Web Scraper for Exam Results

This project uses **Selenium** to automate the scraping of student results from the official results website. The script extracts individual marks, calculates whether each student has passed or failed, and saves the data into an **Excel sheet**.

### ✅ Key Features

- 🌐 Automated web scraping with Selenium
- 📋 Extracts student marks and details from the results website
- 📊 Saves data into an Excel (.xlsx) sheet
- 🎨 Marks each subject as **Pass (green)** or **Fail (red)** using cell background color

### 📁 Output Example

The Excel output includes:

| Student Name | Subject 1 | Subject 2 | Subject 3 | Result |
|--------------|-----------|-----------|-----------|--------|
| John Doe     | 🟩 45     | 🟥 20     | 🟩 32     | Fail   |
| Jane Smith   | 🟩 58     | 🟩 49     | 🟩 51     | Pass   |

> 🟩 Green cell = Pass (mark >= passing score)  
> 🟥 Red cell = Fail (mark < passing score)

### 🛠 Technologies Used

- Python 🐍
- Selenium WebDriver
- openpyxl (for Excel editing)
