**Process Guide: Data Extraction and Analysis Project**

**Step 1: Download and Set Up**

1. **Download Files**
   - Clone or download the repository files from GitHub to your local machine.
   - Save all files to a folder on your **Desktop**.

2. **Set Up Your Environment**
   - Use a good IDE such as **Spyder** or **Jupyter Notebook** to run the Python files.

**Step 2: Run Data Extraction**

1. **Open `data_extraction_sta.py`**
   - In your IDE, open the `data_extraction_sta.py` file.
   
2. **Execute the File**
   - Run the program to begin the extraction process.
   - The program will download 100 articles and save them in a folder named **"blackassign_articles"** on your Desktop.
   
3. **Error Handling**
   - The program will display error messages for links that fail, categorized by the type of error. For example:
     - **HTTP Error**: Occurs when a webpage is not found or an incorrect URL is provided.
     - **Connection Error**: Occurs when there are network issues preventing access to the article.
   
**Step 3: Run Data Analysis**

1. **Open `data_analysis_sta.py`**
   - In a new tab of your IDE, open the `data_analysis_sta.py` file.

2. **Execute the File**
   - Run the program to perform sentiment and textual analysis on the articles extracted in the previous step.
   - The analysis will utilize parameters from the **"master dictionary"** and **"stopwords"** folder to analyze the articles.
   
3. **Output**
   - Once the program completes, it will generate an **Excel file** containing the sentiment and textual analysis results.
   - The program will also print error messages for any files that are missing or cannot be processed.

**Step 4: Review Output**
   - Check the Excel file for a detailed breakdown of the sentiment and textual analysis of the extracted articles.
