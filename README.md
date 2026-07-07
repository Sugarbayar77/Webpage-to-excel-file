# Webpage-to-excel-file
It detects HTML `&lt;table>` tags, extracts them into DataFrames, and saves them into an Excel file. If the file doesn’t exist, it creates one; if it does, new tables are added as separate sheets with unique timestamped names. This way, the Excel file grows into a structured archive of scraped tables over time.
