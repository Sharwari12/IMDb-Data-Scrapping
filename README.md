# Web Scraping Project: IMDb Top Rated Movies

In this project, I scraped data from the IMDb website to compile a list of the top 25 rated movies. 

I utilized the following Python libraries:

•	BeautifulSoup: For parsing HTML and extracting data from the web page.

•	Requests: To send HTTP requests and retrieve web page content.

•	OpenPyXL: To create and manage Excel files, enabling easy storage and presentation of the scraped data.
Project Workflow:

1.	Excel Sheet Creation:
	Created a new Excel workbook using openpyxl.Workbook().
	Retrieved the active sheet with excel.active and renamed it to "Top 25 IMDb Rated Movies".
	Added column headers to the sheet: ['Movie Rank', 'Movie Name', 'Release Year', 'IMDb Rating'].

2.	Sending HTTP Request:
	Defined HTTP headers using a User-Agent to simulate a browser request, improving the chances of successfully accessing the page.
	Sent a GET request to the IMDb top movies page with requests.get() and checked for a successful response using webpage.raise_for_status().

3.	Parsing Web Content:
	Parsed the webpage content using BeautifulSoup: soup = BeautifulSoup(webpage.text, "html.parser").
	Located the list of movies by finding the appropriate HTML elements that contained the relevant data.

4.	Data Extraction:
	Looped through each movie item in the list to extract the movie name, rank, release year, and IMDb rating.
	Printed the extracted details to the console for verification.

5.	Storing Data in Excel:
	Appended each movie's data to the Excel sheet using sheet.append().
	Saved the Excel file as "IMDb Top Rated Movies.xlsx" at the end of the scraping process.

This project demonstrates my ability to extract data from websites, manage it effectively, and present it in a user-friendly format.

 
