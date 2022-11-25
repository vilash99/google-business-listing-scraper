<h1>Google Business Listing Scraper using VBA+Selenium</h1>
<p>When you search company, shop, or any other business at Google, it shows the business name, location, website, contact, etc on the right side of the browser. This is a simple VBA script embedded in Excel Macro that can scrap these business details.</p>
<p>You can read the full article on the Kushal Study blog.</p>

<h3>How to Use this script</h3>
<ol>
<li>Install Selenium on your System. You can check step-by-step instructions on How to <a href="https://www.kushalstudy.com/blog/how-to-install-selenium-webdriver-for-vba-in-windows/">How to Install Selenium Webdriver for VBA in Windows</a> blog.</li>
<li>Open <em>scrap-business-listing-google.xlsm</em> file and put business names in column A, starting from row 3.</li>
<li>Add the row number at cell H4 where you want to START scraping.</li>
<li>Add the row number at cell H6 where you want to END scraping.</li>
<li>Add seconds to pause after each keyword search. If you try to scrap very fast, then Google will show a captcha.</li>
<li>Click on button <b>Start Scraping</b>.</li>
</ol>

<h3>Example of how business listing is displayed on Google:</h3>
<img src="https://github.com/vilash99/google-business-listing-scraper/blob/main/google-business-listing-example.jpg" alt="google-search-result">

<h3>Excel Interface:</h3>
<img src="https://github.com/vilash99/google-business-listing-scraper/blob/main/excel-file-screenshot.jpg" alt="excel-dashboard">
