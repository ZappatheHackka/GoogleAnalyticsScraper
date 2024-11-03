A script I wrote that we use at my company to generate monthly web traffic reports.

Uses xlsxwriter to create Excel spreadsheets with multiple tabs, the data itself being pulled via the Google Search Console API.
The spreadsheet itself was designed with a two-page format in mind, based on the 'query' and 'page' arguments passed to GSC.

The first page focused on the 'query' data; tracking our top-performing monthly keywords (those with more than an insignificant # of clicks) as well as their respective clicks, impressions, and ctr. 
Then the monthly total clicks, impressions, and average ctr are calculated and displayed above last month's totals, before finally the month-to-month % change in all values is displayed.

The second page functions more or less the same albeit with different data. Our top-ranking web pages take the place of keywords. 
Last month's data is displayed, although given the monthly fluctuations of top-ranking pages, month-to-month comparisons are not made beyond totals.

I used company colors in the headers, and primarily made this as a tool to track changes in web traffic as I began to implement more heavy-duty SEO-related web updates. 
In retrospect, I'm not sure why I didn't put some of the repeated functionality(querying the API, combining 'http' & 'https' entries) into functions to make the code cleaner. 
Seems like a logical step. I guess as the scope of the project gradually expanded I just ended up copy-pasting the initial querying, editing it to my needs.

I used some of the functions on the following webpage to help me set up the GSC API: https://engineeringfordatascience.com/posts/google_search_console_api_python/ --Very helpful tutorial!

Plug in your own secret keys and generate your own straightforward SEO reports.
