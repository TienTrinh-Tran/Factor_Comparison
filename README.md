# Factor_Comparison

For competitive analysis, our team performs quarterly study of average premiums of our company against top carriers. We do the review for 10 states for both Auto and Home each quarter. We use a third-party software to calculate the premiums for thousands of hypothetical risks in our model. Since insurance companies often change their rate every six or 12 months, we expect to see the change in average premiums when comparing our current study to the past study. We set a 5% acceptable threshold for the fluctuation. However, often times we see some big changes for some carriers and that require us to find out the reason. The first thing we are looking for is: does our vendor program rates correctly according to carrier's insurance rate filings. This requires us to compare the factors applied in the past and now and see if the changes are correct.

Before having this automated tool, we often used the software to generate an Excel file showing the factor applied. One Excel file is for one risk and one carrier only and we needed to do so to generate the Excel file for past factor applied as well. We then used excel formula to compare the factors from the 2 files. This was a very repetitive & time consuming process and I was always seeking a way to automate it. I finally realized my wish when I started learning Python. 

What does the program do?
- retrieve the csv files, from previous quarter and current quarter, containg rating factors for all the hypothetical risks used in our pricing competitive model. These files are generated from a third-party software.
- based on user's input, extract only the factors for the requested risks, then arrange the past & current factors side by side and by carriers.
- perform the comparison to indicate the factors that changed or unchanged. For factors that are changed, calculate the % increase/decrease
- output the comparison into an Excel workbook, one sheet for each requested risk

Notes: 
There is a script for Auto and another for Home. The majority of the codes are the same. I decided to have 2 separate scripts for ease of maintenance and enhancement. The requirements for these 2 lines can be a lot different in many aspects. In addition, this makes the script more readable and understandble for team members who don't have much experience with Python programming language.
