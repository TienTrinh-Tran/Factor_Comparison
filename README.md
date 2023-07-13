# Factor_Comparison

For competitive analysis, our team performs a quarterly study of average premiums for our company compared to top carriers. We review 10 states for both Auto and Home insurance each quarter. We use third-party software to calculate premiums for thousands of hypothetical risks in our model. Since insurance companies often change their rates every six or 12 months, we expect to observe changes in average premiums when comparing our current study to past studies. We have set an acceptable threshold of 5% for fluctuations. However, we often encounter significant changes for some carriers, which necessitates investigating the reasons behind them. Our initial focus is determining whether our vendor program correctly aligns with the carriers' insurance rate filings. This requires comparing the factors applied in the past with the current ones to verify the accuracy of the changes.

Before having this automated tool, our previous method involved using software to generate individual Excel files that displayed the factors applied. Each Excel file represented one risk with one carrier, and we had to repeat this process to generate files for past factor applications as well. We then used Excel formulas to compare the factors between the two files. This approach was highly repetitive and time-consuming, and I constantly sought a way to automate it. Finally, my wish came true when I began learning Python. 

The program performs the following tasks:
- Retrieves CSV files from the previous quarter and the current quarter. These files contain rating factors for all the hypothetical risks used in your pricing competitive model. The CSV files are generated from a third-party software.
- Based on user input, extracts the factors for the requested risks from the CSV files. It then arranges the past and current factors side by side and organizes them by carriers.
- Performs a comparison of the factors to identify which ones have changed and which ones remain unchanged. For the factors that have changed, the program calculates the percentage increase or decrease.
- Outputs the comparison results into an Excel workbook, with each requested risk having its own sheet.

In summary, the program automates the retrieval, extraction, comparison, and output of rating factors for requested risks, allowing for easy analysis and identification of changes in factors between previous and current quarters.

