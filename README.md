Unit 1 | Assignment - KickStart My Chart

Background
Over two billion dollars have been raised using the massively successful crowdfunding service, Kickstarter, but not every project has found success. Of the over 300,000 projects launched on Kickstarter, only a third have made it through the funding process with a positive outcome.
Since getting funded on Kickstarter requires meeting or exceeding the project's initial goal, many organizations spend months looking through past projects in an attempt to discover some trick to finding success. For this week's homework, you will organize and analyze a database of four thousand past projects in order to uncover any hidden trends.

Instructions



Using the Excel table provided, you will be modifying and analyzing the data of four thousand past Kickstarter projects as you attempt to uncover some of the market trends.


Use conditional formatting to fill each cell in the state column with a different color, depending on whether the associated campaign was "successful," "failed," "cancelled," or is currently "live".


Create a new column at column O called percent funded that uses a formula to uncover how much money a campaign made towards reaching its initial goal.

Use conditional formatting to fill each cell in the percent funded column using a three-color scale. The scale should start at 0 and be a dark shade of red, transitioning to green at 100, and then moving towards blue at 200.



Create a new column at column P called average donation that uses a formula to uncover how much each backer for the project paid on average.


Create two new columns, one called category at Q and another called sub-category at R, which use formulas to split the Category and Sub-Category column into two parts.



Create a new sheet with a pivot table that will analyze your initial worksheet to count how many campaigns were "successful," "failed," "cancelled," or are currently "live" per category.

Create a stacked column pivot chart that can be filtered by country based on the table you have created.






Create a new sheet with a pivot table that will analyze your initial sheet to count how many campaigns were "successful," "failed," "cancelled," or are currently "live" per sub-category.

Create a stacked column pivot chart that can be filtered by country and parent-category based on the table you have created.





The dates stored within the deadline and launched_at columns are using unix timestamps. Fortunately for us, there is a formula out there that can be used to convert these timestamps into a normal date.


Create a new column named Date Created Conversion that will use this formula to convert the data contained within launched_at into Excel's Date format


Create a new column named Date Ended Conversion that will use this formula to convert the data contained within deadline into Excel's Date format





Create a new sheet with a pivot table with a column of state, rows of Date Created Conversion, values based on the count of state, and filters based on parent category and Years.


Now create a pivot chart line graph that visualizes this new table.




Create a report in Microsoft Word and answer the following questions...



What are three conclusions we can make about Kickstarter campaigns given the provided data?
What are some of the limitations of this dataset?
What are some other possible tables/graphs that we could create?
