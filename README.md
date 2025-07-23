**Travel Mate** 
- A simple travel logger app that keeps track of all standby flights as well as their cost. 

**Key Features**
- Keeps track of flights on separate sheets for each month as well as their cost.
- Yearly Total Sheet - Keeps track of your total spending for the year 
- Sends automated email with excel sheet for that month to each person who flew standby (employee, parents, guests, etc) 
    - This way, users can have monthly updates on all their flights and cost
- Sends automated yearly email with total travel cost for the year so users can settle expenses


- Logic to add
    - in config.json, create a recipientList and map the traveler name to their email 
    - at the end of each month, send an email only to each recipient that flew that month (ex if PRAKASH didn't fly in July, he won't receive an email in July). Email will contain an excel file attachment for just that month (ex - AA Travel Log July 2025.xlsx)
        - This excel file will contain sheet for just that month (July 2025)
        - It will also have their current yearly total as a separate sheet (Yearly Total) so this way the user can see how much they owe for the year so far 

- Example Email Body Below 

Subject: American Airlines Travel Log July 2025

Body: 

Your monthly American Airlines travel log is attached to this email. Please view the excel file for detailed information. 

Best Regards,
Travel Mate



Your {Year} American Airlines travel log is attached to this email. Please view the excel file for detailed information and settle your expenses. 

Note: If you have already settled your expenses, please disregard this email. (in bold font)

Best Regards,
Travel Mate