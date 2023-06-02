## Tasks

The robot is split into two tasks, meant to run as separate steps in Control Room. The first task generates (produces) data, and the second one reads (consumes) and processes that data.

### The first task (the producer)

- Load the Excel file with a list of customers (downloads it from Sharepoint?)
- Splits the Excel file into work items for the consumer

### The second task (the consumer)

- Logs into Power BI
- Handle each work item:
  - Checks if data for customer in Power BI is up to date
  - Downloads minutes report in .xlsx
  - Format the data from the .xlsx and create pivot table
  - Saves file and send it to SDM by email