# Midpoint-Prediction-Tool
Automation of Market Pricing. Coded with VBA, Python, and XML for the GUI

One of the processes that used to take a great deal of time daily from our work schedules in the Compensation team was pricing jobs whenever a manager, an HR Business Partner, or a Talent Acquisition partner logged a case to have a job activated for a new location (we're in many countries). The first time I priced a job I had to open several different Excel reports, filter to get the data I needed, populate a template, and then start thinking and making my pricing decisions. Each job would take us about 20 minutes and the reports were massive, which would add to the time while Excel would freeze multiple times during the process. I developed a tool using VBA, Python, and XML for the graphical user interface to automate this, and as a result, we have a more streamlined process that takes less than 2 minutes per job.

It is made available to the whole team through SharePoint, where I load the latest version of all the reports that are used for pricing, along with the latest version of the code so they get all the updates whenever I improve the tool.

The tool is accessed by anyone on the team through an Excel Add-in:<br>
![image](https://github.com/jacksoncaquino/Midpoint-Prediction-Tool/assets/61064363/d438b956-1a54-4fa2-8a9c-a7971a5daede)

It can price one or more jobs at a time. The user can type the job code and compensation market directly on the form, or select a cell or group of cells where the data is:<br>
![image](https://github.com/jacksoncaquino/Midpoint-Prediction-Tool/assets/61064363/ec4e2788-28a1-4295-a734-dcff7dae31de)

VBA then gathers the user input and passes it to Python through a text file with parameters, then calls the Python pricing script.

In summary, the ~1000 lines python code (which is not shared here) includes functions that will do the following:
• Get job family using the job codes<br>
• Get location information from the comp market code<br>
• Verify if we already have jobs priced in the same job family in that location to make sure intra-level progression will be applied<br>
• Calculate the typical progression between levels in that location<br>
• Calculate pay range midpoint through optimal pay range progression<br>
• Get market data available from our Excel reports<br>
• See which level in the job family has the best quality data<br>
• Calculate pay range based on market data<br>
• Get all jobs in our main location (headquarted in the United States) and see which ones are also priced on the target location<br>
• Calculate costs to make sure cost differentials are consistent across jobs in the same geography<br>
• Analyze the data from the above pricing approaches and recommend the final midpoint for the pay range<br>
• Write a log that the Comp analysts can follow and understand the model's recommendations (it's not just a black box that does not justify its decisions)<br>
• Saves the log file for future reference/training/audits<br>

Here's a small snippet of what the code looks like:<br>
![image](https://github.com/jacksoncaquino/Midpoint-Prediction-Tool/assets/61064363/e2e1e1f9-ebce-4f42-a8d3-97439ccabc71)

When the Python script is done pricing the job, it returns it to Excel through the same parameters text file, which then is read by VBA and populates the target cells with the pay range midpoints.

Here's an example of what the text file for the log looks like:<br>
![image](https://github.com/jacksoncaquino/Midpoint-Prediction-Tool/assets/61064363/60f26e8f-ba23-41e6-95b1-828a81999002)

If the user chooses to make changes to the pricing, he or she is more than welcome to, and they can use the log file to keep notes of their changes.
