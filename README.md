# TG_Assignment_T2
- Export Excel from jsonFile provided
- It is a console application
- The projects are in .Net5.
- I'm using OpenXml SDK for excel generation

## Prerequisites:
1. Visual Studio 2019
2. .Net5 installed

## How to run:
1. Open the solution
2. Run the default console project by the name TG_Reporting
3. Follow the steps in console. Just need to press "Enter"
4. Generated excel is saved in the bin folder in "output" folder. The path of the latest generated excel is printed in the console as well along with the complete path

## Run Unit tests:
1. Go to TG_Reporting.Tests project
2. Run all tests

## My Approach:
- My first task was to convert the json data into a favourable format and then transform it into the model needed for the excel format provided. So I'm firt deserializing the json into a C# object.
- *Creating Excel*: Exporting the data for creating a basic excel with header row is straight forward using the openXML sdk. After that applying various formatting and was complex.
- **Conditional Formating**: For setting the alternative row background I used conditional formatting. But doing so from code was tricky. So, I did the formating in the excel manually and then opened the excel in "Open xml Productivity Tool". Thereby understanding the conditional formatting logic needed.
- **Date Format**: I tried using the same approach for setting the format of the date columns to "Date". But it did not work as expected.

## Improvements 

There are several things we could enhance:

- **Error Handling**: We can add custom error handling according to the requirement. Currently I'm just returning empty path if any exception occurs while creating the excel. In a real world application, we would be returning some sort of message and logging the same.
- **UX**: I've used a console application for running the logic. Same could have been done through a WPF/Web app/Web API. Which would open ways to send different types of data and testing
- **Testing**:  File IO and Exception tests should be added. I've added only the file exists tests. Some tests could be added to validate the data exported in excel. e.g. We can add test to verify the headers of or the number of rows exported in the output file.

## Comments
I have added some comments, notes, etc every where in code to explain my mindset while taking that decision.
Usually I want my code to be self-explanatory and avoid comments unless absolutely needed. 
My approach in this project was to start with minimum required and then extend when needed.

Looking forward to the feedback. Many Thanks!
