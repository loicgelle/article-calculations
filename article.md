# How to turn your client's complex spreadsheet into clean calculation code

When it comes to developing a calculation feature in your project, there is always this tricky part of specifying with your Product Owner **the formulas you need to get it right and accurate**. And you may not have the choice regarding the format of this specification, as your Product Owner can decide to give you an existing corporate spreadsheet and say: "there are the formulas, good luck with that".

This is what happened to me in a project of mine a few days ago. A quick look inside the spreadsheet convinced us that the feature we were about to implement was then of O(my god) complexity, and we were even impressed that such a mess managed to get the things done without errors. Finally a little bit of organization and analysis led us to a quite elegant solution, and we ended up with a fine and functional piece of code.

Now it's your turn to get things right in a clean code, but no panic! I will explain to you how you can **turn your horrible Excel formulas into a functional calculation feature** with a little bit of reflexion.

Let's get started!

## The spreadsheet

Let's first take a look at the spreadsheet we were given. It aims at calculating the month-by-month evolution of a life insurance capital depending on given parameters - contract duration, initial capital, exceptional and periodic payments, funds distribution, management costs, tax rates...

Basically, we can identify two areas in the spreadsheet: an area where we can specify the calculation parameters, and the resulting data table that has a fixed number of columns to compute - 58! - and a number of lines that depends on the contract duration.

**TODO : Image here**

## Understand the logic

The goal of this first part is to design our calculation pattern, to choose the data structures accordingly and to write the first functions in the most modular way. To do this, a little bit of reflexion is necessary.

You have to answer one question: in what *order* are the result data calculated? This will allow you to understand the dependencies between the result values and to demystify the apparently magic way Excel updates all the values at the same time, or so.

The general case is as follows: to be computed, a line needs the global calculation parameters and the previous calculated line. You can then design a calculation pattern that proceeds line by line, always keeping trace of the previous line for calculation needs.
