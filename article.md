# How to turn your client's complex spreadsheet into clean calculation code

When it comes to developing a calculation feature in your project, there is always this tricky part of specifying with your Product Owner **the formulas you need to get it right and accurate**. And you may not have the choice regarding the format of this specification, as your Product Owner can decide to give you an existing corporate spreadsheet and say: "there are the formulas, good luck with that".

This is what happened to me in a Javascript-coded project of mine a few days ago. A quick look inside the spreadsheet convinced us that the feature we were about to implement was then of O(my god) complexity, and we were even impressed that such a mess managed to get the things done without errors. Finally a little bit of organization and analysis led us to a quite elegant solution, and we ended up with a fine and functional piece of code.

Now it's your turn to get things right in a clean code, but no panic! I will explain to you how you can **turn your horrible Excel formulas into a functional calculation feature** with a little bit of reflexion.

Let's get started!

## The spreadsheet

Let's first take a look at the spreadsheet we were given. It aims at calculating the month-by-month evolution of a life insurance capital depending on given parameters - contract duration, initial capital, exceptional and periodic payments, funds distribution, management costs, tax rates...

Basically, we can identify two areas in the spreadsheet: an area where we can specify the calculation parameters, and the resulting data table that has a fixed number of columns to compute - 58! - and a number of lines that depends on the contract duration.

## Understand the logic

The goal of this first part is to design our calculation pattern in order to choose the data structures accordingly and to write the first functions in the most modular way. To do this, a little bit of reflexion is necessary.

You have to answer one question: in what *order* are the result data calculated? This will allow you to understand the dependencies between the result values and to demystify the apparently magic way Excel updates all the values at the same time, or so.

The general case is as follows: to be computed, a line needs the global calculation parameters and the previous calculated line. You can then design a calculation pattern that proceeds line by line, always keeping trace of the previous line for calculation needs.

Sometimes - and this was the case for our project -, you have a slightly more complicated situation where a line depends on all the previously calculated lines to be computed. For example, your formula for a column may depend on a sum on the last values and not only on the last value. This is the case we will study, and we will thus **perform a line-by-line calculation that keeps trace of all the previous lines**.

Enough analysis for now, we are ready to code the outline of our feature.

## Data structures and code pattern

We could be tempted to imitate the spreadsheet apparent data structure by putting the result values in a matrix. **I strongly advise you against doing that**, as you will end up with an unreadable code full of ``myMatrix[i][j-1]`` that will be horrible to maintain. Our setup will be:
- an object for the input parameters;
- a table of objects for the results, each line corresponding to a calculation row.

This allow you to label the objets' values to clarify your code. Here is an example of parameters object:

```Javascript
var calculationParams = {
  contractDuration: 8,
  ...
  payments: {
    amount: 1000,
    frequency: 6
  }
};
```

You can provide new lines of your results table using a function:

```Javascript
initLine() {
  const emptyLine = {
    month: 0,
    year: 0,
    firstResultGroup: {
      ...
    },
    secondResultGroup: {
      ...
    }
  };
  return emptyLine;
}
```

Now the calculations structurally boil down to adding and filling new lines to your results table. Let's write the main calculation function:

```Javascript
computeResults(calculationParams) {
  const monthsToCompute = calculationParams.contractDuration * 12;
  var resultsTable = [];

  for(var i=0; i<monthsToCompute; i++) {
    computeLine(calculationParams, resultsTable, i);
  }

  return resultsTable;
}
```

and the line calculator:

```Javascript
computeLine(calculationParams, resultsTable, index) {
  var newLine = initLine();
  resultsTable.push(newLine);

  computeFirstColumn(calculationParams, resultsTable, index);
  computeSecondColumn(calculationParams, resultsTable, index);
  // and so on!
}
```

What we just did looks like nothing, but the pattern we created allows you to handle the formulas and the calculation logic separately. You can now write the core calculation functions to compute the columns.

What I suggest is to take a few minutes to analyse the formulas in your spreadsheet and to **write them down in a documentation file in understandable terms**. This extra step gives you a file to keep track of the formulas - essential if you want someone else to understand your code without the spreadsheet - and it fastens the development. Indeed, there is nothing more annoying that starting to code a feature without being 100% sure of what it will look like in the finest details. To convince yourself of the importance of this step, try to imagine what inspires you the most between


```
=SI($A49="";"";SI($B$42;SI($A49>=12*8;SI($B49=DATE(C49;1;1);$B$43*$O49;$B$43*$O49-SOMME.SI.ENS($AF48:$AF$49;$C48:$C$49;C49;$A48:$A$49;">=96")-SOMME.SI.ENS($AY48:$AY$49;$C48:$C$49;C49;$A48:$A$49;">=96"));0);SI($A49>=12*8;SI($B49=DATE(C49;1;1);$B$43*$B$44;$B$43*$B$44-SOMME.SI.ENS($AH48:$AH$49;$C48:$C$49;C49;$A48:$A$49;">=96")-SOMME.SI.ENS($BA48:$BA$49;$C48:$C$49;C49;$A48:$A$49;">=96"));0)))
```

and

```
IF someOption is activated
	IF the current index is >= 96
		resultValue = someParam * someColumn - (sum of someColumn values of this year) - (sum of someOtherColumn values of this year)
	ELSE
		 resultValue = 0
ELSE
	IF the current index is >= 96
		resultValue = someParam * someOtherParam - (sum of someColumn values of this year) - (sum of someOtherColumn values of this year)
	ELSE
		resultValue = 0
```

When coding, the second option is way better! When turning this formula into a function it gives us, without trying to refactor at first:

```Javascript
computeFirstColumn(calculationParams, resultsTable, index) {
  var line = resultsTable[index];
  line.firstColumn = 0;

  var sum1 = 0;
  var sum2 = 0;
  for (var i=index-1; i <= index - 12; i--) {
    if (resultsTable[i].year == line.year) {
      sum1 += resultsTable[i].someColumn;
      sum2 += resultsTable[i].someOtherColumn;
    }
  }

  if (calculationParams.someOption)
    if (index >= 96)
      line.firstColumn = calculationParams.someParam * resultsTable[index - 1].someColumn - sum1 - sum2;
  } else {
    if (index >= 96)
      line.firstColumn = calculationParams.someParam * calculationParams.someOtherParam - sum1 - sum2;
  }
}
```

which is not so obvious when you take a look at the spreadsheet formula...

## Going further

We drew the most general code pattern for this kind of line-by-line calculations. Of course, **the calculations themselves can be optimized, but this is specific to your formulas**. If you look carefully at my previous implementation of *computeFirstColumn*, you can notice that the calculation of the sums inside the function is completely inefficient, as I will need to recompute the sum every time I add a line to my results table.

An optimization workaround could be to create a helper object to keep track of the intermediate results:

```Javascript
computeResults(calculationParams) {
  const monthsToCompute = calculationParams.contractDuration * 12;
  var resultsTable = [];

  var calculationHelper = {
    someColumnSum: 0,
    someOtherColumnSum: 0
  };

  for(var i=0; i<monthsToCompute; i++) {
    computeLine(calculationParams, resultsTable, calculationHelper, i);
  }

  return resultsTable;
}
```

```Javascript
computeLine(calculationParams, resultsTable, calculationHelper, index) {
  var newLine = initLine();
  resultsTable.push(newLine);

  computeFirstColumn(calculationParams, resultsTable, calculationHelper, index);
  computeSecondColumn(calculationParams, resultsTable, calculationHelper, index);
  // and so on!

  updateCalculationHelper(calculationParams, resultsTable, calculationHelper, index);
}
```

and to update them accordingly using a new function. You end up with a clearer and more optimized code:

```Javascript
computeFirstColumn(calculationParams, resultsTable, calculationHelper, index) {
  var line = resultsTable[index];
  line.firstColumn = 0;

  if (calculationParams.someOption)
    if (index >= 96)
      line.firstColumn = calculationParams.someParam * resultsTable[index - 1].someColumn - calculationHelper.someColumnSum - calculationHelper.someOtherColumnSum;
  } else {
    if (index >= 96)
      line.firstColumn = calculationParams.someParam * calculationParams.someOtherParam -  calculationHelper.someColumnSum - calculationHelper.someOtherColumnSum;
  }
}
```

## Conclusion

Some points to finish with:
- This article is not pretending to be an absolute guide to how to run calculations in a project, but it shows you global patterns and a way to reach them. **Do not hesitate to adapt them to your project!**
- **Cut your work into smaller pieces!** Your first goal should be to display an empty results table with the index filled - it allows you to draw and code the general pattern -, then you can proceed column by column. The optimization steps can be included in future iterations.
- If the spreadsheet is way too complex or if you have any doubt, don't hesitate to contact your Product Owner or the spreadsheet's author. A 30-minute long **workshop on the spreadsheet will help you develop your calculation feature faster.**
