# How to Turn Your Client's Complex Spreadsheet into Clean Calculation Code

When it comes to developing a calculation feature in your project, there is always this tricky part of specifying with your Product Owner **the formulas you need to get it right and accurate**.
And you may not have the choice regarding the format of this specification, as your Product Owner can decide to give you an existing corporate spreadsheet and say: "here are the formulas, good luck with that".

This is what happened to me in a Javascript-coded project of mine a few days ago.
A quick look inside the spreadsheet convinced us that the feature we were about to implement was then of O(my god) complexity, and we were even impressed that such a mess managed to get everything done without errors.
Finally a little bit of organization and analysis led us to a quite elegant solution, and we ended up with a fine and functional piece of code.

Now it's your turn to get things right in a clean code, but no panic! I will explain to you how you can **turn your horrible Excel formulas into a functional calculation feature** with a little bit of thinking.

Let's get started!

## The Example Spreadsheet

Let's describe the example spreadsheet.
It aims at forecasting the month-by-month evolution of a set of parameters - temperature, species distribution, air quality,... - in an ecological niche depending on given the initial state of the environment and some calculation parameters. This evolution can be calculated according to a model that describes the interation between the different species and between the species and their environment.

Basically, we can identify two areas in the spreadsheet: an area where we can specify the calculation parameters, and the resulting data table that has a fixed number of columns to compute - the parameters that we want to compute - and a number of lines that depends on the forecast duration.

## Understand the Logic

The goal of this first part is to design our calculation pattern in order to choose the data structures accordingly and to write the first functions in the most modular way.
To do this, a little bit of thinking is necessary.

You have to answer one question: in what *order* is the result data calculated? This will allow you to understand the dependencies between the result values and to demystify the apparently magic way Excel updates all the values at the same time.

The general case is as follows: to be computed, a line needs the global calculation parameters and the previous calculated line.
You can then design a calculation pattern that proceeds line by line, always keeping trace of the previous line for calculation needs.

Sometimes - and it is the case for our example spreadsheet -, you have a slightly more complicated situation where a line depends on all the previously calculated lines to be computed.
For example, your formula for a column may depend on a sum on the last values and not only on the last value.
This is the case we will study, and we will thus **perform a line-by-line calculation that keeps track of all the previous lines**.

Enough analysis for now, we are ready to code the outline of our feature.

## Data Structures and Code Pattern

We could be tempted to imitate the spreadsheet's apparent data structure by putting the result values in a matrix.
**I strongly advise you against doing that**, as you will end up with an unreadable code full of ``myMatrix[i][j-1]`` that will be horrible to maintain.
Our setup will be:
- an object for the input parameters;
- a table of objects for the results, each line corresponding to a calculation row.

This allow you to label the objets' values to clarify your code.
Here is an example of parameters object:

```Javascript
var calculationParams = {
  forecastDurationInYears: 8,
  foodResources: {
    percentageForA: 10,
    percentageForB: 13
  },
  predation: {
    rateAB: 0.1,
    rateBA: 0,
    considerAB: true,
    considerBA: false
  },
  ...
  initialState: {
    month: 0,
    temperature: 28,
    speciesA: {
      population: 10000,
      ...
    },
    speciesB: {
      population: 700,
      ...
    },
    ...
  }
};
```

You can provide new lines to your results table using a function:

```Javascript
function initLine() {
  const emptyLine = {
    month: 0,
    temperature: 0,
    speciesA: {
      population: 0,
      ...
    },
    speciesB: {
      population: 0,
      ...
    }
  };
  return emptyLine;
}
```

Now the calculations structurally boil down to adding and filling new lines to your results table.
Let's write the main calculation function:

```Javascript
function computeResults(calculationParams) {
  const monthsToCompute = calculationParams.forecastDurationInYears * 12;
  var resultsTable = [];
  resultsTable.push(calculationParams.initialState);

  for(var i=1; i<monthsToCompute; i++) {
    computeLine(calculationParams, resultsTable, i);
  }

  return resultsTable;
}
```

and the line calculator:

```Javascript
function computeLine(calculationParams, resultsTable, index) {
  var newLine = initLine();
  resultsTable.push(newLine);

  computePopulationEvolutionA(calculationParams, resultsTable, index);
  computePopulationEvolutionB(calculationParams, resultsTable, index);
  // and so on!
}
```

What we just did looks like nothing, but the pattern we created allows you to handle the formulas and the calculation logic separately.
You can now write the core calculation functions to compute the columns.

What I suggest is to take a few minutes to analyse the formulas in your spreadsheet and to **write them down in a documentation file in _understandable_ terms**.
This extra step gives you a file to keep track of the formulas - essential if you want someone else to understand your code without the spreadsheet - and it fastens the development.
Indeed, there is nothing more annoying that starting to code a feature without being 100% sure of what it will look like in the finest details.
To convince yourself of the importance of this step, try to imagine what inspires you the most between


```
=IF($B$4;IF($E50>$B$11;($E$2-$G$3*SUMIFS($H$10:$H49;$A$10:$A49;">$A50-36"))*$F50;0);IF($F50>0;($I$5-$I$6*(SUMIFS($H$10:$H49;$X50;true)-SUM($I$10:$I49)));0))
```

and

```
IF take predation A -> B into account
	IF population A > predation threshold
		populationEvolutionB = (reproductionRateA - predationRateAB * (sum on all the births for species B on the past 3 years)) *  population B
	ELSE
		 populationEvolutionB = 0
ELSE
	IF population B > 0
		populationEvolutionB = (birth rate B - death rate B) * ((sum on all the births of B in viable past months) - (sum on all the deaths of B over the past months))
	ELSE
		populationEvolutionB = 0
```

When coding, the second option is way better! When turning this formula into a function it gives us, without trying to refactor at first:

```Javascript
function computePopulationEvolutionB(calculationParams, resultsTable, index) {
  var line = resultsTable[index];
  line.speciesB.evolution = 0;

  var sumBirths3Years = 0;
  var sumBirthsViable = 0;
  var sumDeaths = 0;
  for (var i=0; i < index; i++) {
    var loopLine = resultsTable[i];
    if (loopLine.month >= line.month - 36)
      sumBirths3Years += loopLine.speciesB.births;
    if (loopline.speciesB.viableBirth)
      sumBirthsViable += loopLine.speciesB.births;
    sumDeaths += loopLine.speciesB.deaths;
  }

  if (calculationParams.predation.considerAB)
    if (line.speciesA.population > calculationParams.predation.threshold)
      line.speciesB.evolution = line.speciesB.population * (calculationParams.reproduction.rateA - calculationParams.predation.rateAB * sumBirths3Years);
  else
    if (line.speciesB.population > 0)
      line.speciesB.evolution = (line.speciesB.birthRate - line.speciesB.deathRate) * (sumBirthsViable - sumDeaths);
}
```

which is not so obvious when you take a look at the spreadsheet formula...

## Going Further

We drew the most general code pattern for this kind of line-by-line calculations.
Of course, **the calculations themselves can be optimized, but this is specific to your formulas**.
If you look carefully at my previous implementation of *computePopulationEvolutionB*, you can notice that the calculation of the sums inside the function is completely inefficient, as I will need to recompute the sum every time I add a line to my results table.

An optimization workaround could be to create a helper object to keep track of the intermediate results:

```Javascript
function computeResults(calculationParams) {
  const monthsToCompute = calculationParams.forecastDurationInYears * 12;
  var resultsTable = [];
  resultsTable.push(calculationParams.initialState);

  var calculationHelper = {
    sumBirthsViableB: 0,
    sumDeathsB: 0,
    sumBirths3YearsB: 0
  };

  for(var i=0; i<monthsToCompute; i++) {
    computeLine(calculationParams, resultsTable, calculationHelper, i);
  }

  return resultsTable;
}
```

```Javascript
function computeLine(calculationParams, resultsTable, calculationHelper, index) {
  var newLine = initLine();
  resultsTable.push(newLine);

  computePopulationEvolutionA(calculationParams, resultsTable, index);
  computePopulationEvolutionB(calculationParams, resultsTable, index);
  // and so on!

  updateCalculationHelper(calculationParams, resultsTable, calculationHelper, index);
}
```

and to update them accordingly using a new function.
You end up with a clearer and more optimized code:

```Javascript
function computePopulationEvolutionB(calculationParams, resultsTable, calculationHelper, index) {
  var line = resultsTable[index];
  line.firstColumn = 0;

  if (calculationParams.predation.considerAB)
    if (line.speciesA.population > calculationParams.predation.threshold)
      line.speciesB.evolution = line.speciesB.population * (calculationParams.reproduction.rateA - calculationParams.predation.rateAB * calculationHelper.sumBirths3YearsB);
  else
    if (line.speciesB.population > 0)
      line.speciesB.evolution = (line.speciesB.birthRate - line.speciesB.deathRate) * (calculationHelper.sumBirthsViableB - calculationHelper.sumDeathsB);
}
```

## Conclusion

Some points to finish with:
- This article does not claim to be an absolute guide to how to run calculations in a project, but it shows you global patterns and a way to reach them.
**Do not hesitate to adapt them to your project!**
- **Cut your work into smaller pieces!** Your first goal should be to display an empty results table with the index filled - it allows you to draw and code the general pattern -, then you can proceed column by column.
The optimization steps can be included in future iterations.
- If the spreadsheet is way too complex or if you have any doubt, don't hesitate to contact your Product Owner or the spreadsheet's author.
A 30-minute long **workshop on the spreadsheet will help you develop your calculation feature faster.**
