# How to turn your client's complex spreadsheet into clean calculation code

When it comes to developing a calculation feature in your project, there is always this tricky part of specifying with your Product Owner the formulas you need to get it right and accurate. And you may not have the choice regarding the format of this specification, as your Product Owner can decide to give you an existing corporate spreadsheet and say: "there are the formulas, good luck with that".

This is what happened to me in a project of mine a few days ago. A quick look inside the spreadsheet convinced me that the feature we were about to implement was then of O(my god) complexity, and I was even impressed that such a mess managed to get the things done without errors. Finally a little bit of organization and analysis led us to a quite elegant solution, and we ended up with a fine and functional piece of code.

Now it's your turn to get things right in a clean code, but no panic! I will explain to you how you can turn your horrible Excel formulas into a functional calculation feature with a little bit of reflexion.

Let's get started!
