# AutoMonthlyBudget
AutoMonthlyBudget generates monthly budget automatically by loading a credit card statement in csv file. It's useful when you track multiple credit card expenses or mange expenses with multiple people. It also enables you to monitor expenses breakdown by categories.

# How to Use
1. Download AutoMonthlyBudget.xlsm under bin folder.
1. Launch AutoMonthlyBudget.xlsm with Microsoft Excel.
1. Go to "List" tab and register payer name.
1. Go to "Card" sheet and register your credit card.
    1. Add your credit card name.
    1. Download a credit card statement in csv file format.
    1. Open the csv file.
    1. Identify which number of column from the left describes Date, Amount and Description and put the number into “Card” Sheet.
1. Go to “Main” Tab.
1. Fill out all field and click “run”.
1. “MMMYYYY” sheet will be generated and expenses are populated.

# Tips
- If you register high frequency item and its category in “LUT” sheet, it will fill out category in MMMYYYY automatically.
- You can add fixed cost and income into “Master” sheet so that you don’t have to fill out every month.

# Note
- Summary will not be completed unless you fill out required field(Amount, Payer and Category).
- You can register up to two payer name for this version.
