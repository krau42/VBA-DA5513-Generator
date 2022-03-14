# DA5513 Generator
Here is my first attempt with macros and messing with VBA to generate a DA5513: Key Control Register and Inventory

---

## General Rules:
- Buttons tell you to click
- Blue cells correlate to the steps listed in "Filling in the Excel"
- User inputed data goes into green cells

## Setting Up:
1. Download a blank DA Form 5513 from https://armypubs.army.mil/ProductMaps/PubForm/Details.aspx?PUB_ID=1002447
2. Set Adobe as the default to open PDF's
3. Enable Macros in Excel File

## Filling in the Excel:
1. Select the blank DA Form 5513 as your template (button labled "1. Click to select File")
  * Once selected, the file path will populate the orange tinted cell (D2)
2. Select the folder you would like it saved to (button labled "2. Click to select Folder")
  * Once selected, the folder path will populate the orange tinted cell (D3)
3. Input Unit/Activity data in adjecent cell (C5)
4. Input From Date in adjacent cell (C6)
5. Input To Date in adjacent cell (E6, leave blank if necessary)
6. Input the name you would like your PDF to be saved as in adjacent cell (G6)
7. Input Item Names below (cells C8-C111, leave blank if necessary)
8. Input serial numbers below (cells C8-C111, leave blank if necessary)
9. Input quantities below (cells C8-C111, leave blank if necessary)
* *** CAUTION *** Before proceding to step 10, ensure you will not need your computer for the next 5 minutes. The macro is based on key strokes and
  and any inputs will most likely disrupt the process. If you accidentally click Button 10, try pressing the ESC key to cancel the macro. 
10. Click to generate your DA 5513. Do not touch your computer or click off the PDF.
11. Click for long term key issue. This will make a Word Doc with each page having its own unique DA5513. 
