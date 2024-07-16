# Book-Parsing
RPI undergraduate research project with the Fashion Innovation Center. Using Python and OpenAI API, turn book scans into an excel sheet.

Currently, the project uses the following workflow to create the excel sheet:
1. Download the bookscan, and run OCR.
2. Copy the text of the scan into a text file. Make sure the formatting is correct, and that columns are not spilling
   into each other.
3. Run "openai_parsing_stage1.py". This calls the OpenAI API to format the text. Currently it costs roughly $$0.50 to run,
   so do so sparingly. It will write the contents to a file specified at runtime in the terminal.
4. Double check the output, making sure everything looks correct and is in the correct format (see example below).
5. Run "openai_parsing_stage2.py". This is a text parser that uses the output of the first program to write the information
   to an excel sheet. Input file and excel file are specified at runtime.

Example formatting:

INPUT:
```
  ALLIED CORPORATION, CHEMICAL SECTOR, FIBERS DIVISION
  EXECUTIVE OFFICE: 141 1 Broadway, New York, NY 10018
  REGIONAL SALES: 400 N. Tustin Ave. , Santa Ana, CA 92705
  1775 The Exchange, Suite 160, Atlanta, GA 30339
  2100 Fiber Park Dr., Dalton, GA 30720
  Friendship Center Office Park, Suite 108, Greensboro, NC 27409
  64 South Miller Rd., Akron, OH 44313
  Nylon 6 Columbia, SC & Hopewell, VA A.CE, Anso, Anso-Tex,
  Capra/an, Nypel (yarn, monofil & staple)
  Polyester Moncure, NC A.CE. (yarn & monofil)
```

OUTPUT: 
```
  <Company Name> ALLIED CORPORATION
  <Division> CHEMICAL SECTOR, FIBERS DIVISION
  <Executive Office> 141 1 Broadway, New York, NY 10018
  <Regional Sales> 400 N. Tustin Ave., Santa Ana, CA 92705
  <Regional Sales> 1775 The Exchange, Suite 160, Atlanta, GA 30339
  <Regional Sales> 2100 Fiber Park Dr., Dalton, GA 30720
  <Regional Sales> Friendship Center Office Park, Suite 108, Greensboro, NC 27409
  <Regional Sales> 64 South Miller Rd., Akron, OH 44313
  <Fiber> Nylon 6
  <Plant Location> Columbia, SC
  <Plant Location> Hopewell, VA
  <Additional Fibers> A.CE, Anso, Anso-Tex, Capra/an, Nypel (yarn, monofil & staple)
  <Fiber> Polyester
  <Plant Location> Moncure, NC
  <Additional Fibers> A.CE. (yarn & monofil)
```
