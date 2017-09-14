# Am_fs2xlsx
This program aims to capture financial statements of those corporations which are listed in SEC and transform them into XLSX. You can save a lot of time by running this program to obtain the financial statements in the format of xlsx and avoid endless processes of copying and pasting. Enjoy it!

You just need to input several key parameters about the report that you are interested in including TICKER SYMBOL or CIK of the corporation(e.g. Nuan or MSFT), the year in which the report was released(e.g. 2016), the type of the report(10-K or 10-Q), the number of the QUARTERS(1, 2 or 3) if the type of the report is 10-Q into the program. Then the financial statements will be extracted from the report automatically and corresponding xlsx will be generated.

Generally speaking, the program can perform well under most conditions, which means that the xlsx displayed is clear, complete and in order without mistakes. But what is so-called MOST CONDITIONS? MOST CONDITIONS refers to such conditions in which the headers are given clearly by the HTM in the website of SEC. However, sometimes the information about headers can't be extracted from the HTM smoothly and thereby there will be something wrong with the xlsx. I haven't come up with an effective approach to solving this problem whereas I believe that it's just a matter of time to get the solution.

To make the program run well, there are some requirements about the running environment and please make sure that they can be satisfied:
0.  The version of Python is 3.X.
1.  Some packages including: bs4, re, openpyxl, requests, selenium.
2.  Chrome with chromedriver fitting the version of Chrome.

If you have any good suggestions about the algorithm or program, please contact me with E-mail: m18362928852@163.com. Danke schön!
