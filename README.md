Using the Google api analyze e-mails and insert info to Google sheets.

I'm taking stock order confirmation emails and adding the date, ticker, price and amount to Google sheets.

In my case I have one sheet that contains all orders and separate ones for each year as well.  
The script opens the spreadsheet containing all orders, checks on what date the last order was, then using the gmail api gets all confirmation emails after that date and adds their info to necessary spreadsheets.

Sheets are organized as follows:        
date(D.M.Y) stock_ticker price amount

My sheets are in estonian so some code here is also, like regex and so on.

The email bodies look something like that:

<pre>
Kinnitus:                     order on täidetud allpool toodud tingimustel
Kinnituse kuupäev:            2022-12-24 10:00:00
Teostamise kuupäev:           2022-12-24

Orderi kuupäev:               2022-12-24 10:00:00
Orderi viide:                 1234567 (12345678)
Konto nr.:                    EE123456789- Investeerimiskonto
Sümbol:                       TICKER, COMPANY NAME
Börs:                         STOCK EXCHANGE
Kogus:                        1.100
Hind:                         1.1 EUR
Summa:                        1.1 EUR
Teenustasu:                   1.1 EUR
Kokku:                        1.1 EUR

Orderi tüüp:                  limiit
Piirhind:                     1.1 EUR
Kehtivus:                     tühistamiseni
</pre>
