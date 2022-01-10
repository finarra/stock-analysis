# Stock Analysis Project: A detailed 2017-2018 review.

## Overview.
The purpose of this paper is to portrait a detailed analysis of twelve relevant companies dedicated to green energy, and their performance on the stock market for 2017 and 2018, summarized through the Total Daily Volume and their yearly return.

The analysis presented here, was made through VBA code, which was refactored in some extent, in order to search through relevant data for the stocks presented here, but it can be repurposed to analyse any stock or stocks in the stock market and any year, with a very efficient and reasonable execution time just by adding an independent sheet for each year, and selecting the year in the input box to analyse the desired period of time.

## Results.
### 2017
The year 2017 can be considered a great year for the green energy sector stocks included in the analysis, with most of them showing positive returns, except for TERP, which is the only stock with a negative return that year (see chart below).

![2017 return](https://user-images.githubusercontent.com/95982833/148709114-a29bb54c-235d-46b4-9765-8538aeebc573.png)

The top performing stocks of 2017 were DQ, with a return of 199%, followed closely by SEDG with a return of 184%, and in third place ENPH with a return of 130%.FSLR is worth mentioning, becuase it had a return of 101%. Besides these four stocks with a return greater than 100%, the rest had a modest return in comparison.

Interestingly, as it can be noted in the following chart, the Total Daily Volume did not correlate with the yearly return.

![2017 daily volume](https://user-images.githubusercontent.com/95982833/148709275-ac66726e-daec-49e7-8c78-d26d70a1abc7.png)

### 2018
By performing the same analysis to this group of stocks, we can see that it was not a good year for the whole sector, with just RUN and ENPH making a positive return, of 84% and 81.9% respectively.

![2018 return](https://user-images.githubusercontent.com/95982833/148709607-c65f7118-c12c-45ce-8291-e4ce553f0e23.png)

### Code performance
The performance of the refactored code, compared with the original code was slightly better, in the ranges of hundreths of a second (see performance times below).

![VBA_Challenge_2017](https://user-images.githubusercontent.com/95982833/148709758-08162b08-9575-4ab7-94d3-91872cb87e1e.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/95982833/148709766-c5a7faf0-8277-494c-9a41-7b4b6150e12b.png)

However, as the difference may seem negligible, we must take into consideration that only 12 stocks were included in the analysis. If more stocks are considered for analysis, I am sure that the difference in performance may be more significative.

## Summary
The purpose of this project is to obtain a code that can be used to analyse stocks, in a user-friendly manner, and obtaining objective evidence of the performance of a single stock, or a specific market, by obtaining the Total Daily Volume and yearly return, in order to aid in investment strategies.
As the parameters described above are obtained using common and general arithmetic calculations, refactoring code is a useful strategy, because these calculations can be done through the use of "general purpose" code lines.
The complexity of this kind of analysis can be dramatically increased as the volume of data increases. We must take into consideration that the stock market produces a huge amount of information, with transactions everyday around the globe in different markets. However, the advantage of the stock market, is that every stock is identified by a unique ticker, which makes it easy to sort the data relevant for each stock, and the refactored code lines to identify each unique ticker come in handy, and it becomes easier to analyse each sock, even with different parameters other than the ones used in this code.
