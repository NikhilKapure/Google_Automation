Colomun AE1:AG6

From	To	Band
0	    1	  Less than 1 years
1.1 	2	  Less than 2 years
		
2.1	  3	  Less than 3 years
3.1	 10	  More than 3 years

Column I 
Purchase Date
17-Sep-2014

Column V
Vintage (years) =(today()-I2)/365
4.83

Column W
Vintage band (years) =vlookup(V2,$AE$1:$AG$6,3,1)
More than 3 years

Assest Menu -  
Column A149 
=IFERROR(FILTER(Asset_List!A:T, Asset_List!L:L="Service Tag"))

Column A150 
=if(B6="Service Tag", IFERROR(FILTER(Asset_List!A:T, Asset_List!B:B=B8)), IFERROR(FILTER(Asset_List!A:T, Asset_List!L:L=B6)))

Column E2
=TRANSPOSE('Main Menu'!A149:Z149)

column F2
=TRANSPOSE('Main Menu'!A150:Z150)

Column I14
Get Email ID =IFERROR(FILTER('Master Data'!D:D,'Master Data'!A:A=F3))

![Screenshot](/Assets%20MGMT%20Automation/main_menu.png)
