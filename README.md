# Excel-Portfolio-Project

Data Analytics with Excel

We have been given Data in Excel format and required to derive some insights. Below is the request from Managment:


		
						
![image](https://user-images.githubusercontent.com/108949165/209159351-09322be8-432c-4f4f-a18f-057d0119ccaf.png)



First we look through the data to understand it. The data is in two separate tabs.


<img width="172" alt="image" src="https://user-images.githubusercontent.com/108949165/209160349-0bb84a33-4bda-4644-b9a2-7cb9f5aeaa47.png">



<img width="201" alt="image" src="https://user-images.githubusercontent.com/108949165/209160556-3252b41f-e825-4511-a781-456582aeeaec.png">




In the first Tab, we would notice that there are some missing cells in column A, the dates are inputed as text so needs to be changed and volume is inputed as text as well. 




**#MAKE THE SHEETS INTO TABLES**
Format both sheet as Tables: Select the data range and hit CTRL + T


#RENAME THE SHEETS
Rename first Sheet as Volume Data, and second sheet as GeoData


**#FILL MISSING CELLS**

Select the entire column using CTLR + SPACEBAR
CTRL + G, CLICK SPECIAL AND SELECT BLANKS TO HIGHLIGHT THE BLANK CELLS
PRESS = AND SELECT THE CELL ABOVE TO INSTRUCT EXCEL TO FILL IN THE VALUE ABOVE



**#CONVERT DATE IN COLUMN B TO PROPER DATE**

Select the entire Date column and use text to column feature under Data ctrl + SPACEBAR

<img width="431" alt="image" src="https://user-images.githubusercontent.com/108949165/209162788-313869a7-f4ab-4bf6-bb1a-89ed14593a63.png">







<img width="434" alt="image" src="https://user-images.githubusercontent.com/108949165/209162866-b8078415-01bc-4fa5-97e8-87d11f592002.png">





<img width="442" alt="image" src="https://user-images.githubusercontent.com/108949165/209162968-0cba26c3-b14a-4bf0-b8f1-f104421a3cc9.png">








make sure you select the right date format







**#CONVERT TEXT IN COLUMN C TO NUMBERS**
Select the entire column CTRL + SPACEBAR
CLICK DATA AND TEXT TO COLUMN, CLICK FINISH TO CONVERT TO NUMBERS.




**#CREATE A GEOID TABLE**


<img width="128" alt="image" src="https://user-images.githubusercontent.com/108949165/209167923-20f035f9-b56d-4df4-a054-0b44cb4c5714.png">




From the email we know that: ( I know NAM ends in 1, EMEA ends in 3 and APAC and LATAM are 2 and 4,) so we need to sum the values to be able to determine the one with the lowest value. We use the sumifs() formula



<img width="214" alt="image" src="https://user-images.githubusercontent.com/108949165/209169861-1af66496-219e-45f2-9b82-e5ae9e575364.png">



#INSERT GEOID IN THE VOLUMEDATA TABLE
Using xlookup :=XLOOKUP([@CLID],GeobyClient[Column1],GeobyClient[GEOID])


#INSERT REGION NAME USING VLOOKUP
=VLOOKUP([@GeoID],Table3[[GEOID]:[GeoNames]],2,FALSE)


<img width="482" alt="image" src="https://user-images.githubusercontent.com/108949165/209177629-40051485-3242-4cc8-99f3-4ad3c98198b5.png">




#CREATE COLUMN FOR QUATER
="Q"&ROUNDUP(MONTH([@Date])/3,0)&" "&YEAR([@Date])


  




