<div align="center">

## Calculate nth Row in Excel


</div>

### Description

Provide understanding of how to calculate every nth row (e.g. every other row, every 35th row, etc.) in Excel, even when the data does not begin in cell A1.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2007-08-02 15:10:52
**By**             |[Todd Benson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/todd-benson.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VBA MS Excel
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Calculate\_207810822007\.zip](https://github.com/Planet-Source-Code/todd-benson-calculate-nth-row-in-excel__1-69095/archive/master.zip)





### Source Code

<p>Sometimes, requested data is delivered to me in Excel, but not in tabular
format. Instead, it is formatted like a report where subtotals are already
calculated and the values are inserted. The grand total (sum of the subtotals),
however, is not included.</p>
<p>The format is like the following:</p>
<table border="1" width="26%">
	<tr>
		<td width="65">Group1</td>
		<td width="76">Detail1</td>
		<td>ValueA</td>
	</tr>
	<tr>
		<td width="65">Group1</td>
		<td width="76">Detail2</td>
		<td>ValueB</td>
	</tr>
	<tr>
		<td width="65">Group1</td>
		<td width="76">Detail3</td>
		<td>ValueC</td>
	</tr>
	<tr>
		<td width="65">Group1</td>
		<td width="76">Detail4</td>
		<td>ValueD</td>
	</tr>
	<tr>
		<td width="65">Group1</td>
		<td width="76">Subtotal1</td>
		<td>SubtotalValue1</td>
	</tr>
	<tr>
		<td width="65">Group2</td>
		<td width="76">Detail1</td>
		<td>ValueE</td>
	</tr>
	<tr>
		<td width="65">Group2</td>
		<td width="76">Detail2</td>
		<td>ValueF</td>
	</tr>
	<tr>
		<td width="65">Group2</td>
		<td width="76">Detail3</td>
		<td>ValueG</td>
	</tr>
	<tr>
		<td width="65">Group2</td>
		<td width="76">Detail4</td>
		<td>ValueH</td>
	</tr>
	<tr>
		<td width="65">Group2</td>
		<td width="76">Subtotal2</td>
		<td>SubtotalValue2</td>
	</tr>
</table>
<p>As you can see there are subtotals on lines 5 and 10. I want to sum them into
a grand total. Simply using the =Sum() function on the column will not get me
the result I want since it will likely sum the whole column, including the
subtotals, or require me to loop through every fifth line to create a string to
populate the =Sum() function (ex., =Sum(E5,E10,E15,En)). Why? Because the cell
in which I place that example can throw an error due to the length of the
formula string (imagine if I had 1,000-plus groups)</p>
<p>So what does a person do? A mix of Excel functions will allow me to do this.
Functions like SumProduct(), Mod() and Row().</p>
<p>I will provide an example that has data that doesn’t start in cell A1, but
somewhere else on the sheet, like C3. That way you can learn the flexibility of
this solution. I also chose an example that requires me to use every fifth row.</p>
<p>I freely admit that you can find this solution elsewhere on the net. However, it
is difficult to find. Therefore, I thought I’d help out the community by
providing it here.</p>
<p>The attached Excel example follows the example above with one difference: There
are four groups instead of two.</p>
<p>THE SETUP: For attached Excel example, my value data that I want to sum starts
in Column E, Row 3. It ends in Column E, Row 22. Each Group has four values in
successive rows followed by a subtotal of the group on the fifth row. That
equals four subtotals (E7, E12, E17 and E22).</p>
<p>THE SOLUTION: I want the grand total to appear after that last subtotal. The
last subtotal is located at E22. Therefore, the grand total will appear in E23.
The formula in E23 is: =SUMPRODUCT($E$3:$E$22,IF(MOD(ROW($E$3:$E$22)-2,5),0,1)).
When entered manually, before you leave the cell, be sure to use the keyboard
combination: &#60CTRL&#62+&#60SHIFT&#62+&#60ENTER&#62. Otherwise, this array calculation will not
work. In VB, the code to enter is
<blockquote>oSheet.Cells(lastRow, lngCol).FormulaArray = _ <br>
SUMPRODUCT($E$3:$E$22,IF(MOD(ROW($E$3:$E$22)-2,5),0,1))</blockquote>
where oSheet is a variable for the Excel Sheet Object.</p>
<p>THE EXPLANATION: Yep, you’ve seen this MOD() function before. Many use it to
color every nth row in Excel. In my example, we sum every fifth row (see the “5”
in the MOD() function?). Let’s break this formula down for better understanding:</p>
<blockquote>SUMPRODUCT = A function that sums the products resulting from multiplying values
from one array with corresponding values from a second array. This is seen as
SUMPRODUCT(array1, array2). In the example, the first array is declared as
$E$3:$E$22. But please remember that I only want every fifth row. That is why
the second array in the example uses the MOD() function with 5 as the divisor.</blockquote>
<blockquote>MOD = The MOD() function, nested in the IF() function, answers the question for
the same array ROW($E$3:$E$22): ‘Is this row NOT one of the every fifth row?’
Confusing logic in that question, eh? If the answer is true (“True, it is NOT a
fifth row) then it multiples the corresponding number from the first array by
the value of the IF()’s truepart, or 0. Otherwise, because it is a fifth row
(“False, it is a fifth row), the value from the first array is multiplied by the
IF()’s falsepart, or 1. </blockquote>
<p>Do you recall that I mentioned I started the data on row three ($E$3)? Great!
Well, the MOD() function needs to take that into account for it to work
properly. In this case, there happen to be two rows (row 1 and row 2) above our
starting row. When you look at the MOD() function in the example formula you can
see I accounted for this by including “-2” following the array.</p>
<p>Finally, SUMPRODUCT sums the products of all those calculations (hence, the name
of the function).</p>
<p>EASY EXAMPLE: Let’s use the sample I provided at the top of this article.</p>
<p>Find the products:</p>
<table border="1" width="56%">
	<tr>
		<td width="89">ValueA</td>
		<td width="41" align="center">3</td>
		<td width="336" bgcolor="#FF0000">Not a fifth row, so multiply by 0.
		Product equals</td>
		<td align="center">0</td>
	</tr>
	<tr>
		<td width="89">ValueB</td>
		<td width="41" align="center">2</td>
		<td width="336" bgcolor="#FF0000">Not a fifth row, so multiply by 0.
		Product equals</td>
		<td align="center">0</td>
	</tr>
	<tr>
		<td width="89">ValueC</td>
		<td width="41" align="center">2</td>
		<td width="336" bgcolor="#FF0000">Not a fifth row, so multiply by 0.
		Product equals</td>
		<td align="center">0</td>
	</tr>
	<tr>
		<td width="89">ValueD</td>
		<td width="41" align="center">6</td>
		<td width="336" bgcolor="#FF0000">Not a fifth row, so multiply by 0.
		Product equals</td>
		<td align="center">0</td>
	</tr>
	<tr>
		<td width="89">Subtotal1</td>
		<td width="41" align="center">13</td>
		<td width="336" bgcolor="#008000">This is a fifth row, so multiply by 1
		to equal </td>
		<td align="center">13</td>
	</tr>
	<tr>
		<td width="89">ValueE</td>
		<td width="41" align="center">4</td>
		<td width="336" bgcolor="#FF0000">Not a fifth row, so multiply by 0.
		Product equals</td>
		<td align="center">0</td>
	</tr>
	<tr>
		<td width="89">ValueF</td>
		<td width="41" align="center">1</td>
		<td width="336" bgcolor="#FF0000">Not a fifth row, so multiply by 0.
		Product equals</td>
		<td align="center">0</td>
	</tr>
	<tr>
		<td width="89">ValueG</td>
		<td width="41" align="center">3</td>
		<td width="336" bgcolor="#FF0000">Not a fifth row, so multiply by 0.
		Product equals</td>
		<td align="center">0</td>
	</tr>
	<tr>
		<td width="89">ValueH</td>
		<td width="41" align="center">3</td>
		<td width="336" bgcolor="#FF0000">Not a fifth row, so multiply by 0.
		Product equals</td>
		<td align="center">0</td>
	</tr>
	<tr>
		<td width="89">Subtotal2</td>
		<td width="41" align="center">11</td>
		<td width="336" bgcolor="#008000">This is a fifth row, so multiply by 1
		to equal </td>
		<td align="center">11</td>
	</tr>
	<tr>
		<td width="89">Grand Total</td>
		<td width="41" align="center">**</td>
		<td width="336">Sum the products: (0+0+0+0+13+0+0+0+0+11) =</td>
		<td align="center">24</td>
	</tr>
</table>
<p>** Here the SUMPRODUCT() formula would be inserted and the result would show
24.</p>

