<div align="center">

## Arithmetic operations on memory address pointers

<img src="PIC20096966575076.jpg">
</div>

### Description

Function to enable valid addition and subtraction of unsigned long integers. Treats the passed value as an unsigned long and returns an unsigned long. Allows safe arithmetic operations on memory address pointers. Assumes valid pointer and pointer offset. Rather fast too, has very small performance hit compared to unsafe in-line calculation, even in intensive loops. The code for addition came from PSC (LukeH) but I needed subtraction as well, thought I'd share. Confirmation that this produces valid results would be appreciated.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rde-arithmetic-operations-on-memory-address-pointers__1-70326/archive/master.zip)





### Source Code

<tt>
<p nowrap>
&#160; <br />
<font color="#006600">' Used for unsigned arithmetic</font><br />
<font color="#000099">Private Const</font> <font color="#660000">DW_MSB</font> <font color="#330033">= &H80000000</font> <font color="#006600">' DWord Most Significant Bit</font><br />
&#160; <br />
<font color="#006600">
&#160;' + Sum Unsigned Long ++++++++++++++++++++++<br />
&#160; <br />
&#160;' Enables valid addition and subtraction of unsigned long integers.<br />
&#160;' Treats lPtr as an unsigned long and returns an unsigned long.<br />
&#160;' Allows safe arithmetic operations on memory address pointers.<br />
&#160;' Assumes valid pointer and pointer offset.<br /></font>
&#160; <br />
<font color="#000099">
Private Function</font> <font color="#330033">SumUnsignedLong</font><font color="#000099">(ByVal</font> <font color="#660000">lPtr</font> <font color="#000099">As Long, ByVal</font> <font color="#660000">lOffset</font> <font color="#000099">As Long) As Long<br />
&#160; <br />
&#160; If</font> <font color="#660000">lOffset</font> <font color="#330033">&gt; 0&amp;</font> <font color="#000099">Then<br />
&#160; &#160; &#160;If</font> <font color="#660000">lPtr</font> <font color="#000099">And</font> <font color="#660000">DW_MSB</font> <font color="#000099">Then</font>&#160; &#160; &#160; &#160; &#160; &#160; &#160;<font color="#006600">' if ptr &lt; 0</font><br />
&#160; &#160; &#160; &#160; <font color="#330033">SumUnsignedLong =</font> <font color="#660000">lPtr</font> <font color="#330033">+</font> <font color="#660000">lOffset</font>&#160;<font color="#006600">' ignors &gt; unsigned max (see assumption)</font><font color="#000099"><br />
&#160; <br />
&#160; &#160; &#160;ElseIf (</font><font color="#660000">lPtr</font> <font color="#000099">Or</font> <font color="#660000">DW_MSB</font><font color="#000099">)</font> <font color="#330033">&lt; -</font><font color="#660000">lOffset</font> <font color="#000099">Then</font><br />
&#160; &#160; &#160; &#160; <font color="#330033">SumUnsignedLong =</font> <font color="#660000">lPtr</font> <font color="#330033">+</font> <font color="#660000">lOffset</font>&#160;<font color="#006600">' result is below signed int max</font><font color="#000099"><br />
&#160; <br />
&#160; &#160; &#160;Else</font>&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#006600">' result wraps to min signed int</font><br />
&#160; &#160; &#160; &#160; <font color="#330033">SumUnsignedLong =</font> <font color="#000099">(</font><font color="#660000">lPtr</font> <font color="#330033">+</font> <font color="#660000">DW_MSB</font><font color="#000099">)</font> <font color="#330033">+</font> <font color="#000099">(</font><font color="#660000">lOffset</font> <font color="#330033">+</font> <font color="#660000">DW_MSB</font><font color="#000099">)<br />
&#160; &#160; &#160;End If<br />
&#160; <br />
&#160; ElseIf</font> <font color="#660000">lOffset</font> <font color="#330033">= 0&amp;</font> <font color="#000099">Then</font><br />
&#160; &#160; &#160;<font color="#330033">SumUnsignedLong =</font> <font color="#660000">lPtr</font><font color="#000099"><br />
&#160; <br />
&#160; Else</font> <font color="#006600">'If lOffset &lt; 0 Then</font><font color="#000099"><br />
&#160; &#160; &#160;If (</font><font color="#660000">lPtr</font> <font color="#000099">And</font> <font color="#660000">DW_MSB</font><font color="#000099">)</font> <font color="#330033">= 0&amp;</font> <font color="#000099">Then</font>&#160; &#160; &#160; <font color="#006600">' if ptr &gt; 0</font><br />
&#160; &#160; &#160; &#160; <font color="#330033">SumUnsignedLong =</font> <font color="#660000">lPtr</font> <font color="#330033">+</font> <font color="#660000">lOffset</font> <font color="#006600">' ignors unsigned &lt; zero (see assumption)</font><font color="#000099"><br />
&#160; <br />
&#160; &#160; &#160;ElseIf (</font><font color="#660000">lPtr</font> <font color="#330033">-</font> <font color="#660000">DW_MSB</font><font color="#000099">)</font> <font color="#330033">&gt;= -</font><font color="#660000">lOffset</font> <font color="#000099">Then</font><br />
&#160; &#160; &#160; &#160; <font color="#330033">SumUnsignedLong =</font> <font color="#660000">lPtr</font> <font color="#330033">+</font> <font color="#660000">lOffset</font> <font color="#006600">' result is above signed int min</font><font color="#000099"><br />
&#160; <br />
&#160; &#160; &#160;Else</font>&#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; &#160; <font color="#006600">' result wraps to max signed int</font><br />
&#160; &#160; &#160; &#160; <font color="#330033">SumUnsignedLong =</font> <font color="#000099">(</font><font color="#660000">lOffset</font> <font color="#330033">-</font> <font color="#660000">DW_MSB</font><font color="#000099">)</font> <font color="#330033">+</font> <font color="#000099">(</font><font color="#660000">lPtr</font> <font color="#330033">-</font> <font color="#660000">DW_MSB</font><font color="#000099">)<br />
&#160; &#160; &#160;End If<br />
&#160; End If<br />
End Function<br /></font>
&#160; <br />
<font color="#006600">
&#160;' ++++++++++++++++++++++++++++++++++++++++++<br />
&#160; <br />
&#160;' Extract from the classic ShellSort algorithm<br /></font>
&#160; <br />
<font color="#000099">
  Dim</font> <font color="#660000">s1</font> <font color="#000099">As String,</font> <font color="#660000">lpStr1</font> <font color="#000099">As Long<br />
  Dim</font> <font color="#660000">s2</font> <font color="#000099">As String,</font> <font color="#660000">lpStr2</font> <font color="#000099">As Long</font><br />
&#160; <br />
<font color="#660000">
  lpStr1</font> = <font color="#000099">VarPtr</font>(<font color="#660000">s1</font>)<br />
<font color="#660000">
  lpStr2</font> = <font color="#000099">VarPtr</font>(<font color="#660000">s2</font>)<br />
&#160; <br />
<font color="#660000">
  lpArr</font> = <font color="#000099">VarPtr</font>(<font color="#660000">sArr</font>(<font color="#660000">lb</font>))<br />
&#160; <br />
<font color="#006600">
&#160;'lpLast = lpArr + ((ub - lb) * 4&)
<br />
</font>
<font color="#660000">
&#160;lpLast</font> = <font color="#330033">SumUnsignedLong</font>(<font color="#660000">lpArr</font>, (<font color="#660000">ub - lb</font>) * 4&)<br />
<font color="#006600">
&#160;'...<br /></font>
&#160; <br />
<font color="#000099">
&#160;  CopyMemByV</font> <font color="#660000">lpStr1, lpLast</font>, 4&<br />
&#160; <br />
<font color="#006600">
&#160;  'CopyMemByV lpStr2, lpLast - lRange, 4&
<br />
</font>
<font color="#000099">
&#160;  CopyMemByV</font> <font color="#660000">lpStr2</font>, <font color="#330033">SumUnsignedLong</font>(<font color="#660000">lpLast, -lRange</font>), 4&<br />
&#160; <br />
<font color="#000099">
&#160;  If StrComp</font>(<font color="#660000">s2, s1, eMethod</font>) = <font color="#660000">eComp</font> <font color="#000099">Then</font><br />
<font color="#006600">
&#160;  '...<br /></font>
&#160; <br />
</p>
</tt>

