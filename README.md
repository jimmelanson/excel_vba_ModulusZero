# excel_vba_ModulusZero
An alternative to the limited vba Mod() function.

One of the (multiple) probelms with Excel VBA's Mod() function is that it does not support decimals.
Mod() will evaluate a single or double as an integer.

Since I hate microsoft telling me how to do things, I decided to create a way around that failure
with the native Mod() function. That was why I created ModulusZero().

This method does not return the modulus of two numbers, it simply returns TRUE if the modulus is
zero and FALSE if it is anything else. This is particularly handy if you are evaluationg a large
range of numbers against a divisor.

Here is how you would use it:

**In VBA**

Ex 1.
```
If ModulusZero(Range("A1").Value, 0.25) = True Then
    
  ... do nifty stuff with Range("A1").Value
      
end if
```

Ex 2.
```
Dim i as Long
Dim longCount as Long
For i = 1 to 5000
    If ModulusZero(Cells(i, 1).Value, 0.125) = True Then
        longCount = longCount + 1
    End If
Next i
Debug.Print "There were " & longCount & " values that were equally divisible by 0.125"
```

<br><br>
**As A Formula**
You can also use this method as an Excel formula:
<code>=ModulusZero(B12, $A$1)</code>

<br><br>
**To use this method**
You can copy and paste the code into your workbook. You can also download the *.bas file
and import it into your workbook project.

<br><br>
**Known MS Issues with Excel VBA**
https://github.com/jimmelanson/excel_vba_known_microsoft_issues

   
