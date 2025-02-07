<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="repair_number.class.asp" -->
<% 
    
    Dim repair
    Set repair = new repair_number

    Response.write "<h1> Start Test </h1><br>"

    Response.write "Semple operation: 280 - 279.99 <br>"
    Response.write "Semple number: " & 280 - 279.99 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(280 - 279.99) & "<br> <br>"

    Response.write "Semple operation: 309.99 - 310 <br>"
    Response.write "Semple number: " & 309.99 - 310 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(309.99 - 310) & "<br> <br>"

    Response.write "Semple operation: 9.99 - 10 <br>"
    Response.write "Semple number: " & 9.99 - 10 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(9.99 - 10) & "<br> <br>"

    Response.write "Semple operation: 1.885 - 1.884 <br>"
    Response.write "Semple number: " & 1.885 - 1.884 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(1.885 - 1.884) & "<br> <br>"

    Response.write "Semple operation: 4.33 - 4.28 <br>"
    Response.write "Semple number: " & 4.33 - 4.28 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(4.33 - 4.28) & "<br> <br>"

    Response.write "Semple operation: 0.1 + 0.2 <br>"
    Response.write "Semple number: " & 0.1 + 0.2 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(0.1 + 0.2) & "<br> <br>"

    Response.write "Semple operation: 76.07 - 67 <br>"
    Response.write "Semple number: " & 76.07 - 67 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(76.07 - 67) & "<br> <br>"

    Response.write "Semple operation: 262.672 - 262.67 <br>"
    Response.write "Semple number: " & 262.672 - 262.67 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(262.672 - 262.67) & "<br> <br>"

    Response.write "Semple operation: 0.00085022 - 0.00085050 <br>"
    Response.write "Semple number: " & 0.00085022 - 0.00085050 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(0.00085022 - 0.00085050) & "<br> <br>"

    Response.write "Semple operation: 10125.79 - 224.72 <br>"
    Response.write "Semple number: " & 10125.79 - 224.72 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(10125.79 - 224.72) & "<br> <br>"

    Response.write "Semple operation: 10.12579 - 224.72<br>"
    Response.write "Semple number: " & 10.12579 - 224.72 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(10.12579 - 224.72) & "<br> <br>"
    
    Response.write "Semple operation: 2.408940000000010  - 1.8558 <br>"
    Response.write "Semple number: " & 2.408940000000010  - 1.8558 & "<br>"
    Response.write "Repaired number: " & repair.repair_number(2.408940000000010  - 1.8558) & "<br> <br>"
%> 