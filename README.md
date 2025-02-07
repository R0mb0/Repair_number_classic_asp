# Repair number Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/ec8f3d42d5864719ab9eac9afe3da97b)](https://app.codacy.com/gh/R0mb0/Repair_number_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Repair_number_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Repair_number_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## Cases coverage 

```
Semple operation: 280 - 279.99
Semple number: 9,99999999999091E-03
Repaired number: 0,001

Semple operation: 309.99 - 310
Semple number: -9,99999999999091E-03
Repaired number: -0,001

Semple operation: 1.885 - 1.884
Semple number: 1,00000000000011E-03
Repaired number: 0,001

Semple operation: 4.33 - 4.28
Semple number: 4,99999999999998E-02
Repaired number: 0,05

Semple operation: 0.00085022 - 0.00085050
Semple number: -2,80000000000072E-07
Repaired number: -0,00000028
```

## `repair_number.class.asp`'s avaible functions

- Function to check and repair a number -> ` Public Function repair_number(number)`

## How to use

> From `Test.asp`

1. Initialize the class
  ```
  <%@LANGUAGE="VBSCRIPT"%>
  <!--#include file="repair_number.class.asp" -->
  <% 
    
      Dim repair
      Set repair = new repair_number
  ```

2. Use the class
  ```
     Response.write "Semple operation: 280 - 279.99 <br>"
     Response.write "Semple number: " & 280 - 279.99 & "<br>"
     Response.write "Repaired number: " & repair.repair_number(280 - 279.99) & "<br> <br>"
   %>
  ```
