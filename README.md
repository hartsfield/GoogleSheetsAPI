<!-- Package ohsheet is a Go library for reading and writing data to/from Google -->
<!-- Sheets using the Google Sheets API -->

<!-- Copyright (c) 2023 J. Hartsfield -->

<!-- Permission is hereby granted, free of charge, to any person obtaining a copy -->
<!-- of this software and associated documentation files (the "Software"), to deal -->
<!-- in the Software without restriction, including without limitation the rights -->
<!-- to use, copy, modify, merge, publish, distribute, sublicense, and/or sell -->
<!-- copies of the Software, and to permit persons to whom the Software is -->
<!-- furnished to do so, subject to the following conditions: -->

<!-- The above copyright notice and this permission notice shall be included in all -->
<!-- copies or substantial portions of the Software. -->

<!-- THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR -->
<!-- IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, -->
<!-- FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE -->
<!-- AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER -->
<!-- LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, -->
<!-- OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE -->
<!-- SOFTWARE. -->
# DEPRECATED See: (gsheet)[https://github.com/sigma-firma/gsheet] for the same functionality but for human `beiungs"

# OhSheet - Access Google Sheets API via Go

`ohsheet` is a Go module used for accessing the Google Sheets API, 
and reading and writing to spreadsheets.

# Example Usage

### Validating credentials and connecting to the API:

It's important to note that for this module to work properly, you need to 
**enable the sheets API in Google Cloud Services**, and download the 
**credentials.json** file provided in the **APIs and Services** section of the
Google Cloud console.

If you're unsure how to do any of that or have never used a Google Service API 
such as the SheetsAPI or GmailAPI, please see the following link:

https://developers.google.com/sheets/api/quickstart/go

That link will walk you through enabling the sheets API through the Google 
Cloud console, and creating and downloading your `credentials.json` file.

Once you have enabled the API, downloaded the `credentials.json` file and 
placed it in the root directory of your program, you can connect to the API in 
a Go program using the following code:

```
package main

import "github.com/hartsfield/ohsheet"

func main() {                                                                          
        // Connect to the API                                                          
        sheet := &ohsheet.Access{                                                      
                Token:       "token.json",                                             
                Credentials: "credentials.json",                                       
                Scopes:      []string{"https://www.googleapis.com/auth/spreadsheets"}, 
        }                                                                              
        srv := sheet.Connect()                                                         
}
```

### Reading values from a spreadsheet:

```
func main() {                                                                          
        package main

        import (
                "fmt"
                "log"

                "github.com/hartsfield/ohsheet"
        )

        // Connect to the API                                                          
        sheet := &ohsheet.Access{                                                      
                Token:       "token.json",                                             
                Credentials: "credentials.json",                                       
                // You may want a ReadOnly scope here instead
                Scopes:      []string{"https://www.googleapis.com/auth/spreadsheets"}, 
        }                                                                              
        srv := sheet.Connect()                                                         


        // Prints the names and majors of students in a sample spreadsheet:
        // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
        spreadsheetId := "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
        readRange := "Class Data!A2:E"
        
        resp, err := sheet.Read(srv, spreadsheetId, readRange)
        if err != nil {
                fmt.Fatalf("Unable to retrieve data from sheet: %v", err)
        }

        if len(resp.Values) == 0 {
                fmt.Println("No data found.")
        } else {
                fmt.Println("Name, Major:")
                for _, row := range resp.Values {
                        // Print columns A and E, which correspond to indices 0 and 4.
                        fmt.Printf("%s, %s\n", row[0], row[4])
                }
        }
}
```

### Writing values to a spreadsheet:

```
package main

import (
        "fmt"
        "log"

        "github.com/hartsfield/ohsheet"
)

func main() {                                                                          
        // Connect to the API                                                          
        sheet := &ohsheet.Access{                                                      
                Token:       "token.json",                                             
                Credentials: "credentials.json",                                       
                Scopes:      []string{"https://www.googleapis.com/auth/spreadsheets"}, 
        }                                                                              
        srv := sheet.Connect()                                                         

        spreadsheetId := "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"

        // Write to the sheet
        writeRange := "K2"
        data := []interface{}{"test data 3"}
        res, err := sheet.Write(srv, spreadsheetId, writeRange, data)
        if err != nil {
                log.Fatalf("Unable to retrieve data from sheet: %v", err)
        }
        fmt.Println(res)
}
```
