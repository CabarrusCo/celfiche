# Celfiche

### About Cabarrus County
---
Cabarrus is an ever-growing county in the southcentral area of North Carolina. Cabarrus is part of the Charlotte/Concord/Gastonia NC-SC Metropolitan Statistical Area and has a population of about 210,000. Cabarrus is known for its rich stock car racing history and is home to Reed Gold Mine, the site of the first documented commercial gold find in the United States.

### About our team
---
The Business & Location Innovative Services (BLIS) team for Cabarrus County consists of five members:

+ Joseph Battinelli - Team Supervisor
+ Mark McIntyre - Software Developer
+ Landon Patterson - Software Developer
+ Brittany Yoder - Software Developer
+ Marci Jones - Software Developer

Our team is responsible for software development and support for the [County](https://www.cabarruscounty.us/departments/information-technology). We work under the direction of the Chief Information Officer.

### About
---
Celfiche is a Playwright Client built in Go for turning Excel files into Laserfiche forms. At Cabarrus County, we use Laserfiche variables mostly as a way to collect data from an enhanced front end(I.E. either bootstrap inputs, chatbots, etc), the variables that are collected are usually repetitive and used only for transferring data in the forms lifecycle. Having the ability to edit the forms in excel gives us the ability to use iterators, etc, and all the power that excel offers(Find replace, etc).

### Download the package
---
```
go get github.com/CabarrusCo/celfiche
```

### Spin up a new client
When spinning up a new client, you must pass the URL of your forms home screen as a parameter, as well as wether or not you want to run the browser in headless mode or not.

```
	cf, err := celfiche.NewClient("https://formsserver/forms/", false)
	if err != nil {
		fmt.Println(err)
		return
	}
```

### Logging in
After spinning up a new client, you have to instruct Fischex to login. It's up to you to store your credentials securely.

```
	username := "MySercurelyStoredUsername"
	password := "MySecurelyStoredPassword"

	err = cf.Login(username, password)
	if err != nil {
		fmt.Println(err)
		return
	}
```

### Converting Excel to a form
After logging in, you can then convert excel to a form using the ConvertExcel function. The ConvertExcel function takes three parameters, the form URL, the path to the Excel file, and a wait time pause. Use the wait time pause only if you want to slow things down when running headless mode false. Example

```
	err = cf.ConvertExcel("https://formsserver/Forms/design/layout/540", `C:\form.xlsx`, 0)
	if err != nil {
		fmt.Println(err)
		return
	}
```

### Stopping the instance
After the Excel is converted, you can stop the instance by using the Stop function.

```
	err = cf.Stop()
	if err != nil {
		fmt.Println(err)
		return
	}
```

### Full working example

```
package main

import (
	"fmt"
	"os"

	"github.com/CabarrusCo/celfiche"
)

func main() {
	cf, err := celfiche.NewClient("https://formsserver/forms/", false)
	if err != nil {
		fmt.Println(err)
		return
	}

        defer cf.Stop()

	username := "MySecurelyStoredUsername"
	password := "MySecurelyStoredPassword"

	err = cf.Login(username, password)
	if err != nil {
		fmt.Println(err)
		return
	}

	err = cf.ConvertExcel("https://formserver/Forms/design/layout/540", `C:\form.xlsx`, 0)
	if err != nil {
		fmt.Println(err)
		return
	}

}
```

### The Excel file
---
Currently, the Excel file must be structured a specific way. The first row must be headers. The fields in the Excel file are Label Name, Variable Name, Class Name, Type, Multi Line height, Options, Iteration. Only Excel 2007 and later is supported.

### Using Iteration
---
You can use the iteration feature to create data sets off one row. To do that, set the iterator number in the iterator cell to a number, then use ${i} to inject the iterator.

### Sample excel files
---
Two sample excel files have been added into this repo as an example

### Other
---
This project is still in it's infancy stages. Please check back often for updates as we find/address issues! 
