[![NuGet](https://img.shields.io/nuget/v/GenericExporter.svg)](https://www.nuget.org/packages/GenericExporter)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

# What is GenericExporter

GenericExporter is a generic excel exporter for collection of any strong typed objects.
GenericExporter uses reflecton to get property name as header in the output file. But you can also provide optional parameters for custom headers and even formatters.
# NuGet
```xml
Install-Package GenericExporter
```
# Getting started with GenericExporter

  * Add dependency injection in Startup: 
  ```xml
     services.AddGenericExporter();
  ```
  * Call export function:
  ```xml
    var items=new List<User>(new User{...});
    var result = _exporter.Export(items);
  ``` 
  * Otional parameters for custom headers and formatters:
  ```xml
    var items=new List<User>(new User{...});
    var result = _exporter.Export(items,new[] {"Id","Full Name","Birthday"},x=>new[] {x.Id,x.Name,x.Birthday.ToShortDateString()});
  ``` 
# License
All source code is licensed under MIT license - http://www.opensource.org/licenses/mit-license.php
