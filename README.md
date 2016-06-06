# AppForSharePointOnlineWebToolkit.Compiled #

[![Build status](https://ci.appveyor.com/api/projects/status/f38f9grt8bkt7hm2/branch/dev?svg=true)](https://ci.appveyor.com/project/justinyoo/appforsharepointonlinewebtoolkit-compiled/branch/dev) | [![](https://img.shields.io/nuget/v/AppForSharePointOnlineWebToolkit.Compiled.svg)](https://www.nuget.org/packages/AppForSharePointOnlineWebToolkit.Compiled) | [![](https://img.shields.io/nuget/dt/AppForSharePointOnlineWebToolkit.Compiled.svg)](https://www.nuget.org/packages/AppForSharePointOnlineWebToolkit.Compiled)

This is a compiled package of [AppForSharePointOnlineWebToolkit](https://www.nuget.org/packages/AppForSharePointOnlineWebToolkit). If you only want code bits, download the original package. Compiled library can be downloaded at [**Compiled** App for SharePoint Web Toolkit (for SharePoint Online)](https://www.nuget.org/packages/AppForSharePointOnlineWebToolkit.Compiled).

> **This package is not maintained by Microsoft.**


## More details ##

* Project site: [http://go.microsoft.com/fwlink/?LinkID=267590&clcid=0x409](http://go.microsoft.com/fwlink/?LinkID=267590&clcid=0x409)


## Additional classes ##

In order for this library to run on DNX framework, `System.Web.dll` dependency should be avoided. The current [AppForSharePointOnlineWebToolkit](https://www.nuget.org/packages/AppForSharePointOnlineWebToolkit) package heavily relies on the `Web.config` file that is tied with `System.Web.dll` and `System.Configuration.dll`. Those `.dll` files are not compatible with DNX environment. Therefore additional configuration file as well as classes are introduced.
 

## `spoconfig.json` ##

`spoconfig.json` is basically a replacement of `Web.config`. More specifically, this replaces SharePoint Online settings from `<appSettings>` on `Web.config`. Its basic structure looks like:

```javascript
{
  "appSettings": [
    { "key": "ClientSigningCertificatePath", "value": "" },
    { "key": "ClientSigningCertificatePassword", "value": "" },


    { "key": "HostedAppHostName", "value": "" },
    { "key": "HostedAppHostNameOverride", "value": "" },

    { "key": "HostedAppName", "value": "" },
    { "key": "HostedAppSigningKey", "value": "" },

    { "key": "ClientId", "value": "" },
    { "key": "ClientSecret", "value": "" },
    { "key": "SecondaryClientSecret", "value": "" },

    { "key": "IssuerId", "value": "" },
    { "key": "Realm", "value": "" }
  ]
}
```


## `ClientContextWrapper` ##

In some unit test scenarios, [`ClientContext`](https://msdn.microsoft.com/en-us/library/ee538685.aspx) is required to be mocked. However, it's not mockable as there is properties or methods marked as `virtual`.

This wrapper class lets developers to mock the `ClientContext` class. This wrapper also provides many asynchronous methods such as `LoadAsync()`, `LoadQueryAsync()` and `ExecuteQueryAsync()` to comply your asynchronous programming approach.


## `ClientContextHelper` ##

As the original [AppForSharePointOnlineWebToolkit](https://www.nuget.org/packages/AppForSharePointOnlineWebToolkit) package has a strong dependency on libraries not compatible with DNX environment, it's impossible to create a `ClientContext` instance in ASP.NET Core applications.

This helper class enables to create a wrapper instance to overcome this issue. This is a simple code snippet to create a `ClientContextWrapper` instance:

```csharp
var helper = new ClientContextHelper();
using (var context = helper.CreateAppOnlyClientContext("http://localhost")
{
  ...
  await context.ExecuteQueryAsync();
  ...
}
```

## Notes for ASP.NET Core Web Applications ##

Even though you import this library from NuGet, the `spoconfig.json` file is not automatically copied to your solution. It should be manually copied for your use. 


### ASP.NET Core RC1 Sample ###

If you want to use this library on your ASP.NET Core RC1 web application, the `spoconfig.json` file should be copied to `wwwroot`; otherwise it will throw an exception.


### ASP.NET Core RC2 Sample ###

If you want to use this library on your ASP.NET Core RC2 web application, the `spoconfig.json` file should be copied to your project root; otherwise it will throw an exception.

In addition to this, your `project.json` should be modified for publishing like:

```json
...

"publishOptions": {
  "include": [
    "wwwroot",
    "Views",
    "appsettings.json",
    "spoconfig.json",
    "web.config"
  ]
},

...
```


## License ##

This package is just a compiled library. Microsoft deserves all rights. For more details, visit the [license page](http://go.microsoft.com/fwlink/?LinkID=267589&clcid=0x409).

