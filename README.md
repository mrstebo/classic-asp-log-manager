# classic-asp-log-manager
A Log Manager for Classic ASP

# How to use it

Copy the src folder somewhere in your Classic ASP project and put the following `include` in your main entry point.

```vbs
<!--#include FILE='my_folder/classic_asp_log_manager.asp'-->
```

*or*

```vbs
<!--#include VIRTUAL='/my_folder/classic_asp_log_manager.asp'-->
```

And to post a log you can do the following

```vbs
Dim logger : Set logger = LogManager.GetLogger("Test")
Dim log : Set log = (New LogBuilder)_
  .Trace("Trace Message") _
  .Tag("test") _
  .Meta("hello", "world") _
  .Build()
Call logger.Log(log)
```
