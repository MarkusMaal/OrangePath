# VBA scripting guidelines

Here are some important things to remember when programming using VBA in Sunlight OS.

## Use standardized module names

A standard module for each app should be named in the format `MApp<AppName>`. The AppName should be in PascalCase for consistency.

## Use PascalCase everywhere

Since PowerPoint seems to use PascalCase almost everywhere else when naming stuff, we do recommend sticking with that naming scheme. In pascal case, each word in a variable or function name should be capitalized, including the first word.

## Naming custom macros

Every custom macro should be prefixed with `App<AppName>` unless it's a built-in system application. This helps avoid conflicts with other macros.

## Use built-in functions whenever possible

Do not manually do stuff with window management or file system, use built-in functions for that, because if you don't, stuff will break.
