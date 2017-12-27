## Synopsis
This script collects computer and user information and writes it to the computers attributes.<br>

The following information is collected:<br>
Computer manufacturer, model and serialnumber.<br>
Full name and samAccountName of logged on user.<br>
This information is written to the specified attributes on the <b>computer</b> account.<br>

## Motivation

This makes it easy to see which user last logged on to a computer and what hardware it is running on.<br>
The resulting data in the attributes will be something like:<br>
employeeType = 'Dell Inc. OptiPlex 7020 Serial# AAA123'<br>
info = 'John Doe (jdoe)'<br>

## How to use

There are probably many ways... How I use it:<br>
I have placed the script in SYSVOL, you could place it anywhere it's accessible. Then I have a scheduled task applied via GPO that runs 30 seconds after a user has logged on.
It executes powershell.exe -executionpolicy bypass -file \\local.domain\sysvol\local.domain\scripts\ad-collect.ps1<br>

The context under which the script runs needs permissions to write to the attributes (employeeType and info by default).<br>
I have solved this by running the script under the SYSTEM context and giving SELF permission to write to the attributes in the relevant OUs.

## License

Licensed under the MIT license.