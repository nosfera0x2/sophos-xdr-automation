# Sophos XDR Automation Lab

For Fellow Sophos Employees:
pocTest.py is the final build of the script.

## Author

* Github: [@spence-rat](https://github.com/spence-rat)
* LinkedIn: [@Spencer_Brown](https://www.linkedin.com/in/spencerbrowntx/)

## About
After a silent install of Office:
1. Run autoOutlookSetup.py to launch Outlook for the first time, click through prompts and import .PST file.

Runing autoPOC.py has the following results:
1. Outlook opens
2. Simulates a person clicking through the interface to open email with attachment.
3. The attachment opens, and a simulation of enabling the malicious macro occurs.
4. Generates this threat graph in the Sophos Threat Analysis Center:

![image](https://user-images.githubusercontent.com/82817752/153731681-ed1fcaff-6247-4ccc-96ba-2184e9c3e72a.png)


## Show your support

Give a ⭐️ if this project helped you!

## Documentation
The online documentation for this project can be found at the link below:
https://pywinauto.readthedocs.io/en/latest/

References:

For killing an open process: https://www.geeksforgeeks.org/how-to-terminate-a-running-process-on-windows-in-python/

For determing if a process is running: https://thinktwisted.com/2020/06/19/how-to-check-is-a-program-or-process-is-running-in-python/

Sophos API: https://github.com/0xbennyv/2020-se-bootcamp-api-101

## Up Next:
1. Automating the UI variables presented based on applications opening for first time (prompting for office365 login, highlighting new features, etc.)
  1a. This has been done during the initial setup after install.
  1b. This needs to also done in the event the image presented already has outlook (with profile imported) already present (an activation dialogue will always show    up).




