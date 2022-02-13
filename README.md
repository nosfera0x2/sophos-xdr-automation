# Sophos XDR Automation Lab

## Author

* Github: [@spence-rat](https://github.com/spence-rat)
* LinkedIn: [@Spencer_Brown](https://www.linkedin.com/in/spencerbrowntx/)

## About

Runing openAttachment.py has the following results:
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

## Up Next:
1. Automating the UI variables presented based on applications opening for first time (prompting for office365 login, highlighting new features, etc.)
2. Connecting to outlook if it is open, and navigating to the BTH2/2 subfolder consistently. 

Need to make sure that the 2 subfolder is visible:
![image](https://user-images.githubusercontent.com/82817752/153733757-79538bfb-5e2d-4202-b09b-f18e0dee7794.png)

Possible by doing a if/else statement that searches through the print_control_identifiers to test for the '2 - Sophos - Outlook' item. 


