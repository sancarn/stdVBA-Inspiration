Author: LaVolpe
Src: https://www.vbforums.com/showthread.php?868555-vb6-Project-Scanner



Might be useful to some to help clean up their code before they post it out to the world? A project that started with an experiment in parsing VB files, but evolved to what we have here.

Posts related to this version start after post #215, page 6.

This project has two parts

1) Scan any VB6 project and quickly display a summary of declarations, procedures, controls used, and more. This can be used for a quick review of your own projects or those you download from sites like this one.

2) A deeper analysis of the code content within a project. It currently offers several validation checks. These checks can help tidy up your code and maybe even pinpoint a problem or two. Here is a brief description of the various scans.

- Code files (forms, classes, etc) that do not use Option Explicit
- Procedures with no executable code
- Procedures that contain active End or Stop statements
- Zombie declarations and procedures. Items that are created/declared but not referenced within your code
- VarType checks. Items not declared with a variable type and default to Variant
- Strings that are duplicated within your code. These can be consolidated to constants
- Declarations that are duplicated
- Variant functions used in place of string functions, i.e., Trim() vs. Trim$()
- OPC (other people's code) checks to highlight code that could potentially be used to modify your registry. Also, shows you what, within your code, would be flagged if someone used this tool to scan your project.

Many of the checks could produce false positives. So, consider these checks as informational only. The help menu in the project offers more details on these checks along with known false positives. One thing this project does not do is to look for undeclared variables. Adding Option Explicit statements to your code files will help you in that case. One thing Option Explicit can't help with that this project can, is the dreaded dead declaration that results in a typo with the ReDim statement. Following code creates 2 declared & arrayed variables. In this example, the second one (highlighted in blue) is a typo, but valid. Chances are that one of these would not be used within the routine, due to the typo. In that case, one of them would be reported as a zombie declaration.

![Sample](https://www.vbforums.com/attachment.php?attachmentid=175765&d=1586024643)