# Simple-file-searcher

A simple solution to find the location of previously tagged files.

To add it to the project, drag and drop both all files (except for excel file) from the **Files** folder to the VBA Project window (Developer -> Visual Basic), then add a reference to the **Microsoft Scripting Runtime** (Tools -> References). Then, add a Button from Form Controls to any sheet and assign it the **InitializeForm** function.

To mark the location, in given folder you have to put the file with the .projData extension (or, with the appropriate code modification, any extension), and then in a text editor modify the file according to the following template:

>Project name
>
>*Project name*
>
>Tags
>
>*Tags entered one under the other*
>
>Additional content
>
>*Text specifying information about the location of the file/s; currently not implemented*

To open location of the item in the list window double click on the item.

Example showing the functionality has been attached to the repository. Example has slightly modified code that makes excel application hidden when running.

For demonstration purposes, disk scanning is limited to folders located in the same folder as filesearcher. 

