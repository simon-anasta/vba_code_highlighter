# vba_code_highlighter

In some scenarios, non-coders need to be able to review code to determine whether it contains any sensitive information (for example, comments relating to individual records that might represent a privacy breach).

Expecting non-coders to read the code within an IDE, or to use code tools to check code, has been ineffective. But Excel is widerly available, VBA provides the option to run code, and keyboard shortcuts can be bound to macros. All of this makes Excel a plausible approach to support non-coders reviewing code files.

# method

The excel sheet provides a white list and a black list. When the macro is run it prompts the user to open a file, this is copied into the Excel sheet. Lines of code that contain text from the black list that do not appear on the white list are highlighted to guide the checker's attention to key areas of the document.

For R, SQL and SAS files we also highlight lines of code that contains comments. Such lines are more likely to contain sensitive information than non-comments.

