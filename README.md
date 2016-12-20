# VBScript-REPL
A REPL (Read, Eval, Print, Loop) for VBScript/WSH

Many languages come with their own REPL or interactive shell for reading and evaluating user input without the need to write and execute a source/script file. This project aims to create such a feature for the VBScript language.

##Execution and Evaluation##

Each line of input is both *executed* and *evaluated*. VBScript's **ExecuteGlobal()** function is used first, which executes each statement in the global namespace. Any errors are echoed to the user.
```
>>> strMessage = "Hello, world!"
>>> intCount = 17
>>> blnPassed = False
>>> dblUnitCost = 6005.9752
>>> Set fso = CreateObject("Scripting.FileSystemObject")
>>> WScript.Echo "The count is", intCount
The count is 17
>>> WScript.Print "Hey"
Object doesn't support this property or method
```

Each line of input is then evaluated using VBScript's **Eval()** function. This allows for variable inspection and evaulation of literal expressions. The output is formatted based on the variable type. For example, strings are surrounded in quotes, objects display their class/type, and currency and date values are formatted using the built-in **FormatCurrency()** and **FormatDateTime()** functions, respectively.

```
>>> strMessage                     ' Just enter the name of a variable to see its value 
"Hello, world!"
>>> WScript                        ' Works for objects, too
Object (Windows Script Host)
>>> a = Array("Bob", 37, True)     ' And arrays
>>> a
("Bob", 37, True)
>>> vbWednesday                    ' And constants
4
>>> j
Name is undefined                  ' Attempting to use an undefined var will display an error
>>> c = CCur(dblAmount)
>>> c
$6,005.98                          ' Values are formatted based on their type
>>> 17 - 5 / 2                     ' Literal expressions are evaluated
14.5
>>> LCase(WScript.Name)
"windows script host"
>>> fso.FolderExists("C:\Windows")
True
```

##Assignment or Equality?##
Ideally, the REPL would be able to determine the user's intent and then properly forward all statements to **ExecuteGlobal()** and any expressions to **Eval()**. However, VBScript makes this difficult because it uses `=` for both assignment and for tests of equality. So both functions are called and any errors are captured. If **ExecuteGlobal()** succeeds, it is assumed to be a successful assignment. Knowing this, you can fool the REPL into treating it like an expression by surrounding it in parentheses, which causes **ExecuteGloba()** to fail.

```
>>> i = 1                          ' Assignment
>>> i
1
>>> i = 1                          ' Re-assignment, not comparison
>>> i
1
>>> (i = 1)                        ' Invalid statement. Treated as expression.
True
>>> CBool(i = 2)                   ' Another way to force evaluation
False
>>> 1 = 1                          ' Of course, ExecuteGlobal() fails here, too, so this gets evaluated
True
>>> 0 = False
True
```

##Block Definitions##
Blocks can also be defined within the REPL. This allows for the use of conditionals (`If`, `Select`, `Do`, `While`, `For`), contexts (`With`), classes (`Class`), and functions (`Sub`, `Function`). Ellipses will be displayed instead of the standard REPL prompt to indicate that you're in "block definition mode". Enter an empty line of input to exit. The block will then be executed and any errors reported. (I toyed with using the MSScriptControl to pre-execute the block so that I could identify the line and column where errors occur, but the script control is "hostless" and would fail when attempting to resolve anything `WScript`-related).

```
>>> Sub Increment(i)
...     i = i + 1
... End Sub
...
>>> j = 1
>>> Increment j                    ' Invoke the subroutine
>>> j
2
>>> Call Increment(j)              ' Alternate calling syntax
>>> j
3
>>> Increment(j)                   ' Perhaps a VBScript bug? This is invalid syntax but ExecuteGlobal()
>>> j                              ' throws no error. Unfortunately, it also doesn't call the function.
3
>>> Function Increment(i)
...     Increment = i + 1
... End Function
...
>>> j = 1
>>> j = Increment(j)
>>> j
2
>>> For j = 1 To 3
...     WScript.Echo j
... Next
...
1
2
3
```

##Importing Scripts##
You can also import other scripts with the `Import` keyword (this isn't a part of VBScript, of course, but I think we all wish it was). All names defined within the imported script will be added to the global namespace. `Import` checks the current folder for the file.

```
>>> Import foo.vbs                ' Imagine foo.vbs contains a class definition of Foo
>>> Set f = New Foo()             ' Use the imported class definition
```

In addition, if the file `init.vbs` exists in the current folder, it will be imported automatically when the REPL is initialized.

##Exiting the REPL##
To exit the REPL, simply type `exit`.

```
>>> exit
```

##To Do##
Still have some basic stuff to take care of, like line continuation. It's not something I use often and, therefore, hasn't been missed yet. But it's a planned addition.

- [ ] Line continuation

##Licensing##
This project is licensed under the MIT License. In short, you're welcome to do anything you want with this code as long as you provide attribution and don't hold me liable.

I hope you find it useful.
