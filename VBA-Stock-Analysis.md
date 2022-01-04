This dataset presents  stock data for analysis to make the best investment decisions

## **History of VBA**

- **BASIC** (short for Beginner's All-purpose Symbolic Instruction Code) was a programming language invented in the 1960s to help teach programming concepts. 
- It soon gained traction and started to be used as a full-fledged programming language.
-  In the 1990s, Microsoft created a version of BASIC with a visual form builder so that graphical desktop applications could be built, and Visual Basic was born! It lives on today in VBA and VB.NET.

**Key concepts**

- In developer parlance, automated tasks are called **macros**. Originally, macros were created by "recording" a task that you performed in Excel, and VBA would automatically generate code to repeat the task.
- The macro could be run repeatedly to quickly perform the task over and over again.
- Almost all VBA code is written to create macros, which are sometimes called **subroutines**.
- However, VBA is powerful enough to connect to the internet and run applications in the operating system, which means it can be abused to write malicious code, like viruses and trojan horses.
- This is why it is disabled and you need to enable VBA by adding the Developer 

**Build a Subroutine**

- A **subroutine** is a key building block of a VBA macro. Subroutines are a collection of steps or instructions. 

- A subroutine is given a name so that the subroutine can be **called**, or run. - Basically a function

- Write a subroutine called MacroCheck: 

- ```visual-basic
  Sub MacroCheck()
  
  End Sub
  ```

Here, `Sub` is a statement that tells VBA to create a subroutine. `MacroCheck` is what we're telling VBA this subroutine is called.

- We need the "End Sub" to signify the end. 

**Create a Variable**

- The keyword to create a variable in VBA is `Dim`, which is short for "dimension." 

- We have to tell VBA the name of the variable and what kind of data type it will store. 
- Because the message will display text, we need to use `String` as the data type. So the full line of code is:

```visual-basic
Dim testMessage As String
testMessage = "Hello World!"

#Full code...
Sub MacroCheck()

    Dim testMessage As String

    testMessage = "Hello World!"

    MsgBox (testMessage)

End Sub

```

- The keyword to create a message box is `MsgBox`.
- `MsgBox` takes in an **argument**, which, in our case, is what we want our pop-up box to display.