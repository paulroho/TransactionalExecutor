# TransactionalExecutor
A class for easier execution of code in the context of a transaction without the boilerplate code for VBA/DAO.

## What it does
* Spans a DAO-transaction
* Calls a method
* Commits the transaction if no error occured
* Rolls the transaction back if an error occured
 
For better user interacation several hooks exist in the process. 
 
## Basic Usage

### Preparation
* Copy the class modules `AT_TransactionalExecutor` and `AT_ErrorState` (folder [TransactionalExecutorCreator.accdb.Content/Code/AT](TransactionalExecutorCreator.accdb.Content/Code/AT)) to your VBA project.
* Make sure a reference to the DAO library is set

### Minimal Scenario
The following code snipped shows this in a minimal setup. Note that this code only works in a class module (can be a form's code behind).
````vbnet
Private WithEvents MultistepOperation As AT_TransactionalExecutor
' ...

Public Sub DoTheOperation()
    Set MultistepOperation = New At_TransactionalExecutor
    MultistepOperation.Execute()
    Set MultistepOperation = Nothing
End Sub

' ...

Private Sub MultistepOperation_Execute(ByVal ErrorState As AT_ErrorState)
On Error Goto Err_

    ' YOUR OPERATIONS GO HERE
    ' They are executed in the context of a DAO transaction

    Exit Sub
Err_:
    ErrorState.SetError Err
End Sub
````

### What above code does
1. Declares a module level variable of `TransactionalExecutor` with the keyword `WithEvents`
2. Calls your code that should be run within the transaction in an event handler for the event `Execute`.  
   Include basic error handling (see below). 
3. Assigns an instance of `TransactionalExecutor` to the module level variable
4. Calls the `Execute` method on that instance   

## Features
### Hooks
`TransactionExecutor` provides hooks in the form of events.

The following events can be used. Typically, the `Execute` event is implemented first and other events are used for improved user feedback.
#### Execution Hooks
* [`Public Event BeforeExecute (ByRef Cancel As Boolean)`](#beforeexecute)
* [`Public Event Execute(ByVal ErrorState As AT_ErrorState)`](#execute)
* [`Public Event AfterExecute()`](#afterexecute)

#### Commit Hooks 
* [`Public Event BeforeCommit(ByRef Cancel As Boolean)`](#beforecommit)
* [`Public Event AfterCommit()`](#aftercommit)

#### Rollback Hook 
* [`Public Event AfterRollback(ByVal ErrorState As AT_ErrorState)`](#afterrollback)

#### Overview
![Overview](https://raw.githubusercontent.com/paulroho/TransactionalExecutor/master/Documentation/flowchart.mmd.png)

#### BeforeExecute
Is the first event raised by a call to `Execute()`. Can be used to cancel the operation.

##### Signature
````vbnet
Public Event BeforeExecute(ByRef Cancel As Boolean)
````

##### Typical use cases

* Check preconditions and cancel the operation if they are not met
* Show confirmation message and let the user decide whether to run the operation

##### Sample
````vbnet
Private Sub MultistepOperation_BeforeExecute(ByRef Cancel As Boolean)
    Cancel = (vbCancel = MsgBox("The complicated operation will be started.", vbOkCancel))
End Sub
````

#### Execute
Place the code to be executed within the transaction here. Make sure to follow the error handling pattern below.

##### Signature
````vbnet
Public Event Execute(ByVal ErrorState As AT_ErrorState)
````

##### Error handling pattern
Since an error raised from an event handler cannot be caught, a simple error handling pattern has to be implemented. The important part here is to call `ErrorState.SetError Err` in case of an error:
````vbnet
Private Sub MultistepOperation_Execute(ByVal ErrorState As AT_ErrorState)
On Error Goto Err_

    ' YOUR OPERATIONS GO HERE
    ' They are executed in the context of a DAO transaction

    Exit Sub
Err_:
    ErrorState.SetError Err
End Sub
````

If `ErrorState.SetError` has not been called, the transaction gets committed. Otherwise a rollback is performed.


#### AfterExecute
This event is fired immediately after the even `Execute` regardless of the outcome. Any commit or rollback happens afterwards.

##### Signature
````vbnet
Public Event AfterExecute()
````

##### Typical Use Case
* Hide any busy indications such as hourglass or progress bars

##### Sample
````vbnet
Private Sub MultistepOperation_AfterExecute()
    Screen.MousePointer = 0
End Sub
````


#### BeforeCommit
Is raised after the operation completed without error but before the transaction is commited. This operation can be canceled. The transaction will be cancelled afterwards if the event is not cancelled. If it is cancelled, the operation will be rolled back (without firing the event `AfterRollback`).

##### Signature
````vbnet
Public Event BeforeCommit(ByRef Cancel As Boolean)
````

##### Typical Use Case
Give the user a chance to avoid committing the operation result to the database. Information about the execution that was eventually collected during execution can be provided.

##### Sample
````vbnet
Private Sub MultistepOperation_BeforeCommit(ByRef Cancel As Boolean)
    ' Assume m_RecordsAffected has been set in the handler for the Execute event.
    Cancel = (vbCancel = MsgBox("m_RecordsAffected & " records are going to be updated.", vbOKCancel))
End Sub
````

#### AfterCommit
Is called after the transaction has been successfully committed.

##### Signature
````vbnet
Public Event AfterCommit()
````

##### Typical Use Case
* Log the (number of) affected records
* Refresh UI showing data affected by the operation

##### Sample
````vbnet
Private Sub MultistepOperation_AfterCommit()
    ' Assuming this is called in a form's code behind class
    Me.Requery
End Sub
````


#### AfterRollback
Is called after the transaction has been rolled back due to an error. The `ErrorState` set by the `Execute` event handler is provided as a parameter. 

##### Signature
````vbnet
Public Event AfterRollback(ByVal ErrorState As AT_ErrorState)
````

##### Typical Use Case
* Inform the user about the the error that has occured.

##### Sample
````vbnet
Private Sub MultistepOperation_AfterRollback(ByVal ErrorState As AT_ErrorState)
    MsgBox "An error has occured:" & vbNewLine & _
           ErrorState.Description
End Sub
````
##### Remark
The `AfterRollback` event is not fired if the rollback was done because the `BeforeCommit` event has been cancelled.