Option Compare Text
Option Explicit
Option Private Module

Public Function AccUnitTestClassFactory_zTST_ErrorState() As Object
   Set AccUnitTestClassFactory_zTST_ErrorState = New zTST_ErrorState
End Function

Public Function AccUnitTestClassFactory_zTST_TransactionalExecutorCommonBehaviour() As Object
   Set AccUnitTestClassFactory_zTST_TransactionalExecutorCommonBehaviour = New zTST_TransactionalExecutorCommonBehaviour
End Function

Public Function AccUnitTestClassFactory_zTST_TransactionalExecutorWithError() As Object
   Set AccUnitTestClassFactory_zTST_TransactionalExecutorWithError = New zTST_TransactionalExecutorWithError
End Function

Public Function AccUnitTestClassFactory_zTST_TransactionalExecutorWithoutError() As Object
   Set AccUnitTestClassFactory_zTST_TransactionalExecutorWithoutError = New zTST_TransactionalExecutorWithoutError
End Function