'----------------------------------
' Completion Handler 
'----------------------------------
Imports System.Management

Namespace completionHandler

    Public Class MyHandler
        Private m_isComplete As Boolean = False
        Private m_returnObject As ManagementBaseObject

        ''' <summary>
        ''' Trigger Done event when InvokeMethod is complete
        ''' </summary>
        Public Sub Done(ByVal sender As Object, ByVal e As ObjectReadyEventArgs)
            m_isComplete = True
            m_returnObject = e.NewObject
        End Sub


        ''' <summary>
        ''' Get property IsComplete
        ''' </summary>
        Public ReadOnly Property IsComplete() As Boolean
            Get
                Return m_isComplete
            End Get
        End Property

        ''' <summary>
        ''' Property allows accessing the result 
        ''' object in the main function
        ''' </summary>
        Public ReadOnly Property ReturnObject() As ManagementBaseObject
            Get
                Return m_returnObject
            End Get
        End Property

    End Class
End Namespace
