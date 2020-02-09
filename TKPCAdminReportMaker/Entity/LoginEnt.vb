Namespace TKPC.Entity
    Public Class LoginEnt
        Private mId As String 'id of admin user
        Private mUserId As String
        Private mPassword As String
        Private mRole As String
        Private mUserName As String

        Public Property id() As String
            Get
                Return mId
            End Get
            Set(ByVal a As String)
                mId = a
            End Set
        End Property

        Public Property userId() As String
            Get
                Return mUserId
            End Get
            Set(ByVal a As String)
                mUserId = a
            End Set
        End Property

        Public Property password() As String
            Get
                Return mPassword
            End Get
            Set(ByVal a As String)
                mPassword = a
            End Set
        End Property

        Public Property role() As String
            Get
                Return mRole
            End Get
            Set(ByVal a As String)
                mRole = a
            End Set
        End Property

        Public Property userName() As String
            Get
                Return mUserName
            End Get
            Set(ByVal a As String)
                mUserName = a
            End Set
        End Property
    End Class
End Namespace
