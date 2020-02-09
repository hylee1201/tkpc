Namespace TKPC.Entity

    Public Class CommentEnt
        Private mPersonId As String
        Private mAuthorName As String
        Private mPersonName As String
        Private mBody As String

        Public Property personId() As String
            Get
                Return mPersonId
            End Get
            Set(ByVal a As String)
                mPersonId = a
            End Set
        End Property

        Public Property authorName() As String
            Get
                Return mAuthorName
            End Get
            Set(ByVal a As String)
                mAuthorName = a
            End Set
        End Property

        Public Property personName() As String
            Get
                Return mPersonName
            End Get
            Set(ByVal a As String)
                mPersonName = a
            End Set
        End Property

        Public Property body() As String
            Get
                Return mBody
            End Get
            Set(ByVal a As String)
                mBody = a
            End Set
        End Property

    End Class

End Namespace