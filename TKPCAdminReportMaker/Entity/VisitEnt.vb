Namespace TKPC.Entity
    Public Class VisitEnt
        Private mVisitDate As String
        Private mVisitedKoreanName As String
        Private mNote As String

        Public Property visitDate() As String
            Get
                Return mVisitDate
            End Get
            Set(ByVal a As String)
                mVisitDate = a
            End Set
        End Property

        Public Property visitedKoreanName() As String
            Get
                Return mVisitedKoreanName
            End Get
            Set(ByVal a As String)
                mVisitedKoreanName = a
            End Set
        End Property

        Public Property note() As String
            Get
                Return mNote
            End Get
            Set(ByVal a As String)
                mNote = a
            End Set
        End Property
    End Class
End Namespace