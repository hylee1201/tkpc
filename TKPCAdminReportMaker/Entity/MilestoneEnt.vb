Namespace TKPC.Entity
    Public Class MilestoneEnt
        Private mPersonId As String
        Private mKoreanName As String
        Private mType As String
        Private mName As String
        Private mEffectiveDate As String

        Public Property effectiveDate() As String
            Get
                Return mEffectiveDate
            End Get
            Set(ByVal a As String)
                mEffectiveDate = a
            End Set
        End Property

        Public Property personId() As String
            Get
                Return mPersonId
            End Get
            Set(ByVal a As String)
                mPersonId = a
            End Set
        End Property

        Public Property koreanName() As String
            Get
                Return mKoreanName
            End Get
            Set(ByVal a As String)
                mKoreanName = a
            End Set
        End Property

        Public Property type() As String
            Get
                Return mType
            End Get
            Set(ByVal a As String)
                mType = a
            End Set
        End Property

        Public Property name() As String
            Get
                Return mName
            End Get
            Set(ByVal a As String)
                mName = a
            End Set
        End Property
    End Class
End Namespace