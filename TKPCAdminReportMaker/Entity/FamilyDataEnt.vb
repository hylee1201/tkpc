Namespace TKPC.Entity
    Public Class FamilyDataEnt
        Private mId As String
        Private mKoreanName As String
        Private mFamilyDataTable As DataTable

        Public Property id() As String
            Get
                Return mId
            End Get
            Set(ByVal a As String)
                mId = a
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

        Public Property familyDataTable() As DataTable
            Get
                Return mFamilyDataTable
            End Get
            Set(ByVal a As DataTable)
                mFamilyDataTable = a
            End Set
        End Property

    End Class
End Namespace