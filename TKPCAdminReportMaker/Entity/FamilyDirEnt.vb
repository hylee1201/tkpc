Namespace TKPC.Entity
    Public Class FamilyDirEnt
        Private mPersonId As String
        Private mHeadName As String
        Private mKoreanNames As String
        Private mEngName1 As String
        Private mEngName2 As String
        Private mAddress As String
        Private mCity As String
        Private mPhone1 As String
        Private mPhone2 As String
        Private mEmail As String

        Public Property personId() As String
            Get
                Return mPersonId
            End Get
            Set(ByVal a As String)
                mPersonId = a
            End Set
        End Property

        Public Property headName() As String
            Get
                Return mHeadName
            End Get
            Set(ByVal a As String)
                mHeadName = a
            End Set
        End Property

        Public Property koreanNames() As String
            Get
                Return mKoreanNames
            End Get
            Set(ByVal a As String)
                mKoreanNames = a
            End Set
        End Property

        Public Property engName1() As String
            Get
                Return mEngName1
            End Get
            Set(ByVal a As String)
                mEngName1 = a
            End Set
        End Property

        Public Property engName2() As String
            Get
                Return mEngName2
            End Get
            Set(ByVal a As String)
                mEngName2 = a
            End Set
        End Property

        Public Property address() As String
            Get
                Return mAddress
            End Get
            Set(ByVal a As String)
                mAddress = a
            End Set
        End Property

        Public Property city() As String
            Get
                Return mCity
            End Get
            Set(ByVal a As String)
                mCity = a
            End Set
        End Property

        Public Property phone1() As String
            Get
                Return mPhone1
            End Get
            Set(ByVal a As String)
                mPhone1 = a
            End Set
        End Property

        Public Property phone2() As String
            Get
                Return mPhone2
            End Get
            Set(ByVal a As String)
                mPhone2 = a
            End Set
        End Property

        Public Property email() As String
            Get
                Return mEmail
            End Get
            Set(ByVal a As String)
                mEmail = a
            End Set
        End Property
    End Class
End Namespace
