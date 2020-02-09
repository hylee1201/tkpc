Namespace TKPC.Entity
    Public Class PersonEnt
        Private mPersonId As String
        Private mKoreanName As String
        Private mLastName As String
        Private mFirstName As String
        Private mGender As String
        Private mDOB As String
        Private mAge As String
        Private mCell As String
        Private mHome As String
        Private mWork As String
        Private mEmail As String
        Private mSpouseIds As String
        Private mStatus As String
        Private mAddressType As String
        Private mAddress As String

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

        Public Property LastName() As String
            Get
                Return mLastName
            End Get
            Set(ByVal a As String)
                mLastName = a
            End Set
        End Property

        Public Property firstName() As String
            Get
                Return mFirstName
            End Get
            Set(ByVal a As String)
                mFirstName = a
            End Set
        End Property

        Public Property gender() As String
            Get
                Return mGender
            End Get
            Set(ByVal a As String)
                mGender = a
            End Set
        End Property

        Public Property dob() As String
            Get
                Return mDOB
            End Get
            Set(ByVal a As String)
                mDOB = a
            End Set
        End Property

        Public Property age() As String
            Get
                Return mAge
            End Get
            Set(ByVal a As String)
                mAge = a
            End Set
        End Property

        Public Property cell() As String
            Get
                Return mCell
            End Get
            Set(ByVal a As String)
                mCell = a
            End Set
        End Property

        Public Property home() As String
            Get
                Return mHome
            End Get
            Set(ByVal a As String)
                mHome = a
            End Set
        End Property

        Public Property work() As String
            Get
                Return mWork
            End Get
            Set(ByVal a As String)
                mWork = a
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

        Public Property spouseIds() As String
            Get
                Return mSpouseIds
            End Get
            Set(ByVal a As String)
                mSpouseIds = a
            End Set
        End Property

        Public Property status() As String
            Get
                Return mStatus
            End Get
            Set(ByVal a As String)
                mStatus = a
            End Set
        End Property

        Public Property addressType() As String
            Get
                Return mAddressType
            End Get
            Set(ByVal a As String)
                mAddressType = a
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
    End Class
End Namespace
