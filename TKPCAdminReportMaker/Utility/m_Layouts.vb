Module m_Layouts
    Public Sub FindPeople()
        With mdiTKPC.ugFind
            .DisplayLayout.Bands(0).Columns(0).Hidden = True
            .DisplayLayout.Bands(0).Columns(1).Width = 70 'Korean Name
            .DisplayLayout.Bands(0).Columns(2).Width = 70 'Last Name
            .DisplayLayout.Bands(0).Columns(3).Width = 70 'First Name
            .DisplayLayout.Bands(0).Columns(4).Width = 30 'Gender
            .DisplayLayout.Bands(0).Columns(5).Width = 60 'DOB
            .DisplayLayout.Bands(0).Columns(6).Width = 40 'Age
            .DisplayLayout.Bands(0).Columns(7).Width = 90 'Cell Phone
            .DisplayLayout.Bands(0).Columns(8).Width = 150 'Home address
            .DisplayLayout.Bands(0).Columns(9).Width = 90 'spouses
            .DisplayLayout.Bands(0).Columns(9).Hidden = True
            .DisplayLayout.Override.RowAppearance.TextVAlign = Infragistics.Win.VAlign.Middle
        End With
    End Sub
End Module
