Imports Npgsql
Imports System.Text
Imports System.Collections.Generic
Imports log4net

Namespace TKPC.DAL
    Public Class PersonDAL

        Private Shared Instance As PersonDAL

        Protected Sub New()
        End Sub

        'To make this DAL singleton to use memory efficiently
        Public Shared Function getInstance() As PersonDAL
            ' initialize if not already done
            If Instance Is Nothing Then
                Instance = New PersonDAL
            End If
            ' return the initialized instance of the Singleton Class
            Return Instance
        End Function 'Instance

        Public Function findPeopleList(ByVal flag As Integer, ByVal name As String, ByVal statusFlag As Integer, ByVal filterFlag As Integer, ByVal namesSQL As String, ByVal titleYear As String, ByVal ageFrom As String, ByVal ageTo As String) As DataSet
            Dim sql As String = String.Empty
            findPeopleList = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder
            Dim temp As String = ""

            Logger.LogInfo("flag=[" + CType(flag, String) + "], name=[" + name + "], statusFlag=[" + CType(statusFlag, String) + "]")

            If String.IsNullOrEmpty(name) = False Then
                name = name.Trim
            End If

            If String.IsNullOrEmpty(ageFrom) = False Then
                ageFrom = ageFrom.Trim
                temp = ageFrom
            End If

            If String.IsNullOrEmpty(ageTo) = False Then
                ageTo = ageTo.Trim
            End If

            If flag = 11 Then
                If String.IsNullOrEmpty(ageFrom) = False And String.IsNullOrEmpty(ageTo) = False Then
                    ageFrom = ageTo
                    ageTo = temp
                End If
            End If

            Dim year As String = "z.person_id, "
            If flag = 1 Then
                year = "z.person_id, z.year, "
            End If

            Dim selectStr As String = "select " + year + " z.korean_name, z.kn_cnt, z.title, z.title_date, " +
                                      "  (select string_agg(cast(date_part('year', effective_date) as varchar), ',' ORDER BY date_part('year', effective_date) desc) " +
                                      "   from public.milestones m " +
                                      "   where type = 'TitleMilestone' and name = z.title " +
                                      "   and person_id = z.person_id) title_years, " +
                                      "   (select string_agg(cast(year as varchar), ',' ORDER BY year desc) " +
                                      "   from public.offering_envelopes e where person_id = z.person_id) offering_years, " +
                                      "       z.last_name, z.first_name, z.gender, z.dob, z.dob_year, z.age,   " +
                                      "       	   CASE WHEN cast(z.person_id as varchar) = z.spouses THEN 'Y' " +
                                      "            ELSE 'N' " +
                                      "            END single," +
                                      "       case " + vbCrLf +
                                      "       WHEN (z.age <= 0) THEN '0세 이하' " + vbCrLf +
                                      "       WHEN (z.age > 0) AND (z.age <= 10) THEN '1-10세' " + vbCrLf +
                                      "       WHEN (z.age > 10) AND (z.age <= 20) THEN '11-20세' " + vbCrLf +
                                      "       WHEN (z.age > 20) AND (z.age <= 30) THEN '21-30세' " + vbCrLf +
                                      "       WHEN (z.age > 30) AND (z.age <= 40) THEN '31-40세' " + vbCrLf +
                                      "       WHEN (z.age > 40) AND (z.age <= 50) THEN '41-50세' " + vbCrLf +
                                      "       WHEN (z.age > 50) AND (z.age <= 60) THEN '51-60세' " + vbCrLf +
                                      "       WHEN (z.age > 60) AND (z.age <= 70) THEN '61-70세' " + vbCrLf +
                                      "       WHEN (z.age > 70) AND (z.age <= 80) THEN '71-80세' " + vbCrLf +
                                      "       WHEN (z.age > 80) AND (z.age <= 90) THEN '81-90세' " + vbCrLf +
                                      "       WHEN (z.age > 90) AND (z.age <= 100) THEN '91-100세' " + vbCrLf +
                                      "       WHEN (z.age > 100) AND (z.age <= 110) THEN '101-110세' " + vbCrLf +
                                      "       ELSE '110세 이후' " + vbCrLf +
                                      "       end age_range, z.email, " + vbCrLf +
                                      "      z.cell_phone, z.home_phone, z.work_phone, z.address, z.addr_cnt, " +
                                      "      CASE WHEN z.status = 'church_member' THEN '등록' WHEN z.status = 'left_church' THEN '이적' WHEN z.status = 'new_family' THEN '새교우' ELSE '' END status, " +
                                      "      z.status_date, z.deceased_on, z.spouses, z.photo_file_name "

            With sb
                .Append(selectStr + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("    Select p.id As person_id, korean_name, " + vbCrLf)
                .Append("           COUNT(*) OVER (PARTITION BY korean_name) kn_cnt, " + vbCrLf)
                .Append("           INITCAP(last_name) last_name, INITCAP(first_name) first_name, gender, date_of_birth dob, to_char(date_of_birth, 'yyyy') dob_year,  " + vbCrLf)
                .Append("           date_part('year', age(date_of_birth)) age, cell_phone, home_phone, work_phone, " + vbCrLf)
                .Append("           deceased_on,  " + vbCrLf)

                If flag = 1 Then
                    .Append("       e.year, e.offering_number, ")
                End If

                .Append("           (select x.address  " + vbCrLf)
                .Append("            from (  " + vbCrLf)
                .Append("               select INITCAP(a.street) || ' ' || INITCAP(a.city) || ', ' || INITCAP(CASE WHEN lower(a.province)='ontario' THEN 'Ont.' ELSE a.province END) || ' ' || a.postal_code as address, ROW_NUMBER() OVER (PARTITION BY a.person_id ORDER BY created_at) row_num  " + vbCrLf)
                .Append("               from public.addresses a where p.id = a.person_id and a.address_type = 'home'  " + vbCrLf)
                .Append("             ) x  " + vbCrLf)
                .Append("             where x.row_num = 1  " + vbCrLf)
                .Append("           ) address, " + vbCrLf)
                .Append("		    ( " + vbCrLf)
                .Append("               select COUNT(*) " + vbCrLf)
                .Append("               from public.addresses a where p.id = a.person_id " + vbCrLf)
                .Append("           ) addr_cnt, " + vbCrLf)
                .Append("           (select x.name " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'TitleMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) title, " + vbCrLf)
                .Append("           (select x.effective_date " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'TitleMilestone' " + vbCrLf)

                If flag = 3 And String.IsNullOrEmpty(titleYear) = False And titleYear <> "모든년도" Then '직분명
                    .Append(" and date_part('year', effective_date) = " + CType(titleYear, String) + " " + vbCrLf)
                End If

                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) title_date, " + vbCrLf)
                .Append("           (select x.name " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'MembershipMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) status, " + vbCrLf)
                .Append("           (select x.effective_date " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'MembershipMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("           where x.row_num = 1 " + vbCrLf)
                .Append("           ) status_date, " + vbCrLf)
                .Append("          (select x.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'baptism' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) baptism, " + vbCrLf)
                .Append("          (select x.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'baptism' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) baptism_date, " + vbCrLf)

                .Append("          (select x.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'infant_baptism' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) infant_baptism, " + vbCrLf)
                .Append("          (select x.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'infant_baptism' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) infant_baptism_date, " + vbCrLf)

                .Append("          (select x.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'confirmation' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) confirmation, " + vbCrLf)
                .Append("          (select x.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'confirmation' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) confirmation_date, " + vbCrLf)
                .Append("     (select CASE WHEN string_agg(cast(y.f_id as varchar), ',' ORDER BY y.f_id) IS NULL THEN cast(p.id as varchar) " + vbCrLf)
                .Append("             ELSE string_agg(cast(y.f_id As varchar), ',' ORDER BY y.f_id) END " + vbCrLf)
                .Append("      from (  " + vbCrLf)
                .Append("           select distinct r.person_id f_id, r.relationship_type as rtype  " + vbCrLf)
                .Append("           from Public.relationships r   " + vbCrLf)
                .Append("           where p.id = r.related_person_id   " + vbCrLf)
                .Append("           union all   " + vbCrLf)
                .Append("           select distinct r.related_person_id f_id, r.relationship_type as rtype  " + vbCrLf)
                .Append("           from Public.relationships r   " + vbCrLf)
                .Append("           where p.id = r.person_id   " + vbCrLf)
                .Append("           union all   " + vbCrLf)
                .Append("           select distinct r.person_id f_id, r.relationship_type as rtype  " + vbCrLf)
                .Append("           from Public.relationships r   " + vbCrLf)
                .Append("           where p.id = r.person_id   " + vbCrLf)
                .Append("           union all   " + vbCrLf)
                .Append("           select distinct r.related_person_id f_id, r.relationship_type as rtype  " + vbCrLf)
                .Append("           from Public.relationships r   " + vbCrLf)
                .Append("           where p.id = r.related_person_id   " + vbCrLf)
                .Append("       ) y    " + vbCrLf)
                .Append("       where y.rtype = 'spouse') spouses, photo_file_name, email " + vbCrLf)

                If flag = 1 Then
                    .Append("   from public.people p, public.offering_envelopes e  " + vbCrLf)
                Else
                    .Append("   from public.people p  " + vbCrLf)
                End If

                If flag = 0 Then
                    If String.IsNullOrEmpty(name) = False Then
                        If name = "샘플" Then
                            .Append("where p.korean_name in ('주문자','김경희B','김경희D','김진희','백수정','오화연','이승범','김동규','장성훈','손희전','강영태','선희주','이규대','채성민','배원욱','양원칠','나형주','배기욱','신학경','손명수','송남성','손미라','유희문','김동균','김성겸','원석희','유명인','김해천','김형근','김춘호','고우길','원태웅','오브람','김상우','신한준','이해영','류세진','강영숙','나형주','고우길','김미령','심홍섭','우석균','원태웅','강태식','고흥태','강영희','고일용','이민선') " + vbCrLf)
                            'ElseIf name = "교샘플" Then
                            '    .Append("where p.korean_name in ('강영숙','나형주','고우길','김미령','심홍섭','우석균','원태웅','오브람', '심홍섭', '신학경', '신지숙') " + vbCrLf)
                            'ElseIf name = "장로" Then
                            '    .Append("where p.korean_name in ('손명수','장성훈','김동규','김성겸','나형주','배덕출','양원칠','이규대','이승범','이용식','채성민','김춘호','김성겸') " + vbCrLf)
                        ElseIf name = "샘플2" Then
                            .Append("where p.korean_name in ('강태식','고우길','오화연','주문자','신학경','김진희','백수정') " + vbCrLf)

                        Else
                            If filterFlag = 3 Then
                                .Append("where (lower(p.korean_name)) = @name or (lower(last_name) = @name) or (lower(first_name) = @name) " + vbCrLf)
                            Else
                                .Append("where lower(p.korean_name) like @name or lower(last_name) like @name or lower(first_name) like @name " + vbCrLf)
                            End If
                        End If
                    End If
                ElseIf flag = 1 Then
                    .Append("where p.id = e.person_id " + vbCrLf)
                    If String.IsNullOrEmpty(name) = False Then
                        .Append("and e.offering_number = @name " + vbCrLf)
                    End If
                ElseIf flag = 5 Then 'From excel file
                    If String.IsNullOrEmpty(name) = True Or
                      (String.IsNullOrEmpty(name) = False And name.IndexOf(".csv") > 0) Or
                      (String.IsNullOrEmpty(name) = False And name.IndexOf(".txt") > 0) Then
                        .Append("where p.korean_name in (" + namesSQL + ") " + vbCrLf)
                    End If
                End If

                .Append(") z " + vbCrLf)

                '0 - 전체 데이터
                '1 - 등록교인만
                '2 - 등록교인 + 새교우
                '3 - 등록교인 + 새교우 + 이적교인
                '4 - 등록교인 + 새교우 + 고인
                '5 - 등록교인 + 이적교인
                '6 - 등록교인 + 이적교인 + 고인
                '7 - 등록교인 + 고인
                '8 - 새교우만
                '9 - 새교우 + 이적교인
                '10 - 새교우 + 이적교인 + 고인
                '11 - 새교우 + 고인
                '12 - 이적교인만
                '13 - 이적교인 + 고인
                '14 - 고인만
                '15 - 모두 아님

                If statusFlag = 0 Then '모두
                    .Append(" where 1=1 " + vbCrLf)
                ElseIf statusFlag = 1 Then '등록교인만
                    .Append(" where z.status = 'church_member' and z.deceased_on is null " + vbCrLf)
                ElseIf statusFlag = 2 Then '등록교인 + 새교우
                    .Append(" where ((z.status = 'church_member' and z.deceased_on is null) or (z.status = 'new_family' and  z.deceased_on is null)) " + vbCrLf)
                ElseIf statusFlag = 3 Then '등록교인 + 새교우 + 이적교인
                    .Append(" where ((z.status = 'church_member' and z.deceased_on is null) or (z.status = 'new_family' and  z.deceased_on is null) or (z.status = 'left_church' and  z.deceased_on is null))  " + vbCrLf)
                ElseIf statusFlag = 4 Then '등록교인 + 새교우 + 고인
                    .Append(" where ((z.status = 'church_member') or (z.status = 'left_church') or (z.deceased_on is not null)) " + vbCrLf)
                ElseIf statusFlag = 5 Then '등록교인 + 이적교인
                    .Append(" where ((z.status = 'church_member' and z.deceased_on is null) or (z.status = 'left_church' and  z.deceased_on is null))  " + vbCrLf)
                ElseIf statusFlag = 6 Then '등록교인 + 이적교인 + 고인
                    .Append(" where ((z.status = 'church_member') or (z.status = 'left_church')  or (z.deceased_on is not null)) " + vbCrLf)
                ElseIf statusFlag = 7 Then '등록교인 + 고인
                    .Append(" where ((z.status = 'church_member') or (z.deceased_on is not null)) " + vbCrLf)
                ElseIf statusFlag = 8 Then '새교우
                    .Append(" where ((z.status = 'new_family' and z.deceased_on is null)) " + vbCrLf)
                ElseIf statusFlag = 9 Then '새교우 + 이적교인
                    .Append(" where ((z.status = 'new_family' and z.deceased_on is null) or (z.status = 'left_church' and  z.deceased_on is null))  " + vbCrLf)
                ElseIf statusFlag = 10 Then '새교우 + 이적교인 + 고인
                    .Append(" where ((z.status = 'church_member') or (z.status = 'left_church') or (z.deceased_on is not null)) " + vbCrLf)
                ElseIf statusFlag = 11 Then '새교우 + 고인
                    .Append(" where ((z.status = 'new_family') or (z.deceased_on is not null))  " + vbCrLf)
                ElseIf statusFlag = 12 Then '이적교인
                    .Append(" where ((z.status = 'left_church' and  z.deceased_on is null))  " + vbCrLf)
                ElseIf statusFlag = 13 Then '이적교인 + 고인
                    .Append(" where ((z.status = 'left_church') or (z.deceased_on is not null))  " + vbCrLf)
                ElseIf statusFlag = 14 Then '고인
                    .Append(" where z.deceased_on is not null  " + vbCrLf)
                ElseIf statusFlag = 15 Then '무적
                    .Append(" where z.status is null " + vbCrLf)
                End If

                If String.IsNullOrEmpty(name) = False Then
                    If flag = 2 Then '전화번호
                        If filterFlag = 3 Then
                            .Append("and translate(z.cell_phone, '- ', '') = @name or translate(z.home_phone, '- ', '') = @name or translate(z.work_phone, '- ', '') = @name " + vbCrLf)
                        Else
                            .Append("and translate(z.cell_phone, '- ', '') like @name or translate(z.home_phone, '- ', '') like @name or translate(z.work_phone, '- ', '') like @name " + vbCrLf)
                        End If
                    ElseIf flag = 3 Then '직분명
                        If filterFlag = 3 Then
                            If name = "중직" Or name = "중직자" Then
                                .Append("and z.title in ('시무장로','시무권사','안수집사') " + vbCrLf)
                            Else
                                .Append("and z.title = @name " + vbCrLf)
                            End If
                        Else
                            If name = "중직" Or name = "중직자" Then
                                .Append("and z.title in ('시무장로','시무권사','안수집사') " + vbCrLf)
                            Else
                                .Append("and z.title like @name " + vbCrLf)
                            End If

                        End If

                        If titleYear <> "모든년도" And String.IsNullOrEmpty(titleYear) = False Then
                            .Append("and date_part('year', z.title_date) =@titleYear ")
                        End If

                    ElseIf flag = 4 Then '도시명
                        If filterFlag = 3 Then
                            .Append("and lower(z.address) = @name " + vbCrLf)
                        Else
                            .Append("and lower(z.address) like @name " + vbCrLf)
                        End If
                    ElseIf flag = 7 Then '이메일
                        If filterFlag = 3 Then
                            .Append("and lower(z.email) = @name " + vbCrLf)
                        Else
                            .Append("and lower(z.email) like @name " + vbCrLf)
                        End If
                    ElseIf flag = 9 Then '독신여부
                        If name.ToLower = "y" Or name.ToLower = "yes" Or name = "네" Or name = "예" Or name = "예스" Then
                            .Append("and cast(z.person_id as varchar) = z.spouses " + vbCrLf)
                        Else
                            .Append("and cast(z.person_id as varchar) <> z.spouses " + vbCrLf)
                        End If
                    End If
                End If

                If flag = 6 Then '나이
                    If String.IsNullOrEmpty(ageFrom) = False And String.IsNullOrEmpty(ageTo) = False Then
                        .Append("and z.age between @ageFrom and @ageTo " + vbCrLf)
                    Else
                        If String.IsNullOrEmpty(ageFrom) = False Then
                            .Append("and z.age = @ageFrom " + vbCrLf)
                        Else
                            .Append("and z.age = @ageTo " + vbCrLf)
                        End If
                    End If
                ElseIf flag = 8 Then '한글이름중복
                    If String.IsNullOrEmpty(ageFrom) = False And String.IsNullOrEmpty(ageTo) = False Then
                        .Append("and z.kn_cnt between @ageFrom and @ageTo " + vbCrLf)
                    Else
                        If String.IsNullOrEmpty(ageFrom) = False Then
                            .Append("and z.kn_cnt = @ageFrom " + vbCrLf)
                        Else
                            .Append("and z.kn_cnt = @ageTo " + vbCrLf)
                        End If
                    End If
                ElseIf flag = 10 Then '주소갯수
                    If String.IsNullOrEmpty(ageFrom) = False And String.IsNullOrEmpty(ageTo) = False Then
                        .Append("and z.addr_cnt between @ageFrom and @ageTo " + vbCrLf)
                    Else
                        If String.IsNullOrEmpty(ageFrom) = False Then
                            .Append("and z.addr_cnt = @ageFrom " + vbCrLf)
                        Else
                            .Append("and z.addr_cnt = @ageTo " + vbCrLf)
                        End If
                    End If

                ElseIf flag = 11 Then '출생년도
                    If String.IsNullOrEmpty(ageFrom) = False And String.IsNullOrEmpty(ageTo) = False Then
                        .Append("and z.dob_year between @ageFrom and @ageTo " + vbCrLf)
                    Else
                        If String.IsNullOrEmpty(ageFrom) = False Then
                            .Append("and z.dob_year = @ageFrom " + vbCrLf)
                        Else
                            .Append("and z.dob_year = @ageTo " + vbCrLf)
                        End If
                    End If
                End If

                If flag = 1 Then
                    .Append("order by z.year desc, z.korean_name COLLATE " + ControlChars.Quote + "C" + ControlChars.Quote)
                Else
                    .Append("order by z.korean_name COLLATE " + ControlChars.Quote + "C" + ControlChars.Quote)
                End If

            End With


            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                If String.IsNullOrEmpty(name) = False Then
                    If flag = 0 Then 'contains - by name
                        If filterFlag = 0 Then 'containing
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = "%" + name.ToLower + "%"
                        ElseIf filterFlag = 1 Then 'beginning
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = name.ToLower + "%"
                        ElseIf filterFlag = 2 Then 'ending
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = "%" + name.ToLower
                        ElseIf filterFlag = 3 Then 'matching
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = name.ToLower
                        End If

                    ElseIf flag = 1 Then 'by offering_number
                        cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(name, Integer)
                    ElseIf flag = 2 Then 'by telephone
                        If filterFlag = 0 Then
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = "%" + Replace(name, "-", "") + "%"
                        ElseIf filterFlag = 1 Then
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = Replace(name, "-", "") + "%"
                        ElseIf filterFlag = 2 Then
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = "%" + Replace(name, "-", "")
                        ElseIf filterFlag = 3 Then
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = Replace(name, "-", "")
                        End If

                    ElseIf flag = 3 Then 'by title
                        If name <> "중직" And name <> "중직자" Then
                            If filterFlag = 0 Then
                                cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = "%" + name + "%"
                            ElseIf filterFlag = 1 Then
                                cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = name + "%"
                            ElseIf filterFlag = 2 Then
                                cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = "%" + name
                            ElseIf filterFlag = 3 Then
                                cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = name
                            End If
                        End If

                        If titleYear <> "모든년도" And String.IsNullOrEmpty(titleYear) = False Then
                            cmd.Parameters.Add(New NpgsqlParameter("@titleYear", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(titleYear, Integer)
                        End If

                    ElseIf flag = 4 Or flag = 7 Then 'by city or email
                        If filterFlag = 0 Then
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = "%" + name.ToLower + "%"
                        ElseIf filterFlag = 1 Then
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = name.ToLower + "%"
                        ElseIf filterFlag = 2 Then
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = "%" + name.ToLower
                        ElseIf filterFlag = 3 Then
                            cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = name.ToLower
                        End If
                    End If
                Else
                    If flag = 6 Or flag = 8 Or flag = 10 Then '나이, 한글이름 중복, 주소갯수
                        If String.IsNullOrEmpty(ageFrom) = False And String.IsNullOrEmpty(ageTo) = False Then
                            cmd.Parameters.Add(New NpgsqlParameter("@ageFrom", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(ageFrom, Integer)
                            cmd.Parameters.Add(New NpgsqlParameter("@ageTo", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(ageTo, Integer)
                        Else
                            If String.IsNullOrEmpty(ageFrom) = False Then
                                cmd.Parameters.Add(New NpgsqlParameter("@ageFrom", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(ageFrom, Integer)
                            Else
                                cmd.Parameters.Add(New NpgsqlParameter("@ageTo", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(ageTo, Integer)
                            End If
                        End If
                    ElseIf flag = 11 Then '출생년도
                        If String.IsNullOrEmpty(ageFrom) = False And String.IsNullOrEmpty(ageTo) = False Then
                            cmd.Parameters.Add(New NpgsqlParameter("@ageFrom", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = CType(ageFrom, String)
                            cmd.Parameters.Add(New NpgsqlParameter("@ageTo", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = CType(ageTo, String)
                        Else
                            If String.IsNullOrEmpty(ageFrom) = False Then
                                cmd.Parameters.Add(New NpgsqlParameter("@ageFrom", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = CType(ageFrom, String)
                            Else
                                cmd.Parameters.Add(New NpgsqlParameter("@ageTo", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = CType(ageTo, String)
                            End If
                        End If
                    End If
                End If

                da.SelectCommand = cmd
                da.Fill(findPeopleList, "findPeopleList")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.findPeopleList : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return findPeopleList

        End Function

        Public Function getPeopleList(ByVal flag As Integer, ByVal personIds As String, ByVal startDate As String, ByVal endDate As String, ByVal statusFlag As Integer) As DataSet
            Dim sql As String = String.Empty
            getPeopleList = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder
            Dim other_status As String = "church_member"

            Logger.LogInfo("flag=[" + CType(flag, String) + "], personIds=[" + personIds + "], startDate=[" + startDate + "], endDate = [" + endDate + "], statusFlag=[" + CType(statusFlag, String) + "]")

            If flag = 5 Then '등록 리스트
                other_status = "left_church"
            End If

            Dim selectStr As String = "select z.person_id, z.korean_name, z.title, INITCAP(z.last_name) last_name, INITCAP(z.first_name) first_name, z.gender, z.dob, z.age,   " + vbCrLf +
                            "       case " + vbCrLf +
                            "       WHEN (z.age <= 0) THEN '0세 이하' " + vbCrLf +
                            "       WHEN (z.age > 0) AND (z.age <= 10) THEN '1-10세' " + vbCrLf +
                            "       WHEN (z.age > 10) AND (z.age <= 20) THEN '11-20세' " + vbCrLf +
                            "       WHEN (z.age > 20) AND (z.age <= 30) THEN '21-30세' " + vbCrLf +
                            "       WHEN (z.age > 30) AND (z.age <= 40) THEN '31-40세' " + vbCrLf +
                            "       WHEN (z.age > 40) AND (z.age <= 50) THEN '41-50세' " + vbCrLf +
                            "       WHEN (z.age > 50) AND (z.age <= 60) THEN '51-60세' " + vbCrLf +
                            "       WHEN (z.age > 60) AND (z.age <= 70) THEN '61-70세' " + vbCrLf +
                            "       WHEN (z.age > 70) AND (z.age <= 80) THEN '71-80세' " + vbCrLf +
                            "       WHEN (z.age > 80) AND (z.age <= 90) THEN '81-90세' " + vbCrLf +
                            "       WHEN (z.age > 90) AND (z.age <= 100) THEN '91-100세' " + vbCrLf +
                            "       WHEN (z.age > 100) AND (z.age <= 110) THEN '101-110세' " + vbCrLf +
                            "       ELSE '110세 이후' " + vbCrLf +
                            "       end year_range,	 " + vbCrLf +
                            "       z.cell_phone, z.home_phone, z.work_phone, " +
                            "       z.address, z.infant_baptism, z.infant_baptism_date, z.confirmation, z.confirmation_date, z.baptism, z.baptism_date, " + vbCrLf +
                            "       z.status, z.status_date, " + vbCrLf +
                            "       CASE WHEN z.status = 'church_member' THEN age(z.status_date) " + vbCrLf +
                            "            WHEN z.status = 'left_church' THEN age(z.status_date, z.other_status_date) " + vbCrLf +
                            "            WHEN z.status = 'new_family' THEN age(z.status_date) " + vbCrLf +
                            "            END stay, " + vbCrLf +
                            "       z.deceased_on as deceased_date, z.spouses, z.photo_file_name "

            With sb
                .Append(selectStr + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("    Select p.id As person_id, korean_name, last_name, first_name, gender, date_of_birth dob,  " + vbCrLf)
                .Append("           date_part('year', age(date_of_birth)) age, cell_phone, home_phone, work_phone, " + vbCrLf)
                .Append("           deceased_on,  " + vbCrLf)
                .Append("           (select x.address  " + vbCrLf)
                .Append("            from (  " + vbCrLf)
                .Append("               select INITCAP(a.street) || ' ' || INITCAP(a.city) || ' ' || INITCAP(CASE WHEN lower(a.province)='ontario' THEN 'Ont.' ELSE a.province END) || ' ' || a.postal_code as address, ROW_NUMBER() OVER (PARTITION BY a.person_id ORDER BY created_at) row_num  " + vbCrLf)
                .Append("               from public.addresses a where p.id = a.person_id and a.address_type = 'home'  " + vbCrLf)
                .Append("             ) x  " + vbCrLf)
                .Append("             where x.row_num = 1  " + vbCrLf)
                .Append("           ) address, " + vbCrLf)
                .Append("           (select x.name " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'TitleMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) title, " + vbCrLf)
                .Append("           (select x.name " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'MembershipMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) status, " + vbCrLf)
                .Append("           (select x.effective_date " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'MembershipMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("           where x.row_num = 1 " + vbCrLf)
                .Append("           ) status_date, " + vbCrLf)
                .Append("          (select x.effective_date " + vbCrLf)
                .Append("           from( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num  " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id And name='" + other_status + "' and type = 'MembershipMilestone'  " + vbCrLf)
                .Append("           ) x  " + vbCrLf)
                .Append("           where x.row_num = 1  " + vbCrLf)
                .Append("           ) other_status_date,  " + vbCrLf)
                .Append("          (select x.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'baptism' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) baptism, " + vbCrLf)
                .Append("          (select x.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'baptism' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) baptism_date, " + vbCrLf)

                .Append("          (select x.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'infant_baptism' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) infant_baptism, " + vbCrLf)
                .Append("          (select x.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'infant_baptism' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) infant_baptism_date, " + vbCrLf)

                .Append("          (select x.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'confirmation' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) confirmation, " + vbCrLf)
                .Append("          (select x.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where p.id = m.person_id and type = 'BaptismMilestone' and name = 'confirmation' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) confirmation_date, " + vbCrLf)
                .Append("     (select CASE WHEN string_agg(cast(y.f_id as varchar), ',' ORDER BY y.f_id) IS NULL THEN cast(p.id as varchar) " + vbCrLf)
                .Append("             ELSE string_agg(cast(y.f_id As varchar), ',' ORDER BY y.f_id) END " + vbCrLf)
                .Append("      from (  " + vbCrLf)
                .Append("           select distinct r.person_id f_id, r.relationship_type as rtype  " + vbCrLf)
                .Append("           from Public.relationships r   " + vbCrLf)
                .Append("           where p.id = r.related_person_id   " + vbCrLf)
                .Append("           union all   " + vbCrLf)
                .Append("           select distinct r.related_person_id f_id, r.relationship_type as rtype  " + vbCrLf)
                .Append("           from Public.relationships r   " + vbCrLf)
                .Append("           where p.id = r.person_id   " + vbCrLf)
                .Append("           union all   " + vbCrLf)
                .Append("           select distinct r.person_id f_id, r.relationship_type as rtype  " + vbCrLf)
                .Append("           from Public.relationships r   " + vbCrLf)
                .Append("           where p.id = r.person_id   " + vbCrLf)
                .Append("           union all   " + vbCrLf)
                .Append("           select distinct r.related_person_id f_id, r.relationship_type as rtype  " + vbCrLf)
                .Append("           from Public.relationships r   " + vbCrLf)
                .Append("           where p.id = r.related_person_id   " + vbCrLf)
                .Append("       ) y    " + vbCrLf)
                .Append("       where y.rtype = 'spouse') spouses, photo_file_name " + vbCrLf)
                .Append("   from public.people p  " + vbCrLf)

                If String.IsNullOrEmpty(personIds) = False Then
                    .Append("   where p.id in (" + personIds + ") " + vbCrLf)
                End If

                .Append(") z " + vbCrLf)

                If flag = 0 Then '고인 리스트
                    .Append(" where z.deceased_on between @startDate and @endDate " + vbCrLf)

                ElseIf flag = 1 Then '유아세례 리스트
                    .Append(" where z.infant_baptism_date between @startDate and @endDate " + vbCrLf)

                ElseIf flag = 2 Then '입교 리스트
                    .Append(" where z.confirmation_date between @startDate and @endDate " + vbCrLf)

                ElseIf flag = 3 Then '세례 리스트
                    .Append(" where z.baptism_date between @startDate and @endDate " + vbCrLf)

                ElseIf flag = 4 Then '새교우 리스트
                    .Append(" where z.status_date between @startDate and @endDate " + vbCrLf)

                ElseIf flag = 5 Then '등록 리스트
                    .Append(" where z.status_date between @startDate and @endDate " + vbCrLf)

                ElseIf flag = 6 Then '이적 리스트
                    .Append(" where z.status_date between @startDate and @endDate " + vbCrLf)

                ElseIf flag = 7 Then '생일 리스트
                    .Append(" where to_char(z.dob, 'MM') between @startDate and @endDate " + vbCrLf)

                End If

                If statusFlag = 0 Then '모두
                ElseIf statusFlag = 1 Then '등록교인만
                    .Append(" and (z.status = 'church_member' and z.deceased_on is null) " + vbCrLf)
                ElseIf statusFlag = 2 Then '등록교인 + 새교우
                    .Append(" and ((z.status = 'church_member' and z.deceased_on is null) or (z.status = 'new_family' and  z.deceased_on is null)) " + vbCrLf)
                ElseIf statusFlag = 3 Then '등록교인 + 새교우 + 이적교인
                    .Append(" and ((z.status = 'church_member' and z.deceased_on is null) or (z.status = 'new_family' and  z.deceased_on is null) or (z.status = 'left_church' and  z.deceased_on is null))  " + vbCrLf)
                ElseIf statusFlag = 4 Then '등록교인 + 새교우 + 고인
                    .Append(" and ((z.status = 'church_member') or (z.status = 'left_church') or (z.deceased_on is not null)) " + vbCrLf)
                ElseIf statusFlag = 5 Then '등록교인 + 이적교인
                    .Append(" and ((z.status = 'church_member' and z.deceased_on is null) or (z.status = 'left_church' and  z.deceased_on is null))  " + vbCrLf)
                ElseIf statusFlag = 6 Then '등록교인 + 이적교인 + 고인
                    .Append(" and ((z.status = 'church_member') or (z.status = 'left_church')  or (z.deceased_on is not null)) " + vbCrLf)
                ElseIf statusFlag = 7 Then '등록교인 + 고인
                    .Append(" and ((z.status = 'church_member') or (z.deceased_on is not null)) " + vbCrLf)
                ElseIf statusFlag = 8 Then '새교우
                    .Append(" and ((z.status = 'new_family' and z.deceased_on is null)) " + vbCrLf)
                ElseIf statusFlag = 9 Then '새교우 + 이적교인
                    .Append(" and ((z.status = 'new_family' and z.deceased_on is null) or (z.status = 'left_church' and  z.deceased_on is null))  " + vbCrLf)
                ElseIf statusFlag = 10 Then '새교우 + 이적교인 + 고인
                    .Append(" and ((z.status = 'church_member') or (z.status = 'left_church') or (z.deceased_on is not null)) " + vbCrLf)
                ElseIf statusFlag = 11 Then '새교우 + 고인
                    .Append(" and ((z.status = 'new_family') or (z.deceased_on is not null))  " + vbCrLf)
                ElseIf statusFlag = 12 Then '이적교인
                    .Append(" and ((z.status = 'left_church' and  z.deceased_on is null))  " + vbCrLf)
                ElseIf statusFlag = 13 Then '이적교인 + 고인
                    .Append(" and ((z.status = 'left_church') or (z.deceased_on is not null))  " + vbCrLf)
                ElseIf statusFlag = 14 Then '고인
                    .Append(" and z.deceased_on is not null  " + vbCrLf)
                ElseIf statusFlag = 15 Then '무적
                    .Append(" and z.status is null  " + vbCrLf)
                End If

                .Append("order by z.korean_name COLLATE " + ControlChars.Quote + "C" + ControlChars.Quote)
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                If flag <> 8 Then
                    If flag = 7 Then '생일 리스트
                        cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = startDate
                        cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = endDate
                    Else
                        cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(startDate, Date)
                        cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(endDate, Date)
                    End If
                End If

                da.SelectCommand = cmd
                da.Fill(getPeopleList, "getPeopleList")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getPeopleList : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getPeopleList
        End Function

        Public Function getPeopleListWithOfferingNumber(ByRef offeringNumber As String) As DataSet
            Dim sql As String = String.Empty
            getPeopleListWithOfferingNumber = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder
            Dim paramStr As String = ""

            If String.IsNullOrEmpty(offeringNumber) = False Then
                offeringNumber = offeringNumber.Trim
            End If

            Logger.LogInfo("offeringNumber=[" + offeringNumber + "]")

            With sb
                .Append("select e.person_id, e.offering_number, e.year, p.korean_name, a.street, a.city, a.province, a.postal_code  ")
                .Append("from public.people p, public.addresses a, public.offering_envelopes e ")
                .Append("where p.id = a.person_id and a.person_id = e.person_id and e.person_id = p.id ")
                If String.IsNullOrEmpty(offeringNumber) = False Then
                    .Append("and e.offering_number = @offeringNumber  ")
                End If
                .Append("order by e.year desc, e.offering_number, p.korean_name COLLATE " + ControlChars.Quote + "C" + ControlChars.Quote + ", e.person_id")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()

                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                If String.IsNullOrEmpty(offeringNumber) = False Then
                    cmd.Parameters.Add(New NpgsqlParameter("@offeringNumber", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(offeringNumber, Integer)
                End If

                da.SelectCommand = cmd
                da.Fill(getPeopleListWithOfferingNumber, "getPeopleListWithOfferingNumber")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getPeopleListWithOfferingNumber :  " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getPeopleListWithOfferingNumber
        End Function

        Public Function getCoupleList(ByVal flag As Integer, ByVal names As String, ByVal statusFlag As Integer) As DataSet
            Dim sql As String = String.Empty
            getCoupleList = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("flag=[" + CType(flag, String) + "], names=[" + names + "]")

            Dim whereStr As String = ""
            Dim wherePrefix As String = " where "
            Dim whichStatus As String = "y.k1_status "
            Dim whichDeceased As String = "y.k1_deceased "

            If String.IsNullOrEmpty(names) = False Then
                wherePrefix = " and "
            End If

            If flag = 1 Then
                whichStatus = "y.k2_status "
                whichDeceased = "y.k2_deceased "
            End If

            If statusFlag = 0 Then '모두
            ElseIf statusFlag = 1 Or statusFlag = 15 Then '등록교인만
                whereStr = wherePrefix + " (" + whichStatus + " = 'church_member' and " + whichDeceased + " is null) " + vbCrLf
            ElseIf statusFlag = 2 Then '등록교인 + 새교우
                whereStr = wherePrefix + " ((" + whichStatus + " = 'church_member' and " + whichDeceased + " is null) or (" + whichStatus + " = 'new_family' and " + whichDeceased + " is null)) " + vbCrLf
            ElseIf statusFlag = 3 Then '등록교인 + 새교우 + 이적교인
                whereStr = wherePrefix + " ((" + whichStatus + " = 'church_member' and " + whichDeceased + " is null) or (" + whichStatus + " = 'new_family' and " + whichDeceased + " is null) or (" + whichStatus + " = 'left_church' and " + whichDeceased + " is null))  " + vbCrLf
            ElseIf statusFlag = 4 Then '등록교인 + 새교우 + 고인
                whereStr = wherePrefix + " ((" + whichStatus + " = 'church_member') or (" + whichStatus + " = 'left_church') or ( " + whichDeceased + " is not null)) " + vbCrLf
            ElseIf statusFlag = 5 Then '등록교인 + 이적교인
                whereStr = wherePrefix + " ((" + whichStatus + " = 'church_member' and " + whichDeceased + " is null) or (" + whichStatus + " = 'left_church' and " + whichDeceased + " is null))  " + vbCrLf
            ElseIf statusFlag = 6 Then '등록교인 + 이적교인 + 고인
                whereStr = wherePrefix + " ((" + whichStatus + " = 'church_member') or (" + whichStatus + " = 'left_church')  or ( " + whichDeceased + " is not null)) " + vbCrLf
            ElseIf statusFlag = 7 Then '등록교인 + 고인
                whereStr = wherePrefix + " ((" + whichStatus + " = 'church_member') or ( " + whichDeceased + " is not null)) " + vbCrLf
            ElseIf statusFlag = 8 Then '새교우
                whereStr = wherePrefix + " ((" + whichStatus + " = 'new_family' and " + whichDeceased + " is null)) " + vbCrLf
            ElseIf statusFlag = 9 Then '새교우 + 이적교인
                whereStr = wherePrefix + " ((" + whichStatus + " = 'new_family' and " + whichDeceased + " is null) or (" + whichStatus + " = 'left_church' and " + whichDeceased + " is null))  " + vbCrLf
            ElseIf statusFlag = 10 Then '새교우 + 이적교인 + 고인
                whereStr = wherePrefix + " ((" + whichStatus + " = 'church_member') or (" + whichStatus + " = 'left_church') or ( " + whichDeceased + " is not null)) " + vbCrLf
            ElseIf statusFlag = 11 Then '새교우 + 고인
                whereStr = wherePrefix + " ((" + whichStatus + " = 'new_family') or ( " + whichDeceased + " is not null))  " + vbCrLf
            ElseIf statusFlag = 12 Then '이적교인
                whereStr = wherePrefix + " ((" + whichStatus + " = 'left_church' and " + whichDeceased + " is null))  " + vbCrLf
            ElseIf statusFlag = 13 Then '이적교인 + 고인
                whereStr = wherePrefix + " ((" + whichStatus + " = 'left_church') or ( " + whichDeceased + " is not null))  " + vbCrLf
            ElseIf statusFlag = 14 Then '고인
                whereStr = wherePrefix + "  " + whichDeceased + " is not null  " + vbCrLf
            End If

            Dim selectStr As String = ""

            If flag = 0 Then 'Husband first
                selectStr = "select y.k1 As 남편이름, CASE WHEN y.k1_deceased is not null THEN '고인' ELSE CASE When y.k1_status = 'church_member' THEN '등록' WHEN y.k1_status = 'left_church' THEN '이적' WHEN y.k1_status = 'new_family' THEN '새교우' ELSE '' END END as 남편_등록상태, " +
                            "       y.k2 as 부인이름, CASE WHEN y.k2_deceased is not null THEN '고인' ELSE CASE WHEN y.k2_status = 'church_member' THEN '등록' WHEN y.k2_status = 'left_church' THEN '이적' WHEN y.k2_status = 'new_family' THEN '새교우' ELSE '' END END as 부인_등록상태 "
            Else
                selectStr = "select y.k2 as 부인이름, CASE WHEN y.k2_deceased is not null THEN '고인' ELSE CASE WHEN y.k2_status = 'church_member' THEN '등록' WHEN y.k2_status = 'left_church' THEN '이적' WHEN y.k2_status = 'new_family' THEN '새교우' ELSE '' END END as 부인_등록상태, " +
                            "       y.k1 As 남편이름, CASE WHEN y.k1_deceased is not null THEN '고인' ELSE CASE When y.k1_status = 'church_member' THEN '등록' WHEN y.k1_status = 'left_church' THEN '이적' WHEN y.k1_status = 'new_family' THEN '새교우' ELSE '' END END as 남편_등록상태 "
            End If

            With sb
                .Append(selectStr)
                .Append("from (" + vbCrLf)
                .Append("   select" + vbCrLf)
                .Append("       case z.g1 when 'F' then z.k2" + vbCrLf)
                .Append("           when 'M' then z.k1" + vbCrLf)
                .Append("       end k1," + vbCrLf)
                .Append("       case z.g2 when 'F' then z.k2" + vbCrLf)
                .Append("           when 'M' then z.k1" + vbCrLf)
                .Append("       end k2, " + vbCrLf)
                .Append("       case z.g1 when 'F' then z.k2_status" + vbCrLf)
                .Append("           when 'M' then z.k1_status" + vbCrLf)
                .Append("       end k1_status," + vbCrLf)
                .Append("       case z.g2 when 'F' then z.k2_status" + vbCrLf)
                .Append("           when 'M' then z.k1_status" + vbCrLf)
                .Append("       end k2_status," + vbCrLf)
                .Append("       case z.g1 when 'F' then z.k2_deceased" + vbCrLf)
                .Append("           when 'M' then z.k1_deceased" + vbCrLf)
                .Append("       end k1_deceased," + vbCrLf)
                .Append("       case z.g2 when 'F' then z.k2_deceased" + vbCrLf)
                .Append("           when 'M' then z.k1_deceased" + vbCrLf)
                .Append("       end k2_deceased" + vbCrLf)
                .Append("   from (" + vbCrLf)
                .Append("       select " + vbCrLf)
                .Append("           (select korean_name from public.people where id = r.person_id) k1," + vbCrLf)
                .Append("           (select gender from public.people where id = r.person_id) g1," + vbCrLf)
                .Append("           (select (select x.name " + vbCrLf)
                .Append("                    from ( " + vbCrLf)
                .Append("                        select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("                        from public.milestones m where p.id = m.person_id And type = 'MembershipMilestone' " + vbCrLf)
                .Append("                    ) x " + vbCrLf)
                .Append("                    where x.row_num = 1) status " + vbCrLf)
                .Append("            from public.people p where id = r.person_id) k1_status, " + vbCrLf)
                .Append("           (select deceased_on from public.people where id = r.person_id) k1_deceased," + vbCrLf)
                .Append("           (select korean_name from public.people where id = r.related_person_id) k2," + vbCrLf)
                .Append("           (select gender from public.people where id = r.related_person_id) g2, " + vbCrLf)
                .Append("           (select (select x.name " + vbCrLf)
                .Append("                    from ( " + vbCrLf)
                .Append("                        select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("                        from public.milestones m where p.id = m.person_id And type = 'MembershipMilestone' " + vbCrLf)
                .Append("                    ) x " + vbCrLf)
                .Append("                    where x.row_num = 1) status " + vbCrLf)
                .Append("            from public.people p where id = r.related_person_id) k2_status, " + vbCrLf)
                .Append("           (select deceased_on from public.people where id = r.related_person_id) k2_deceased" + vbCrLf)
                .Append("       from public.relationships r" + vbCrLf)
                .Append("       where relationship_type = 'spouse'" + vbCrLf)
                .Append("   ) z" + vbCrLf)
                .Append(") y " + vbCrLf)

                If String.IsNullOrEmpty(names) = False Then
                    .Append("where y.k1 in (" + names + ") or y.k2 in (" + names + ") " + vbCrLf)
                End If

                .Append(whereStr)
                .Append("group by y.k1, y.k2, y.k1_status, y.k2_status, y.k1_deceased, y.k2_deceased " + vbCrLf)

                If flag = 0 Then
                    .Append("order by y.k1 COLLATE " + ControlChars.Quote + "C" + ControlChars.Quote)
                Else
                    .Append("order by y.k2 COLLATE " + ControlChars.Quote + "C" + ControlChars.Quote)
                End If
            End With


            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0
                da.SelectCommand = cmd
                da.Fill(getCoupleList, "getCoupleList")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getCoupleList : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getCoupleList

        End Function

        Public Function getPeopleFindFamilyKey(ByVal flag As Integer, ByVal personId As String) As List(Of String)
            Dim sql As String = String.Empty
            getPeopleFindFamilyKey = New List(Of String)
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("flag=[" + CType(flag, String) + "], personId=[" + personId + "]")

            'check if spouse first first
            'and if spouse is not there, then f_id != id

            If flag = 0 Then 'spouse only
                With sb
                    .Append("select distinct z.f_id " + vbCrLf)
                    .Append("from ( " + vbCrLf)
                    .Append("   select distinct x.f_id, x.id, " + vbCrLf)
                    .Append("          x.korean_name, x.rtype " + vbCrLf)
                    .Append("   from (  " + vbCrLf)
                    .Append("       select distinct r.person_id f_id, p.id, p.korean_name, r.relationship_type as rtype " + vbCrLf)
                    .Append("       from public.people p, public.relationships r  " + vbCrLf)
                    .Append("       where p.id = r.related_person_id  " + vbCrLf)

                    If (String.IsNullOrEmpty(personId)) = False Then
                        .Append("   and p.id in ( " + personId + ") " + vbCrLf)
                    End If

                    .Append("       union all  " + vbCrLf)
                    .Append("       select distinct r.related_person_id f_id, p.id, p.korean_name, r.relationship_type as rtype " + vbCrLf)
                    .Append("       from public.people p, public.relationships r  " + vbCrLf)
                    .Append("       where p.id = r.person_id  " + vbCrLf)

                    If (String.IsNullOrEmpty(personId)) = False Then
                        .Append("   and p.id in ( " + personId + ") " + vbCrLf)
                    End If

                    .Append("       union all  " + vbCrLf)
                    .Append("       select distinct r.person_id f_id, p.id, p.korean_name, r.relationship_type as rtype " + vbCrLf)
                    .Append("       from public.people p, public.relationships r  " + vbCrLf)
                    .Append("       where p.id = r.person_id  " + vbCrLf)

                    If (String.IsNullOrEmpty(personId)) = False Then
                        .Append("   and p.id in ( " + personId + ") " + vbCrLf)
                    End If

                    .Append("       union all  " + vbCrLf)
                    .Append("       select distinct r.related_person_id f_id, p.id, p.korean_name, r.relationship_type as rtype " + vbCrLf)
                    .Append("       from public.people p, public.relationships r  " + vbCrLf)
                    .Append("       where p.id = r.related_person_id  " + vbCrLf)

                    If (String.IsNullOrEmpty(personId)) = False Then
                        .Append("   and p.id in ( " + personId + ") " + vbCrLf)
                    End If

                    .Append("   ) x  " + vbCrLf)
                    .Append(") z " + vbCrLf)
                    .Append("where z.rtype = 'spouse'  " + vbCrLf)
                    .Append("order by z.f_id " + vbCrLf)
                End With
            ElseIf flag = 1 Then 'Other family except the spouse
                With sb
                    .Append("select z.f_id, z.id, z.rank, z.korean_name, z.rtype " + vbCrLf)
                    .Append("from ( " + vbCrLf)
                    .Append("   select distinct x.f_id, x.id, " + vbCrLf)
                    .Append("          ROW_NUMBER() over (partition by x.f_id) rank, " + vbCrLf)
                    .Append("          x.korean_name, x.rtype " + vbCrLf)
                    .Append("   from (  " + vbCrLf)
                    .Append("       select distinct r.person_id f_id, p.id, p.korean_name, r.relationship_type as rtype " + vbCrLf)
                    .Append("       from public.people p, public.relationships r  " + vbCrLf)
                    .Append("       where p.id = r.related_person_id  " + vbCrLf)

                    If (String.IsNullOrEmpty(personId)) = False Then
                        .Append("   and p.id in (" + personId + ") " + vbCrLf)
                    End If

                    .Append("       union all  " + vbCrLf)
                    .Append("       select distinct r.related_person_id f_id, p.id, p.korean_name, r.relationship_type as rtype " + vbCrLf)
                    .Append("       from public.people p, public.relationships r  " + vbCrLf)
                    .Append("       where p.id = r.person_id  " + vbCrLf)

                    If (String.IsNullOrEmpty(personId)) = False Then
                        .Append("   and p.id in (" + personId + ") " + vbCrLf)
                    End If

                    '.Append("       union all  " + vbCrLf)
                    '.Append("       select distinct r.person_id f_id, p.id, p.korean_name, r.relationship_type as rtype " + vbCrLf)
                    '.Append("       from Public.people p, Public.relationships r  " + vbCrLf)
                    '.Append("       where p.id = r.person_id  " + vbCrLf)

                    'If (String.IsNullOrEmpty(personId)) = False Then
                    '    .Append("   and p.id = @personId " + vbCrLf)
                    'End If

                    .Append("       union all  " + vbCrLf)
                    .Append("       select distinct r.related_person_id f_id, p.id, p.korean_name, r.relationship_type as rtype " + vbCrLf)
                    .Append("       from public.people p, public.relationships r  " + vbCrLf)
                    .Append("       where p.id = r.related_person_id  " + vbCrLf)

                    If (String.IsNullOrEmpty(personId)) = False Then
                        .Append("   and p.id in (" + personId + ") " + vbCrLf)
                    End If

                    .Append("   ) x  " + vbCrLf)
                    .Append(") z " + vbCrLf)
                    .Append("where z.f_id != z.id " + vbCrLf)
                    '.Append("order by z.korean_name " + vbCrLf)
                End With
            Else 'check if this is parent or child
                With sb
                    .Append("       select distinct r.person_id f_id, p.id, p.korean_name, r.relationship_type as rtype " + vbCrLf)
                    .Append("       from public.people p, public.relationships r  " + vbCrLf)
                    .Append("       where p.id = r.related_person_id  " + vbCrLf)

                    If (String.IsNullOrEmpty(personId)) = False Then
                        .Append("   and p.id in (" + personId + ") " + vbCrLf)
                    End If
                End With
            End If

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                'If (String.IsNullOrEmpty(personId)) = False Then
                '    cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personId, Integer)
                'End If

                dr = cmd.ExecuteReader
                'Dim parentIds As String = ""
                'Dim spouseIds As String = ""
                While dr.Read()
                    getPeopleFindFamilyKey.Add(dr("f_id").ToString)
                    'If dr("rtype").ToString = "child" Then
                    '    parentIds += dr("f_id").ToString + ","
                    'Else
                    '    spouseIds += dr("f_id").ToString + ","
                    'End If
                End While

                'If String.IsNullOrEmpty(parentIds) = False Then
                '    parentIds = parentIds.Substring(0, parentIds.Length - 1)
                '    getPeopleFindFamilyKey.Add(parentIds)
                'End If

                'If String.IsNullOrEmpty(spouseIds) = False Then
                '    spouseIds = spouseIds.Substring(0, spouseIds.Length - 1)
                '    getPeopleFindFamilyKey.Add(spouseIds)
                'End If

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getPeopleFindFamilyKey : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getPeopleFindFamilyKey
        End Function
        Public Function getAllPeopleId() As List(Of String)
            Dim sql As String = String.Empty
            getAllPeopleId = New List(Of String)
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            'check if spouse first first
            'and if spouse is not there, then f_id != id

            With sb
                .Append("select p.id ")
                .Append("from public.people p ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                dr = cmd.ExecuteReader
                'Dim parentIds As String = ""
                'Dim spouseIds As String = ""
                While dr.Read()
                    getAllPeopleId.Add(dr("id").ToString)
                End While

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getPeopleFindFamilyKey : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getAllPeopleId
        End Function
        Public Function getPeopleFindWithFamily(ByVal personIds As String, ByVal singleParentPersonId As String) As DataTable
            Dim sql As String = String.Empty
            getPeopleFindWithFamily = New DataTable
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("personIds=[" + personIds + "], singleParentPersonId=[" + singleParentPersonId + "]")

            Dim relationshipType As String = "spouse"

            If singleParentPersonId.Contains("ns") = True Then
                relationshipType = "parent"
            End If

            With sb
                .Append("select z.lvl, z.pid as id, z.korean_name, INITCAP(z.last_name) last_name, INITCAP(z.first_name) first_name, z.gender, " + vbCrLf)
                .Append("       z.dob, z.age, z.title, z.cell, z.home, z.work, z.email, z.rtype as type, z.address, z.address2, z.status, z.status_date, " + vbCrLf)
                .Append("       z.infant_baptism, z.infant_baptism_date, z.confirmation, z.confirmation_date, " + vbCrLf)
                .Append("       z.baptism, z.baptism_date, z.photo_file_name, " + vbCrLf)
                .Append("       COUNT(*) OVER (PARTITION BY 1) cnt " + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("   select ROW_NUMBER() OVER (PARTITION BY x.lvl) row_id, x.lvl, " + vbCrLf)
                .Append("           x.id, x.head_name, x.pid,  " + vbCrLf)
                .Append("           x.korean_name, x.last_name, x.first_name, x.gender, " + vbCrLf)
                .Append("           x.date_of_birth dob, date_part('year', x.age) age, x.cell_phone cell,  " + vbCrLf)
                .Append("           x.home_phone home, x.work_phone as work, x.email, x.rtype, " + vbCrLf)
                .Append("           (select y.address  " + vbCrLf)
                .Append("            from (  " + vbCrLf)
                .Append("               select INITCAP(a.street) as address, ROW_NUMBER() OVER (PARTITION BY a.person_id ORDER BY created_at) row_num  " + vbCrLf)
                .Append("               from public.addresses a where x.pid = a.person_id and a.address_type = 'home'  " + vbCrLf)
                .Append("             ) y  " + vbCrLf)
                .Append("             where y.row_num = 1  " + vbCrLf)
                .Append("           ) address, " + vbCrLf)
                .Append("           (select y.address  " + vbCrLf)
                .Append("            from (  " + vbCrLf)
                .Append("               select INITCAP(a.city) || ' ' || INITCAP(CASE WHEN lower(a.province)='ontario' THEN 'Ont.' ELSE a.province END) || ' ' || UPPER(a.postal_code) as address, ROW_NUMBER() OVER (PARTITION BY a.person_id ORDER BY created_at) row_num  " + vbCrLf)
                .Append("               from public.addresses a where x.pid = a.person_id and a.address_type = 'home'  " + vbCrLf)
                .Append("             ) y  " + vbCrLf)
                .Append("             where y.row_num = 1  " + vbCrLf)
                .Append("           ) address2, " + vbCrLf)
                .Append("       (select name " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where x.pid = m.person_id and type = 'MembershipMilestone' " + vbCrLf)
                .Append("            ) y " + vbCrLf)
                .Append("           where y.row_num = 1 " + vbCrLf)
                .Append("        ) status, " + vbCrLf)
                .Append("       (select effective_date " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where x.pid = m.person_id and type = 'MembershipMilestone' " + vbCrLf)
                .Append("            ) y " + vbCrLf)
                .Append("           where y.row_num = 1 " + vbCrLf)
                .Append("        ) status_date, " + vbCrLf)
                .Append("          (select y.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where x.pid = m.person_id and name = 'infant_baptism' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) y " + vbCrLf)
                .Append("            where y.row_num = 1 " + vbCrLf)
                .Append("           ) infant_baptism, " + vbCrLf)
                .Append("          (select y.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where x.pid = m.person_id and name = 'infant_baptism' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) y " + vbCrLf)
                .Append("            where y.row_num = 1 " + vbCrLf)
                .Append("           ) infant_baptism_date, " + vbCrLf)
                .Append("          (select y.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where x.pid = m.person_id and name = 'confirmation' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) y " + vbCrLf)
                .Append("            where y.row_num = 1 " + vbCrLf)
                .Append("           ) confirmation, " + vbCrLf)
                .Append("          (select y.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where x.pid = m.person_id and name = 'confirmation' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) y " + vbCrLf)
                .Append("            where y.row_num = 1 " + vbCrLf)
                .Append("           ) confirmation_date, " + vbCrLf)
                .Append("          (select y.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where x.pid = m.person_id and name = 'baptism' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) y " + vbCrLf)
                .Append("            where y.row_num = 1 " + vbCrLf)
                .Append("           ) baptism, " + vbCrLf)
                .Append("          (select y.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where x.pid = m.person_id and name = 'baptism' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) y " + vbCrLf)
                .Append("            where y.row_num = 1 " + vbCrLf)
                .Append("           ) baptism_date, " + vbCrLf)
                .Append("          (select y.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where x.pid = m.person_id and type = 'TitleMilestone' " + vbCrLf)
                .Append("            ) y " + vbCrLf)
                .Append("            where y.row_num = 1 " + vbCrLf)
                .Append("           ) title, x.photo_file_name " + vbCrLf)
                .Append("   from ( " + vbCrLf)
                .Append("       select distinct '1-세대주' lvl, p.id, korean_name head_name, p.id pid,  " + vbCrLf)
                .Append("               korean_name, last_name, first_name, gender,  " + vbCrLf)
                .Append("               date_of_birth, cell_phone, home_phone, work_phone, email, age(date_of_birth) age, p.photo_file_name, " + vbCrLf)
                .Append("               '세대주' rtype, ROW_NUMBER() OVER (PARTITION BY p.id) pcnt " + vbCrLf)
                .Append("       from public.people p, public.relationships r " + vbCrLf)
                .Append("       where p.id = r.related_person_id  " + vbCrLf)
                .Append("       and r.relationship_type = '" + relationshipType + "' " + vbCrLf)
                .Append("       and r.related_person_id in (" + personIds + ") " + vbCrLf)
                .Append("       union all " + vbCrLf)
                .Append("       select '2-부양가족' lvl, r.related_person_id id, korean_name head_name, " + vbCrLf)
                .Append("               r.person_id pid, " + vbCrLf)
                .Append("               (select korean_name from public.people where id = r.person_id) korean_name,  " + vbCrLf)
                .Append("               (select last_name from public.people where id = r.person_id) last_name, " + vbCrLf)
                .Append("               (select first_name from public.people where id = r.person_id) first_name, " + vbCrLf)
                .Append("               (select gender from public.people where id = r.person_id) gender, " + vbCrLf)
                .Append("               (select date_of_birth from public.people where id = r.person_id) date_of_birth, " + vbCrLf)
                .Append("               (select cell_phone from public.people where id = r.person_id) cell_phone, " + vbCrLf)
                .Append("               (select home_phone from public.people where id = r.person_id) home_phone, " + vbCrLf)
                .Append("               (select work_phone from public.people where id = r.person_id) work_phone, " + vbCrLf)
                .Append("               (select email from public.people where id = r.person_id) email, " + vbCrLf)
                .Append("               (select age(date_of_birth) from public.people where id = r.person_id) age, " + vbCrLf)
                .Append("               (select photo_file_name from public.people where id = r.person_id) photo_file_name, " + vbCrLf)
                .Append("               CASE WHEN r.relationship_type = 'spouse' THEN '배우자' ELSE '자녀' END rtype, ROW_NUMBER() OVER (PARTITION BY r.person_id) pcnt " + vbCrLf)
                .Append("       from public.people p, public.relationships r " + vbCrLf)
                .Append("       where p.id = r.related_person_id " + vbCrLf)
                .Append("       and r.related_person_id in (" + personIds + ") " + vbCrLf)
                .Append("       ) x " + vbCrLf)
                .Append("       where x.korean_name is not null and x.lvl || x.pcnt != '2-부양가족2' " + vbCrLf)
                .Append("   ) z " + vbCrLf)
                .Append("where z.lvl || z.row_id = '1-세대주1' Or z.lvl = '2-부양가족' " + vbCrLf)
                .Append("order by z.lvl,  CASE WHEN z.rtype='배우자' THEN 1 ELSE 2 END, z.id, z.dob ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                'cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personId, Integer)

                dr = cmd.ExecuteReader

                If dr.HasRows = True Then
                    With getPeopleFindWithFamily
                        .Load(dr)
                        .Columns(0).ColumnName = "lvl"
                        .Columns(1).ColumnName = "id"
                        .Columns(2).ColumnName = "KoreanName"
                        .Columns(3).ColumnName = "LastName"
                        .Columns(4).ColumnName = "FirstName"
                        .Columns(5).ColumnName = "Gender"
                        .Columns(6).ColumnName = "DOB"
                        .Columns(7).ColumnName = "Age"
                        .Columns(8).ColumnName = "Title"
                        .Columns(9).ColumnName = "Cell"
                        .Columns(10).ColumnName = "Home"
                        .Columns(11).ColumnName = "Work"
                        .Columns(12).ColumnName = "Email"
                        .Columns(13).ColumnName = "Type"
                        .Columns(14).ColumnName = "Address"
                        .Columns(15).ColumnName = "City"
                        .Columns(16).ColumnName = "Status"
                        .Columns(17).ColumnName = "StatusDate"
                        .Columns(18).ColumnName = "Infant_Baptism"
                        .Columns(19).ColumnName = "Infant_Baptism_Date"
                        .Columns(20).ColumnName = "Confirmation"
                        .Columns(21).ColumnName = "Confirmation_Date"
                        .Columns(22).ColumnName = "Baptism"
                        .Columns(23).ColumnName = "Baptism_Date"
                        .Columns(24).ColumnName = "PhotoName"
                        .Columns(25).ColumnName = "cnt"
                    End With
                End If

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getPeopleFindWithFamily : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getPeopleFindWithFamily

        End Function

        Public Function getPeopleFindWithoutFamily(ByVal personId As String) As DataTable
            Dim sql As String = String.Empty
            getPeopleFindWithoutFamily = New DataTable
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("personId=[" + personId + "]")

            With sb
                .Append("select '0-단독개인' lvl, id, korean_name, INITCAP(last_name) last_name, INITCAP(first_name) first_name, gender, " + vbCrLf)
                .Append("       date_of_birth as dob, date_part('year', age(date_of_birth)) age, " + vbCrLf)
                .Append("          (select x.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where id = m.person_id and m.person_id = @personId and type = 'TitleMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) title, " + vbCrLf)
                .Append("       cell_phone, home_phone, work_phone, email, '본인' rtype, " + vbCrLf)
                .Append("           (select x.address  " + vbCrLf)
                .Append("            from (  " + vbCrLf)
                .Append("               select INITCAP(a.street) as address, ROW_NUMBER() OVER (PARTITION BY a.person_id ORDER BY created_at) row_num  " + vbCrLf)
                .Append("               from public.addresses a where id = a.person_id and a.person_id = @personId and a.address_type = 'home'  " + vbCrLf)
                .Append("             ) x  " + vbCrLf)
                .Append("             where x.row_num = 1  " + vbCrLf)
                .Append("           ) address, " + vbCrLf)
                .Append("           (select x.address  " + vbCrLf)
                .Append("            from (  " + vbCrLf)
                .Append("               select INITCAP(a.city) || ' ' || INITCAP(CASE WHEN lower(a.province)='ontario' THEN 'Ont.' ELSE a.province END) || ' ' || a.postal_code as address, ROW_NUMBER() OVER (PARTITION BY a.person_id ORDER BY created_at) row_num  " + vbCrLf)
                .Append("               from public.addresses a where id = a.person_id and a.person_id = @personId and a.address_type = 'home'  " + vbCrLf)
                .Append("             ) x  " + vbCrLf)
                .Append("             where x.row_num = 1  " + vbCrLf)
                .Append("           ) address2, " + vbCrLf)
                .Append("       (select x.name " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where id = m.person_id and m.person_id = @personId and type = 'MembershipMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("           where x.row_num = 1 " + vbCrLf)
                .Append("        ) status, " + vbCrLf)
                .Append("       (select x.effective_date " + vbCrLf)
                .Append("            from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where id = m.person_id and m.person_id = @personId and type = 'MembershipMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("           where x.row_num = 1 " + vbCrLf)
                .Append("        ) status_date, " + vbCrLf)
                .Append("          (select x.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where id = m.person_id and m.person_id = @personId and name = 'infant_baptism' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) infant_baptism, " + vbCrLf)
                .Append("          (select x.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where id = m.person_id and m.person_id = @personId and name = 'infant_baptism' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) infant_baptism_date, " + vbCrLf)
                .Append("          (select x.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where id = m.person_id and m.person_id = @personId and name = 'confirmation' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) confirmation, " + vbCrLf)
                .Append("          (select x.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where id = m.person_id and m.person_id = @personId and name = 'confirmation' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) confirmation_date, " + vbCrLf)
                .Append("          (select x.name " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select name, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where id = m.person_id and m.person_id = @personId and name = 'baptism' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) baptism, " + vbCrLf)
                .Append("          (select x.effective_date " + vbCrLf)
                .Append("           from ( " + vbCrLf)
                .Append("               select effective_date, ROW_NUMBER() OVER (PARTITION BY m.person_id ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("               from public.milestones m where id = m.person_id and m.person_id = @personId and name = 'baptism' and type = 'BaptismMilestone' " + vbCrLf)
                .Append("            ) x " + vbCrLf)
                .Append("            where x.row_num = 1 " + vbCrLf)
                .Append("           ) baptism_date, photo_file_name, " + vbCrLf)
                .Append("        COUNT(*) OVER (PARTITION BY 1) cnt " + vbCrLf)
                .Append("from public.people " + vbCrLf)
                .Append("where id = @personId")

            End With

            sql = sb.ToString

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personId, Integer)

                dr = cmd.ExecuteReader

                With getPeopleFindWithoutFamily
                    .Load(dr)
                    .Columns(0).ColumnName = "lvl"
                    .Columns(1).ColumnName = "id"
                    .Columns(2).ColumnName = "KoreanName"
                    .Columns(3).ColumnName = "LastName"
                    .Columns(4).ColumnName = "FirstName"
                    .Columns(5).ColumnName = "Gender"
                    .Columns(6).ColumnName = "DOB"
                    .Columns(7).ColumnName = "Age"
                    .Columns(8).ColumnName = "Title"
                    .Columns(9).ColumnName = "Cell"
                    .Columns(10).ColumnName = "Home"
                    .Columns(11).ColumnName = "Work"
                    .Columns(12).ColumnName = "Email"
                    .Columns(13).ColumnName = "Type"
                    .Columns(14).ColumnName = "Address"
                    .Columns(15).ColumnName = "City"
                    .Columns(16).ColumnName = "Status"
                    .Columns(17).ColumnName = "StatusDate"
                    .Columns(18).ColumnName = "Infant_Baptism"
                    .Columns(19).ColumnName = "Infant_Baptism_Date"
                    .Columns(20).ColumnName = "Confirmation"
                    .Columns(21).ColumnName = "Confirmation_Date"
                    .Columns(22).ColumnName = "Baptism"
                    .Columns(23).ColumnName = "Baptism_Date"
                    .Columns(24).ColumnName = "PhotoName"
                    .Columns(25).ColumnName = "cnt"
                End With

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getPeopleFindWithoutFamily : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getPeopleFindWithoutFamily

        End Function

        Public Function getVisitedFamily(ByVal personIds As String) As List(Of TKPC.Entity.VisitEnt)
            Dim sql As String = String.Empty
            getVisitedFamily = New List(Of TKPC.Entity.VisitEnt)
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("personIds=[" + personIds + "]")

            With sb
                .Append("select x.korean_name, TO_CHAR(x.date, 'YYYY-MM-DD') date, x.note " + vbCrLf)
                .Append("from " + vbCrLf)
                .Append("( " + vbCrLf)
                .Append("   select p.korean_name, date, note, ROW_NUMBER() OVER (PARTITION BY p.id) rank " + vbCrLf)
                .Append("   from public.pastoral_visits v, public.people p " + vbCrLf)
                .Append("   where p.id = v.person_id " + vbCrLf)
                .Append("   and v.person_id in (" + personIds + ")" + vbCrLf)
                .Append(") x " + vbCrLf)
                .Append("where x.rank <= 2 " + vbCrLf)
                .Append("order by date desc ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                dr = cmd.ExecuteReader
                While dr.Read()
                    Dim visitEnt As TKPC.Entity.VisitEnt = New TKPC.Entity.VisitEnt
                    With visitEnt
                        .visitedKoreanName = dr("korean_name").ToString
                        .visitDate = dr("date").ToString
                        .note = dr("note").ToString
                    End With
                    getVisitedFamily.Add(visitEnt)
                End While

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getVisitedFamily : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getVisitedFamily
        End Function

        Public Function restoreUpdatedSpousePositions() As String
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder
            restoreUpdatedSpousePositions = "update public.relationships set related_person_id = person_id, person_id = related_person_id where id in ("

            With sb
                .Append("select x.id " + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("       select id, related_person_id,  " + vbCrLf)
                .Append("          (select korean_name from public.people p where r.related_person_id = p.id) korean_name1, " + vbCrLf)
                .Append("          (select gender from public.people p where r.related_person_id = p.id) gender1, " + vbCrLf)
                .Append("          person_id,  " + vbCrLf)
                .Append("          (select korean_name from public.people p where r.person_id = p.id) korean_name2, " + vbCrLf)
                .Append("          (select gender from public.people p where r.person_id = p.id) gender2, " + vbCrLf)
                .Append("          r.relationship_type " + vbCrLf)
                .Append("       from public.relationships r " + vbCrLf)
                .Append(") x " + vbCrLf)
                .Append("where x.relationship_type = 'spouse' and x.gender1 = 'F' and x.gender2 = 'M' ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                dr = cmd.ExecuteReader
                While dr.Read()
                    restoreUpdatedSpousePositions += dr("id").ToString + ","
                End While

                restoreUpdatedSpousePositions = restoreUpdatedSpousePositions.Substring(0, restoreUpdatedSpousePositions.Length - 1) + ")"

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.restoreUpdatedSpousePositions : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
        End Function

        Public Function updateSpousePosition() As Integer
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim ta As NpgsqlTransaction = Nothing
            Dim sb As StringBuilder = New StringBuilder
            updateSpousePosition = 0

            With sb
                .Append("update public.relationships set related_person_id = person_id, person_id = related_person_id ")
                .Append("where id in (")
                .Append("   select x.id " + vbCrLf)
                .Append("   from ( " + vbCrLf)
                .Append("       select id, related_person_id,  " + vbCrLf)
                .Append("          (select korean_name from public.people p where r.related_person_id = p.id) korean_name1, " + vbCrLf)
                .Append("          (select gender from public.people p where r.related_person_id = p.id) gender1, " + vbCrLf)
                .Append("          person_id,  " + vbCrLf)
                .Append("          (select korean_name from public.people p where r.person_id = p.id) korean_name2, " + vbCrLf)
                .Append("          (select gender from public.people p where r.person_id = p.id) gender2, " + vbCrLf)
                .Append("          r.relationship_type " + vbCrLf)
                .Append("       from public.relationships r " + vbCrLf)
                .Append("   ) x " + vbCrLf)
                .Append("   where x.relationship_type = 'spouse' and x.gender1 = 'F' and x.gender2 = 'M') ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                ta = cn.BeginTransaction()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0
                cmd.Transaction = ta

                updateSpousePosition = cmd.ExecuteNonQuery()
                ta.Commit()

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.updateSpousePosition : " + ex.Message)
                ta.Rollback()
                Logger.LogError("Update was rollbacked. " + ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(cmd, cn, ta)
                sb = Nothing
            End Try
        End Function

        Public Function updateFamilyType(ByVal id As Integer) As Integer
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim ta As NpgsqlTransaction = Nothing
            Dim sb As StringBuilder = New StringBuilder
            updateFamilyType = 0

            With sb
                .Append("update public.relationships set related_person_id = person_id, person_id = related_person_id, relationship_type = 'parent' ")
                .Append("where id = @id ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                ta = cn.BeginTransaction()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@id", NpgsqlTypes.NpgsqlDbType.Integer)).Value = id
                cmd.Transaction = ta

                updateFamilyType = cmd.ExecuteNonQuery()
                ta.Commit()

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.updateFamilyType : " + ex.Message)
                ta.Rollback()
                Logger.LogError("Update was rollbacked. " + ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(cmd, cn, ta)
                sb = Nothing
            End Try
        End Function

        Public Function getSwitchedFamilyHeads() As DataTable
            Dim sql As String = String.Empty
            getSwitchedFamilyHeads = New DataTable
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            With sb
                .Append("select x.korean_name1, x.gender1, x.korean_name2, x.gender2 " + vbCrLf)
                .Append("from (  " + vbCrLf)
                .Append("   select id, related_person_id,  " + vbCrLf)
                .Append("         (select korean_name from public.people p where r.related_person_id = p.id) korean_name1,  " + vbCrLf)
                .Append("         (select gender from public.people p where r.related_person_id = p.id) gender1, " + vbCrLf)
                .Append("         person_id,  " + vbCrLf)
                .Append("         (select korean_name from public.people p where r.person_id = p.id) korean_name2,  " + vbCrLf)
                .Append("         (select gender from public.people p where r.person_id = p.id) gender2,  " + vbCrLf)
                .Append("   r.relationship_type  " + vbCrLf)
                .Append("   from public.relationships r  " + vbCrLf)
                .Append(") x  " + vbCrLf)
                .Append("where x.relationship_type = 'spouse' and x.gender1 = 'F' and x.gender2 = 'M'")
            End With

            sql = sb.ToString

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                dr = cmd.ExecuteReader

                With getSwitchedFamilyHeads
                    .Load(dr)
                    .Columns(0).ColumnName = "세대주 이름"
                    .Columns(1).ColumnName = "세대주 성별"
                    .Columns(2).ColumnName = "배우자 이름"
                    .Columns(3).ColumnName = "배우자 성별"
                End With

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getSwitchedFamilyHeads : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getSwitchedFamilyHeads

        End Function

        Public Function getSwitchedFamilyType() As DataTable
            Dim sql As String = String.Empty
            getSwitchedFamilyType = New DataTable
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            With sb
                .Append("select x.id, x.korean_name1, x.gender1, x.korean_name2, x.gender2 " + vbCrLf)
                .Append("from (  " + vbCrLf)
                .Append("   select id, related_person_id,  " + vbCrLf)
                .Append("         (select korean_name from public.people p where r.related_person_id = p.id) korean_name1,  " + vbCrLf)
                .Append("         (select gender from public.people p where r.related_person_id = p.id) gender1, " + vbCrLf)
                .Append("         person_id,  " + vbCrLf)
                .Append("         (select korean_name from public.people p where r.person_id = p.id) korean_name2,  " + vbCrLf)
                .Append("         (select gender from public.people p where r.person_id = p.id) gender2,  " + vbCrLf)
                .Append("   r.relationship_type  " + vbCrLf)
                .Append("   from public.relationships r  " + vbCrLf)
                .Append("   where r.relationship_type = 'child' " + vbCrLf)
                .Append(") x  " + vbCrLf)
                .Append("where x.relationship_type = 'child'")
            End With

            sql = sb.ToString

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                dr = cmd.ExecuteReader

                With getSwitchedFamilyType
                    .Load(dr)
                    .Columns(0).ColumnName = "ID"
                    .Columns(1).ColumnName = "자녀이름"
                    .Columns(2).ColumnName = "자녀성별"
                    .Columns(3).ColumnName = "부모이름"
                    .Columns(4).ColumnName = "부모성별"
                End With

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getSwitchedFamilyHeads : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getSwitchedFamilyType

        End Function

        Function isMileStoneThere(ByVal personId As String, ByVal milestoneType As String, ByVal milestoneName As String, ByVal effectiveDate As String) As Boolean
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder
            isMileStoneThere = False

            Logger.LogInfo("personId=[" + personId + "], milestoneType=[" + milestoneType + "], milestoneName=[" + milestoneName + "]")

            With sb
                .Append("select person_id, type, name, effective_date " + vbCrLf)
                .Append("from public.milestones " + vbCrLf)
                .Append("where person_id =@personId and type =@milestoneType and name =@milestoneName")
            End With

            If milestoneName = m_Constant.MILESTONE_NAME_SEORI Then
                sb.Append(" and effective_date=@effectiveDate ")
            End If

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personId, Integer)
                cmd.Parameters.Add(New NpgsqlParameter("@milestoneType", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = milestoneType
                cmd.Parameters.Add(New NpgsqlParameter("@milestoneName", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = milestoneName

                If milestoneName = m_Constant.MILESTONE_NAME_SEORI Then
                    cmd.Parameters.Add(New NpgsqlParameter("@effectiveDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(effectiveDate, Date)
                End If

                dr = cmd.ExecuteReader
                While dr.Read()
                    isMileStoneThere = True
                End While

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.isMileStoneThere : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
        End Function

        Function updateMilestone(ByVal personId As String, ByVal milestoneType As String, ByVal milestoneName As String, ByVal effectiveDate As String) As Integer
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim ta As NpgsqlTransaction = Nothing
            Dim sb As StringBuilder = New StringBuilder
            updateMilestone = 0

            Logger.LogInfo("personId=[" + personId + "], milestoneType=[" + milestoneType + "], milestoneName=[" + milestoneName + "], effectiveDate=[" + effectiveDate + "]")

            With sb
                .Append("update public.milestones " + vbCrLf)
                .Append("set effective_date =@effectiveDate" + vbCrLf)
                .Append("where person_id =@personId and type =@milestoneType and name =@milestoneName")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                ta = cn.BeginTransaction()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0
                cmd.Transaction = ta

                cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personId, Integer)
                cmd.Parameters.Add(New NpgsqlParameter("@milestoneType", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = milestoneType
                cmd.Parameters.Add(New NpgsqlParameter("@milestoneName", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = milestoneName
                cmd.Parameters.Add(New NpgsqlParameter("@effectiveDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(effectiveDate, Date)

                updateMilestone = cmd.ExecuteNonQuery()
                ta.Commit()

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.UpdateMilestone : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
                ta.Rollback()
            Finally
                m_DBUtil.closeConnection(cmd, cn, ta)
                sb = Nothing
            End Try
            Return updateMilestone
        End Function

        Function insertComment(ByVal personId As String, ByVal comment As String) As Integer
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim ta As NpgsqlTransaction = Nothing
            Dim sb As StringBuilder = New StringBuilder
            insertComment = 0

            Logger.LogInfo("personId=[" + personId + "]")

            With sb
                .Append("insert into public.active_admin_comments (namespace, body, resource_type, resource_id, author_type, author_id, created_at, updated_at) " + vbCrLf)
                .Append("values ('admin', @comment, 'Person', @personId, 'AdminUser', @authorId, now(), now()) " + vbCrLf)
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                ta = cn.BeginTransaction()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0
                cmd.Transaction = ta

                cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personId, Integer)
                cmd.Parameters.Add(New NpgsqlParameter("@comment", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = comment
                cmd.Parameters.Add(New NpgsqlParameter("@authorId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(m_Global.loginUserAdminId, Integer)

                insertComment = cmd.ExecuteNonQuery()
                ta.Commit()

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.UpdateMilestone : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
                ta.Rollback()
            Finally
                m_DBUtil.closeConnection(cmd, cn, ta)
                sb = Nothing
            End Try
            Return insertComment
        End Function

        Function deleteRelationship(ByVal personId As String) As Integer
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim ta As NpgsqlTransaction = Nothing
            Dim sb As StringBuilder = New StringBuilder
            deleteRelationship = 0

            Logger.LogInfo("personId=[" + personId + "]")

            With sb
                .Append("delete from public.relationships " + vbCrLf)
                .Append("where person_id = @personId or related_person_id = @personId ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                ta = cn.BeginTransaction()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0
                cmd.Transaction = ta

                cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personId, Integer)

                deleteRelationship = cmd.ExecuteNonQuery()
                ta.Commit()

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.UpdateMilestone : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
                ta.Rollback()
            Finally
                m_DBUtil.closeConnection(cmd, cn, ta)
                sb = Nothing
            End Try
            Return deleteRelationship
        End Function

        Function insertMilestone(ByVal personId As String, ByVal milestoneType As String, ByVal milestoneName As String, ByVal effectiveDate As String) As Integer
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim ta As NpgsqlTransaction = Nothing
            Dim sb As StringBuilder = New StringBuilder
            insertMilestone = 0

            Logger.LogInfo("personId=[" + personId + "], milestoneType=[" + milestoneType + "], milestoneName=[" + milestoneName + "], effectiveDate=[" + effectiveDate + "]")

            With sb
                .Append("insert into public.milestones (person_id, type, name, effective_date, created_at, updated_at) " + vbCrLf)
                .Append("values (@personId, @milestoneType, @milestoneName, @effectiveDate, now(), now()) " + vbCrLf)
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                ta = cn.BeginTransaction()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0
                cmd.Transaction = ta

                cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personId, Integer)
                cmd.Parameters.Add(New NpgsqlParameter("@milestoneType", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = milestoneType
                cmd.Parameters.Add(New NpgsqlParameter("@milestoneName", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = milestoneName
                cmd.Parameters.Add(New NpgsqlParameter("@effectiveDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(effectiveDate, Date)

                insertMilestone = cmd.ExecuteNonQuery()
                ta.Commit()

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.UpdateMilestone : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
                ta.Rollback()
            Finally
                m_DBUtil.closeConnection(cmd, cn, ta)
                sb = Nothing
            End Try
            Return insertMilestone
        End Function

        Public Function getTitleYears(ByVal personId As String, ByVal titleName As String) As String
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder
            getTitleYears = ""

            With sb
                .Append("select string_agg(cast(date_part('year', effective_date) as varchar), ',' ORDER BY date_part('year', effective_date)) as years" + vbCrLf)
                .Append("from public.milestones m " + vbCrLf)
                .Append("where type = 'TitleMilestone' and name = @titleName " + vbCrLf)
                .Append("and person_id = @personId ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personId, Integer)
                cmd.Parameters.Add(New NpgsqlParameter("@titleName", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = titleName.ToLower

                dr = cmd.ExecuteReader
                While dr.Read()
                    getTitleYears = dr("years").ToString
                End While

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getTitleYears : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try

            Return getTitleYears
        End Function

        Public Function getMilestones(ByVal personIds As String) As List(Of TKPC.Entity.MilestoneEnt)
            Dim sql As String = String.Empty
            getMilestones = New List(Of TKPC.Entity.MilestoneEnt)
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("personIds=[" + personIds + "]")

            With sb
                .Append("select p.id, p.korean_name, to_char(effective_date, 'yyyy-mm-dd') effective_date, type, name " + vbCrLf)
                .Append("from public.milestones m, public.people p " + vbCrLf)
                .Append("where p.id = m.person_id " + vbCrLf)
                .Append("and person_id in (" + personIds + ") " + vbCrLf)
                .Append("order by p.korean_name, effective_date ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                dr = cmd.ExecuteReader
                While dr.Read()
                    Dim milestoneEnt As TKPC.Entity.MilestoneEnt = New TKPC.Entity.MilestoneEnt
                    With milestoneEnt
                        .personId = dr("id").ToString
                        .koreanName = dr("korean_name").ToString
                        .type = dr("type").ToString
                        .name = dr("name").ToString
                        .effectiveDate = dr("effective_date").ToString
                    End With
                    getMilestones.Add(milestoneEnt)
                End While

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getMilestones : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getMilestones
        End Function

        Public Function getAddresses(ByVal personIds As String) As List(Of TKPC.Entity.PersonEnt)
            Dim sql As String = String.Empty
            getAddresses = New List(Of Entity.PersonEnt)
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("personIds=[" + personIds + "]")

            With sb
                .Append("select p.id, p.korean_name, age(p.date_of_birth) age, a.address_type, INITCAP(a.street) || ', ' || INITCAP(a.city) || ', ' || INITCAP(a.province) || ', ' || a.postal_code as address" + vbCrLf)
                .Append("from public.people p, public.addresses a " + vbCrLf)
                .Append("where p.id = a.person_id " + vbCrLf)
                .Append("and p.id in (" + personIds + ") " + vbCrLf)
                .Append("order by 3 desc")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                dr = cmd.ExecuteReader
                While dr.Read()
                    Dim personEnt As TKPC.Entity.PersonEnt = New TKPC.Entity.PersonEnt
                    With personEnt
                        .personId = dr("id").ToString
                        .koreanName = dr("korean_name").ToString
                        .addressType = dr("address_type").ToString
                        .address = dr("address").ToString
                    End With
                    getAddresses.Add(personEnt)
                End While

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getAddresses : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getAddresses
        End Function

        Public Function getComments(ByVal personIds As String) As List(Of TKPC.Entity.CommentEnt)
            Dim sql As String = String.Empty
            getComments = New List(Of Entity.CommentEnt)
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("personIds=[" + personIds + "]")

            With sb
                .Append("select p.id, p.korean_name, a.name, c.body " + vbCrLf)
                .Append("from active_admin_comments c, public.people p, admin_users a " + vbCrLf)
                .Append("where resource_id = p.id " + vbCrLf)
                .Append("and a.id = c.author_id " + vbCrLf)
                .Append("and resource_type = 'Person' " + vbCrLf)
                .Append("and p.id in (" + personIds + ") " + vbCrLf)
                .Append("order by p.korean_name")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                dr = cmd.ExecuteReader
                While dr.Read()
                    Dim commentEnt As TKPC.Entity.CommentEnt = New TKPC.Entity.CommentEnt
                    With commentEnt
                        .personId = dr("id").ToString
                        .personName = dr("korean_name").ToString
                        .authorName = dr("name").ToString
                        .body = dr("body").ToString
                    End With
                    getComments.Add(commentEnt)
                End While

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at PersonDAL.getAddresses : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getComments
        End Function

    End Class
End Namespace