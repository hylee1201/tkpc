Imports Npgsql
Imports System.Data
Imports System.Text
Imports System.Collections.Generic
Imports System.Configuration

Namespace TKPC.DAL
    Public Class OfferingDAL
        Private Shared Instance As OfferingDAL

        Protected Sub New()
        End Sub

        'To make this DAL singleton to use memory efficiently
        Public Shared Function getInstance() As OfferingDAL
            ' initialize if not already done
            If Instance Is Nothing Then
                Instance = New OfferingDAL
            End If
            ' return the initialized instance of the Singleton Class
            Return Instance
        End Function 'Instance

        Public Function getOfferingByTypeForSelectedDate(selectedStartDate As String, selectedEndDate As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingByTypeForSelectedDate = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("selectedStartDate=[" + selectedStartDate + "], selectedEndDate=[" + selectedEndDate + "]")

            With sb
                .Append("select a.date, a.offering_type, to_char(a.amount, 'FM9,999,999,999.00') amount, a.offering_numbers ")
                .Append("from (")
                .Append("   select z.lvl, z.date, z.offering_type, z.amount, z.offering_numbers ")
                .Append("   from ( ")
                .Append("       select t.date, ")
                .Append("               MAX(CASE WHEN t.offering_type = 'weekly' THEN 1 ")
                .Append("               WHEN t.offering_type = 'tithe' THEN 2 ")
                .Append("               WHEN t.offering_type = 'missions' THEN 3 ")
                .Append("               WHEN t.offering_type = 'thanksgiving' THEN 4 ")
                .Append("               Else 5 End) lvl, ")
                .Append("               Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("               SUM(t.amount) amount, string_agg(cast(e.offering_number as varchar), ',' ORDER BY e.offering_number) offering_numbers  ")
                .Append("       from public.offering t, public.offering_envelopes e ")
                .Append("       where t.offering_envelope_id = e.id ")
                .Append("       and t.date between @startDate and @endDate ")
                .Append("       group by t.date, CASE WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END ")
                .Append("   ) z ")
                .Append("   union all ")
                .Append("   select null as lvl, t.date as date, 'Total', SUM(amount), null as offering_numbers ")
                .Append("   from public.offering t, public.offering_envelopes e ")
                .Append("   where t.offering_envelope_id = e.id  ")
                .Append("   and t.date between @startDate and @endDate ")
                .Append("   group by t.date ")
                .Append(") a ")
                .Append("order by a.date, a.lvl, a.offering_type, a.amount desc ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(selectedStartDate, Date)
                cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(selectedEndDate, Date)

                da.SelectCommand = cmd
                da.Fill(getOfferingByTypeForSelectedDate, "getOfferingByTypeForSelectedDate")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingByTypeForSelectedDate : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingByTypeForSelectedDate

        End Function

        'This function is to get offering receipts
        Public Function getOfferingReceipt(ByVal flag As Integer, ByVal startYearOrDate As String, ByVal endYearOrDate As String, ByVal personIds As String, ByVal street As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingReceipt = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("flag=[" + CType(flag, String) + "], startYearOrDate=[" + startYearOrDate + "], endYearOrDate=[" + endYearOrDate + "], persondIds=[" + personIds + "], street=[" + street + "]")

            With sb
                .Append("select x.offering_number, x.year, x.person_id, x.korean_name, UPPER(x.last_name) as last_name, UPPER(x.first_name) first_name, " + vbCrLf)
                .Append("UPPER(a.street) as street, UPPER(a.city) as city, UPPER(a.province) as province, UPPER(a.postal_code) as postal_code, a.country, to_char(x.total, 'FM9,999,999,999.00') total " + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("       select e.offering_number, e.year, e.person_id, p.korean_name, p.last_name, p.first_name, SUM(t.amount) total " + vbCrLf)
                .Append("       from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("       where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("       and e.person_id = p.id " + vbCrLf)

                If String.IsNullOrEmpty(personIds) = False Then
                    .Append("       and e.person_id in (" + personIds + ") " + vbCrLf)
                End If

                If flag = 0 Then 'date range
                    .Append("       and t.date between @startYearOrDate and @endYearOrDate " + vbCrLf)
                Else 'year range
                    .Append("       and e.year between @startYearOrDate and @endYearOrDate " + vbCrLf)
                End If

                .Append("       group by p.korean_name, p.last_name, p.first_name, e.offering_number, e.year, e.person_id " + vbCrLf)
                .Append(") x, addresses a " + vbCrLf)
                .Append("where x.person_id = a.person_id " + vbCrLf)

                If flag = 1 Then
                    If String.IsNullOrEmpty(street) = False Then
                        .Append("and LOWER(a.street) = @street ")
                    End If
                End If

                    .Append("order by x.offering_number, x.year, x.korean_name ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                If String.IsNullOrEmpty(personIds) = False Then
                    cmd.Parameters.Add(New NpgsqlParameter("@personIds", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = CType(personIds, String)
                End If

                If flag = 0 Then 'date range
                    cmd.Parameters.Add(New NpgsqlParameter("@startYearOrDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(startYearOrDate, Date)
                    cmd.Parameters.Add(New NpgsqlParameter("@endYearOrDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(endYearOrDate, Date)
                Else 'year range
                    cmd.Parameters.Add(New NpgsqlParameter("@startYearOrDate", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(startYearOrDate, Integer)
                    cmd.Parameters.Add(New NpgsqlParameter("@endYearOrDate", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(endYearOrDate, Integer)

                    If String.IsNullOrEmpty(street) = False Then
                        cmd.Parameters.Add(New NpgsqlParameter("@street", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = street.ToLower
                    End If
                End If
                da.SelectCommand = cmd
                da.Fill(getOfferingReceipt, "getOfferingReceipt")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingReceipt : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingReceipt
        End Function

        'This function is to get offering receipts
        Public Function getOfferingByMonth(ByVal startYear As String, ByVal endYear As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingByMonth = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("startYear=[" + startYear + "], endYear=[" + endYear + "]")

            With sb
                .Append("select z.year, z.offering_type, to_char(z.jan, 'FM9,999,999,999.00') as January, to_char(z.feb, 'FM9,999,999,999.00') as February, to_char(z.mar, 'FM9,999,999,999.00') as March, to_char(z.apr, 'FM9,999,999,999.00') as April, to_char(z.may, 'FM9,999,999,999.00') as May, ")
                .Append("       to_char(z.jun, 'FM9,999,999,999.00') as Jun, to_char(z.jul, 'FM9,999,999,999.00') as July, to_char(z.aug, 'FM9,999,999,999.00') as August, to_char(z.sept, 'FM9,999,999,999.00') as September, to_char(z.nov, 'FM9,999,999,999.00') as November, to_char(z.decm, 'FM9,999,999,999.00') as December, z.total ")
                .Append("from ( ")
                .Append("       select x.year, x.type_lvl, x.offering_type, SUM(x.Jan) jan, SUM(x.Feb) feb, SUM(x.Mar) mar, SUM(x.Apr) apr, SUM(x.May) may, ")
                .Append("              SUM(x.Jun) jun, SUM(x.Jul) jul, SUM(x.Aug) aug, SUM(x.Sept) sept, SUM(x.Nov) nov, SUM(x.Dec) decm, ")
                .Append("              SUM(x.Jan+x.Feb+x.Mar+x.Apr+x.May+x.Jun+x.Jul+x.Aug+x.Sept+x.Nov+x.Dec) total ")
                .Append("       from ( ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl,  ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  SUM(t.amount) as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jan' and e.year between @startYear and @endYear ")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, SUM(t.amount) as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Feb' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason  ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl,  ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, SUM(t.amount) as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Mar' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, SUM(t.amount) as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Apr' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, SUM(t.amount) as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'May' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  SUM(t.amount) as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jun' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, SUM(t.amount) as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jul' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, SUM(t.amount) as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Aug' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, SUM(t.amount) as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Sep' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, SUM(t.amount) as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Oct' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, SUM(t.amount) as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Nov' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, SUM(t.amount) as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Dec' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("       ) x ")
                .Append("       group by x.year, x.type_lvl, x.offering_type ")
                .Append("       union all ")
                .Append("       select x.year, null, 'Total', SUM(x.Jan) jan, SUM(x.Feb) feb, SUM(x.Mar) mar, SUM(x.Apr) apr, SUM(x.May) may, ")
                .Append("              SUM(x.Jun) jun, SUM(x.Jul) jul, SUM(x.Aug) aug, SUM(x.Sept) sept, SUM(x.Nov) nov, SUM(x.Dec) decm, ")
                .Append("              SUM(x.Jan+x.Feb+x.Mar+x.Apr+x.May+x.Jun+x.Jul+x.Aug+x.Sept+x.Nov+x.Dec) total ")
                .Append("       from ( ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl,  ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  SUM(t.amount) as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jan' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, SUM(t.amount) as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Feb' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason  ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl,  ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, SUM(t.amount) as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Mar' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, SUM(t.amount) as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Apr' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, SUM(t.amount) as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'May' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  SUM(t.amount) as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jun' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, SUM(t.amount) as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jul' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, SUM(t.amount) as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Aug' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, SUM(t.amount) as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Sep' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, SUM(t.amount) as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Oct' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, SUM(t.amount) as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Nov' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, SUM(t.amount) as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Dec' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("       ) x ")
                .Append("       group by x.year ")
                .Append(") z ")
                .Append("order by z.year, z.type_lvl")

            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@startYear", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(startYear, Integer)
                cmd.Parameters.Add(New NpgsqlParameter("@endYear", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(endYear, Integer)

                da.SelectCommand = cmd
                da.Fill(getOfferingByMonth, "getOfferingByMonth")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingByMonth : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingByMonth
        End Function

        'This function is to get offering receipts
        Public Function getOfferingByMonthWithoutType(ByVal startYear As String, ByVal endYear As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingByMonthWithoutType = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("startYear=[" + startYear + "], endYear=[" + endYear + "]")

            With sb
                .Append("select CASE WHEN z.offering_type = 'Total' THEN 'Total' ELSE CAST(z.year as VARCHAR) END as year, to_char(z.jan, 'FM9,999,999,999.00') as January, to_char(z.feb, 'FM9,999,999,999.00') as February, to_char(z.mar, 'FM9,999,999,999.00') as March, to_char(z.apr, 'FM9,999,999,999.00') as April, to_char(z.may, 'FM9,999,999,999.00') as May, ")
                .Append("       to_char(z.jun, 'FM9,999,999,999.00') as Jun, to_char(z.jul, 'FM9,999,999,999.00') as July, to_char(z.aug, 'FM9,999,999,999.00') as August, to_char(z.sept, 'FM9,999,999,999.00') as September, to_char(z.nov, 'FM9,999,999,999.00') as November, to_char(z.decm, 'FM9,999,999,999.00') as December, z.total ")
                .Append("from ( ")
                .Append("       select x.year, x.type_lvl, x.offering_type, SUM(x.Jan) jan, SUM(x.Feb) feb, SUM(x.Mar) mar, SUM(x.Apr) apr, SUM(x.May) may, ")
                .Append("              SUM(x.Jun) jun, SUM(x.Jul) jul, SUM(x.Aug) aug, SUM(x.Sept) sept, SUM(x.Nov) nov, SUM(x.Dec) decm, ")
                .Append("              SUM(x.Jan+x.Feb+x.Mar+x.Apr+x.May+x.Jun+x.Jul+x.Aug+x.Sept+x.Nov+x.Dec) total ")
                .Append("       from ( ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl,  ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  SUM(t.amount) as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jan' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, SUM(t.amount) as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Feb' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason  ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl,  ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, SUM(t.amount) as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Mar' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, SUM(t.amount) as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Apr' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, SUM(t.amount) as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'May' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  SUM(t.amount) as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jun' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, SUM(t.amount) as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jul' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, SUM(t.amount) as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Aug' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, SUM(t.amount) as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Sep' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, SUM(t.amount) as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Oct' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, SUM(t.amount) as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Nov' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, SUM(t.amount) as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Dec' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("       ) x ")
                .Append("       group by x.year, x.type_lvl, x.offering_type ")
                .Append("       union all ")
                .Append("       select x.year, null, 'Total', SUM(x.Jan) jan, SUM(x.Feb) feb, SUM(x.Mar) mar, SUM(x.Apr) apr, SUM(x.May) may, ")
                .Append("              SUM(x.Jun) jun, SUM(x.Jul) jul, SUM(x.Aug) aug, SUM(x.Sept) sept, SUM(x.Nov) nov, SUM(x.Dec) decm, ")
                .Append("              SUM(x.Jan+x.Feb+x.Mar+x.Apr+x.May+x.Jun+x.Jul+x.Aug+x.Sept+x.Nov+x.Dec) total ")
                .Append("       from ( ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl,  ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  SUM(t.amount) as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jan' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, SUM(t.amount) as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Feb' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason  ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl,  ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, SUM(t.amount) as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Mar' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, SUM(t.amount) as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Apr' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, SUM(t.amount) as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'May' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  SUM(t.amount) as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jun' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, SUM(t.amount) as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jul' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, SUM(t.amount) as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Aug' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, SUM(t.amount) as Sept, 0 as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Sep' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, SUM(t.amount) as Oct, 0 as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Oct' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, SUM(t.amount) as Nov, 0 as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Nov' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("           union all ")
                .Append("           select '1' lvl, e.year, ")
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 ")
                .Append("                  WHEN offering_type = 'tithe' THEN 2 ")
                .Append("                  WHEN offering_type = 'missions' THEN 3 ")
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 ")
                .Append("                  Else 5 End type_lvl, ")
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type, ")
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, ")
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, SUM(t.amount) as Dec ")
                .Append("           from public.offering t, public.offering_envelopes e ")
                .Append("           where t.offering_envelope_id = e.id ")
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Dec' and e.year between @startYear and @endYear")
                .Append("           group by e.year, t.offering_type, t.other_reason ")
                .Append("       ) x ")
                .Append("       group by x.year ")
                .Append(") z ")
                .Append("order by z.year, z.type_lvl")

            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@startYear", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(startYear, Integer)
                cmd.Parameters.Add(New NpgsqlParameter("@endYear", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(endYear, Integer)

                da.SelectCommand = cmd
                da.Fill(getOfferingByMonthWithoutType, "getOfferingByMonthWithoutType")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingByMonthWithoutType : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingByMonthWithoutType
        End Function

        Public Function getOfferingInfoByPersonIdOrOfferingId(ByRef flag As Integer, ByRef personIds As String, ByRef offeringNumber As String, ByRef selectedYear As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingInfoByPersonIdOrOfferingId = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder
            Dim paramStr As String = ""

            Logger.LogInfo("flag=[" + CType(flag, String) + "], personIds=[" + personIds + "], offeringNumber=[" + offeringNumber + "], selectedYear=[" + selectedYear + "]")

            If flag = 0 Then '이름으로 찾기
                With sb
                    .Append("select z.offering_number, z.year, z.korean_name, z.cell_phone, z.home_phone, z.work_phone, z.address, z.person_id, z.cnt " + vbCrLf)
                    .Append("from ( " + vbCrLf)
                    .Append("   select x.offering_number, x.year, x.korean_name, x.person_id, x.address, x.cell_phone, x.home_phone, x.work_phone, " + vbCrLf)
                    .Append("          COUNT(*) OVER (PARTITION BY x.offering_number ORDER BY x.offering_number desc) cnt " + vbCrLf)
                    .Append("   from ( " + vbCrLf)
                    .Append("       select e.offering_number, e.year, p.korean_name, e.person_id,  " + vbCrLf)
                    .Append("           (select x.address " + vbCrLf)
                    .Append("            from ( " + vbCrLf)
                    .Append("               select INITCAP(a.street) || ' ' || INITCAP(a.city) || ' ' || INITCAP(a.province) || ' ' || a.postal_code as address, " + vbCrLf)
                    .Append("                      ROW_NUMBER() OVER (PARTITION BY a.person_id ORDER BY created_at) row_num " + vbCrLf)
                    .Append("               from public.addresses a where p.id = a.person_id and a.address_type = 'home' " + vbCrLf)
                    .Append("           ) x " + vbCrLf)
                    .Append("           where x.row_num = 1 " + vbCrLf)
                    .Append("       ) address, p.cell_phone, p.home_phone, p.work_phone " + vbCrLf)
                    .Append("   from public.people p " + vbCrLf)
                    .Append("   right outer join public.offering_envelopes e on e.person_id = p.id " + vbCrLf)
                    .Append("   where e.year = @selectedYear " + vbCrLf)

                    If String.IsNullOrEmpty(personIds) = False Then
                        '.Append("and e.person_id = ANY(@name) ")

                        .Append("   and e.person_id in (" + personIds + ") " + vbCrLf)
                    Else
                        Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
                        Dim max As Integer = CType(ConfigurationManager.AppSettings(m_Constant.MAX_OFFERING_NUMBER).ToString(), Integer)

                        .Append("   union all " + m_DBUtil.getOfferingNumbersSQL(max, selectedYear))
                    End If
                    .Append("   ) x " + vbCrLf)
                    .Append(") z " + vbCrLf)
                    .Append("where  (z.cnt >= 2 and person_id is not null) or z.cnt = 1 " + vbCrLf)
                    .Append("order by z.offering_number, z.korean_name COLLATE " + ControlChars.Quote + "C" + ControlChars.Quote)
                End With
            Else '헌금봉투로 찾기
                With sb
                    .Append("select z.offering_number, z.year, z.korean_name, z.cell_phone, z.home_phone, z.work_phone, z.address, z.person_id, z.cnt " + vbCrLf)
                    .Append("from ( " + vbCrLf)
                    .Append("   select x.offering_number, x.year, x.korean_name, x.person_id, x.address, " + vbCrLf)
                    .Append("   from ( " + vbCrLf)
                    .Append("       select e.offering_number, e.year, p.korean_name, e.person_id,  " + vbCrLf)
                    .Append("             (select x.address " + vbCrLf)
                    .Append("              from ( " + vbCrLf)
                    .Append("                   select INITCAP(a.street) || ' ' || INITCAP(a.city) || ' ' || INITCAP(a.province) || ' ' || a.postal_code as address, " + vbCrLf)
                    .Append("                          ROW_NUMBER() OVER (PARTITION BY a.person_id ORDER BY created_at) row_num " + vbCrLf)
                    .Append("                   from public.addresses a where p.id = a.person_id and a.address_type = 'home' " + vbCrLf)
                    .Append("             ) x " + vbCrLf)
                    .Append("             where x.row_num = 1 " + vbCrLf)
                    .Append("             ) address " + vbCrLf)
                    .Append("       from public.people p " + vbCrLf)
                    .Append("       right outer join public.offering_envelopes e on e.person_id = p.id " + vbCrLf)
                    .Append("       where e.year = @selectedYear " + vbCrLf)

                    If String.IsNullOrEmpty(offeringNumber) = False Then
                        .Append("   and e.offering_number = @name " + vbCrLf)
                    Else
                        Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
                        Dim max As Integer = CType(ConfigurationManager.AppSettings(m_Constant.MAX_OFFERING_NUMBER).ToString(), Integer)

                        .Append("   union all " + m_DBUtil.getOfferingNumbersSQL(max, selectedYear))
                    End If
                    .Append("   ) x " + vbCrLf)
                    .Append(") z " + vbCrLf)
                    .Append("where  (z.cnt >= 2 and person_id is not null) or z.cnt = 1 " + vbCrLf)
                    .Append("order by x.offering_number, x.korean_name COLLATE " + ControlChars.Quote + "C" + ControlChars.Quote)
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

                If flag = 1 Then
                    If String.IsNullOrEmpty(offeringNumber) = False Then
                        cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(offeringNumber, Integer)
                    End If
                End If
                cmd.Parameters.Add(New NpgsqlParameter("@selectedYear", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(selectedYear, Integer)

                da.SelectCommand = cmd
                da.Fill(getOfferingInfoByPersonIdOrOfferingId, "getOfferingInfoByPersonIdOrOfferingId")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingInfoByPersonIdOrOfferingId : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingInfoByPersonIdOrOfferingId
        End Function

        'This function is to get offering number by type and date
        Public Function getOfferingNumbersByType(ByVal offeringType As String, ByVal offeringDate As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingNumbersByType = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("offeringType=[" + offeringType + "], offeringDate=[" + offeringDate + "]")

            With sb
                .Append("select e.offering_number ")
                .Append("from public.offering_transactions t, public.offering_envelopes e ")
                .Append("where t.offering_envelope_id = e.id ")
                .Append("and t.date = @date ")
                .Append("and t.offering_type = @type or t.other_reason = @type")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@type", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = CType(offeringType, String)
                cmd.Parameters.Add(New NpgsqlParameter("@date", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = CType(offeringDate, String)

                da.SelectCommand = cmd
                da.Fill(getOfferingNumbersByType, "getOfferingNumbersByType")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingNumbersByType : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingNumbersByType
        End Function

        'This function is to create a offering report flexibly
        Public Function getOfferingFlexibleReport(ByVal typeFlag As Boolean, ByVal startDate As String, ByVal endDate As String, ByVal personIds As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingFlexibleReport = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("typeFlag=[" + CType(typeFlag, String) + "], startDate=[" + startDate + "], endDate=[" + endDate + "], persondIds=[" + personIds + "]")

            '0 - 주단위 헌금 내역 (CRA)

            With sb
                .Append("select z.date, z.offer_num, z.korean_name, z.last_name, z.first_name, z.month, " + vbCrLf)
                .Append("       z.offering_type, to_char(z.amount, 'FM9,999,999,999.00') amount " + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("       select a.lvl, a.date, to_char(a.date, 'YYYY Mon') as month,  " + vbCrLf)
                .Append("              to_char(a.date, 'YYYY') as year, a.offering_type,  " + vbCrLf)
                .Append("              a.korean_name, a.last_name, a.first_name, a.offer_num, a.amount " + vbCrLf)
                .Append("       from (  " + vbCrLf)
                .Append("           select t.date,  " + vbCrLf)
                .Append("                  MAX(CASE WHEN t.offering_type = 'weekly' THEN 1  " + vbCrLf)
                .Append("                  WHEN t.offering_type = 'tithe' THEN 2  " + vbCrLf)
                .Append("                  WHEN t.offering_type = 'missions' THEN 3  " + vbCrLf)
                .Append("                  WHEN t.offering_type = 'thanksgiving' THEN 4  " + vbCrLf)
                .Append("                  Else 5 End) lvl,  " + vbCrLf)
                .Append("                  Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END offering_type,  " + vbCrLf)
                .Append("                  p.korean_name as korean_name, p.last_name, p.first_name, e.offering_number as offer_num, " + vbCrLf)
                .Append("                  SUM(t.amount) amount  " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id  " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and t.date between @startDate and @endDate  " + vbCrLf)
                .Append("           and p.id = @personIds" + vbCrLf)
                .Append("           group by p.korean_name, p.last_name, p.first_name, e.offering_number, t.date, " + vbCrLf)
                .Append("                    Case WHEN t.offering_type = 'other' THEN t.other_reason ELSE t.offering_type END  " + vbCrLf)
                .Append("       ) a  " + vbCrLf)
                '.Append("       Union all " + vbCrLf)
                '.Append("       select null as lvl, null as date, to_char(date, 'YYYY Mon') as month,  " + vbCrLf)
                '.Append("              null as year, 'Total', MAX(p.korean_name), null, null, MAX(e.offering_number) as offer_num, " + vbCrLf)
                '.Append("              SUM(amount) amount " + vbCrLf)
                '.Append("       from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                '.Append("       where t.offering_envelope_id = e.id  " + vbCrLf)
                '.Append("       and p.id = e.person_id " + vbCrLf)
                '.Append("       and t.date between @startDate and @endDate " + vbCrLf)
                '.Append("       and p.id = @personIds " + vbCrLf)
                '.Append("       group by to_char(date, 'YYYY Mon')  " + vbCrLf)
                .Append(") z " + vbCrLf)
                .Append("order by z.date, z.lvl, z.offering_type desc " + vbCrLf)
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(startDate, Date)
                cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(endDate, Date)
                cmd.Parameters.Add(New NpgsqlParameter("@personIds", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personIds, Integer)

                da.SelectCommand = cmd
                da.Fill(getOfferingFlexibleReport, "getOfferingFlexibleReport")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingFlexibleReport : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingFlexibleReport
        End Function

        'This function is to create a offering report flexibly
        Public Function getOfferingFlexibleReport1(ByVal typeFlag As Boolean, ByVal startDate As String, ByVal endDate As String, ByVal personIds As String, ByVal allFlag As Boolean, ByVal datesList As List(Of String)) As DataSet
            Dim sql As String = String.Empty
            getOfferingFlexibleReport1 = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("typeFlag=[" + CType(typeFlag, String) + "], startDate=[" + startDate + "], endDate=[" + endDate + "], persondIds=[" + personIds + "], allFlag=[" + CType(allFlag, String) + "]")

            '1 - 주단위 헌금

            If typeFlag = False And allFlag = True Then
                With sb
                    .Append("select date as date, " + vbCrLf)
                    .Append("       e.offering_number as offer_num, p.korean_name as korean_name,  " + vbCrLf)
                    .Append("       to_char(SUM(amount), 'FM9,999,999,999.00') total, " + vbCrLf)
                    .Append("       to_char(SUM(amount) OVER (PARTITION BY date, e.offering_number), 'FM9,999,999,999.00') total_date_offerNumber, " + vbCrLf)
                    .Append("       to_char(SUM(amount) OVER (PARTITION BY date), 'FM9,999,999,999.00') total_date, " + vbCrLf)
                    .Append("       to_char(SUM(amount) OVER (PARTITION BY e.offering_number), 'FM9,999,999,999.00') total_person " + vbCrLf)
                    .Append("from public.offering t, public.offering_envelopes e, public.people p  " + vbCrLf)
                    .Append("where t.offering_envelope_id = e.id   " + vbCrLf)
                    .Append("and p.id = e.person_id  " + vbCrLf)

                    'If String.IsNullOrEmpty(personIds) = False Then
                    '    .Append("and p.id in (" + personIds + ") " + vbCrLf)
                    'End If

                    .Append("and t.date between @startDate and @endDate " + vbCrLf)
                    .Append("group by date, e.offering_number, p.korean_name, t.amount " + vbCrLf)

                    For i As Integer = 0 To datesList.Count - 1
                        .Append("       Union all " + vbCrLf)
                        .Append("       select distinct to_date('" + datesList.Item(i) + "', 'YYYY-MM-DD') as date,  " + vbCrLf)
                        .Append("              e.offering_number as offer_num, p.korean_name as korean_name, " + vbCrLf)
                        .Append("              '0.00' total, '0.00' total_date_offerNumber, " + vbCrLf)
                        .Append("              '0.00' total_date, " + vbCrLf)
                        .Append("              '0.00' total_person " + vbCrLf)
                        .Append("       from public.offering_envelopes e, public.people p " + vbCrLf)
                        .Append("       where p.id = e.person_id " + vbCrLf)
                        .Append("       And e.year in (" + m_Global.getOfferingYears(startDate, endDate) + ") " + vbCrLf)
                        .Append("       And p.id Not in (select distinct p.id " + vbCrLf)
                        .Append("                  from public.offering t, public.offering_envelopes e, public.people p  " + vbCrLf)
                        .Append("                  where t.offering_envelope_id = e.id   " + vbCrLf)
                        .Append("                  And p.id = e.person_id  " + vbCrLf)

                        If String.IsNullOrEmpty(personIds) = False Then
                            .Append("and p.id in (" + personIds + ") " + vbCrLf)
                        End If

                        .Append("                  And t.date = '" + datesList.Item(i) + "')" + vbCrLf)
                        Logger.LogInfo("datesList=[" + datesList.Item(i) + "]")
                    Next
                    .Append("order by 1, 2, 3 " + vbCrLf)
                End With
            Else
                Dim OfferingTypeStr As String = ""
                Dim TypeSumPartitionStr As String = ""

                If typeFlag = True Then
                    OfferingTypeStr = "z.offering_type, z.other_reason, "
                    TypeSumPartitionStr = ", to_char(SUM(z.amount) OVER (PARTITION BY z.offering_type), 'FM9,999,999,999.00') total_type "
                End If

                Dim PersonMainStr As String = ""
                Dim PersonSubStr As String = ""
                Dim PersonGroupbyStr As String = ""
                Dim OfferingNumbersMainStr As String = ", z.offering_numbers "
                Dim PersonDateSumPartitionStr As String = ""
                Dim GrandSumPartitionStr As String = ""
                Dim PersonSumPartitionStr As String = ""

                If String.IsNullOrEmpty(personIds) = False Or allFlag = True Then
                    PersonMainStr = "z.offer_num, z.korean_name, "
                    PersonDateSumPartitionStr = ", to_char(SUM(z.amount) OVER (PARTITION BY z.date, z.offer_num), 'FM9,999,999,999.00') total_date_offerNumber "
                    GrandSumPartitionStr = ", to_char(SUM(z.amount) OVER (PARTITION BY z.date), 'FM9,999,999,999.00') total_date "
                    PersonSumPartitionStr = ", to_char(SUM(z.amount) OVER (PARTITION BY z.offer_num), 'FM9,999,999,999.00') total_person "
                End If

                If String.IsNullOrEmpty(personIds) = False Or allFlag = True Then
                    OfferingNumbersMainStr = ""
                    PersonSubStr = "  e.offering_number as offer_num, p.korean_name as korean_name, "
                    PersonGroupbyStr = " p.korean_name, e.offering_number, "
                Else
                    PersonSubStr = "  MAX(e.offering_number) as offer_num, MAX(p.korean_name) as korean_name, "
                    PersonGroupbyStr = ""
                End If

                With sb
                    .Append("select z.date, " + vbCrLf)
                    .Append(PersonMainStr + vbCrLf)
                    .Append(OfferingTypeStr + vbCrLf)
                    .Append("       to_char(z.amount, 'FM9,999,999,999.00') amount  " + vbCrLf)
                    .Append(TypeSumPartitionStr + vbCrLf)
                    .Append(PersonDateSumPartitionStr + vbCrLf)
                    .Append(OfferingNumbersMainStr + vbCrLf)
                    .Append(GrandSumPartitionStr + vbCrLf)
                    .Append(PersonSumPartitionStr + vbCrLf)
                    .Append("from ( " + vbCrLf)
                    .Append("       select a.lvl, a.date, to_char(a.date, 'YYYY Mon') as month,  " + vbCrLf)
                    .Append("              to_char(a.date, 'YYYY') as year, a.offering_type, a.other_reason,  " + vbCrLf)
                    .Append("              a.offer_num, a.korean_name, a.amount, a.offering_numbers " + vbCrLf)
                    .Append("       from (  " + vbCrLf)
                    .Append("           select t.date,  " + vbCrLf)
                    .Append("                  MAX(CASE WHEN t.offering_type = 'weekly' THEN 1  " + vbCrLf)
                    .Append("                  WHEN t.offering_type = 'tithe' THEN 2  " + vbCrLf)
                    .Append("                  WHEN t.offering_type = 'missions' THEN 3  " + vbCrLf)
                    .Append("                  WHEN t.offering_type = 'thanksgiving' THEN 4  " + vbCrLf)
                    .Append("                  Else 5 End) lvl,  " + vbCrLf)
                    .Append("                  t.offering_type, t.other_reason,  " + vbCrLf)
                    .Append(PersonSubStr + vbCrLf)
                    .Append("                  SUM(t.amount) amount,  string_agg(cast(e.offering_number as varchar), ',' ORDER BY e.offering_number) offering_numbers   " + vbCrLf)
                    .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                    .Append("           where t.offering_envelope_id = e.id  " + vbCrLf)
                    .Append("           and p.id = e.person_id " + vbCrLf)
                    .Append("           and t.date between @startDate and @endDate  " + vbCrLf)
                    If String.IsNullOrEmpty(personIds) = False Then
                        .Append("           and p.id in ( " + personIds + ") " + vbCrLf)
                    End If
                    .Append("           group by " + vbCrLf)
                    .Append(PersonGroupbyStr + vbCrLf)
                    .Append(" t.date, " + vbCrLf)
                    .Append(" t.offering_type, t.other_reason  " + vbCrLf)
                    .Append("       ) a  " + vbCrLf)

                    If allFlag = False Then
                        If String.IsNullOrEmpty(personIds) = True Then
                            .Append("       Union all " + vbCrLf)
                            .Append("       select null as lvl, date as date, null as month,  " + vbCrLf)
                            .Append("              null as year, 'Total', null as other_reason, " + vbCrLf)
                            .Append(PersonSubStr + vbCrLf)
                            .Append("              SUM(amount) amount, NULL offering_numbers " + vbCrLf)
                            .Append("       from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                            .Append("       where t.offering_envelope_id = e.id  " + vbCrLf)
                            .Append("       and p.id = e.person_id " + vbCrLf)
                            .Append("       and t.date between @startDate and @endDate " + vbCrLf)
                            .Append("       group by " + vbCrLf)
                            .Append(PersonGroupbyStr + vbCrLf)
                            .Append(" date " + vbCrLf)
                        Else
                            For i As Integer = 0 To datesList.Count - 1
                                .Append("       Union all " + vbCrLf)
                                .Append("       select null as lvl, to_date('" + datesList.Item(i) + "', 'YYYY-MM-DD') as date, null as month,  " + vbCrLf)
                                .Append("              null as year, '', '', " + vbCrLf)
                                .Append(PersonSubStr + vbCrLf)
                                .Append("              0 amount, NULL offering_numbers " + vbCrLf)
                                .Append("       from public.offering_envelopes e, public.people p " + vbCrLf)
                                .Append("       where p.id = e.person_id " + vbCrLf)
                                .Append("       and p.id in (" + personIds + ") " + vbCrLf)
                                .Append("       and p.id not in (select distinct p.id " + vbCrLf)
                                .Append("                  from public.offering t, public.offering_envelopes e, public.people p  " + vbCrLf)
                                .Append("                  where t.offering_envelope_id = e.id   " + vbCrLf)
                                .Append("                  and p.id = e.person_id  " + vbCrLf)
                                .Append("                  and t.date = '" + datesList.Item(i) + "') " + vbCrLf)
                                .Append("       group by " + vbCrLf)
                                .Append(PersonGroupbyStr + vbCrLf)
                                .Append(" date " + vbCrLf)
                                Logger.LogInfo("datesList=[" + datesList.Item(i) + "]")
                            Next
                        End If
                    Else
                        For i As Integer = 0 To datesList.Count - 1
                            .Append("       Union all " + vbCrLf)
                            .Append("       select null as lvl, to_date('" + datesList.Item(i) + "', 'YYYY-MM-DD') as date, null as month,  " + vbCrLf)
                            .Append("              null as year, '', '', " + vbCrLf)
                            .Append(PersonSubStr + vbCrLf)
                            .Append("              0 amount, NULL offering_numbers " + vbCrLf)
                            .Append("       from public.offering_envelopes e, public.people p " + vbCrLf)
                            .Append("       where p.id = e.person_id " + vbCrLf)

                            If String.IsNullOrEmpty(personIds) = False Then
                                .Append("       and p.id in (" + personIds + ") " + vbCrLf)
                            Else
                                .Append("       and e.year in (" + m_Global.getOfferingYears(startDate, endDate) + ") " + vbCrLf)
                            End If

                            .Append("       and p.id not in (select distinct p.id " + vbCrLf)
                            .Append("                  from public.offering t, public.offering_envelopes e, public.people p  " + vbCrLf)
                            .Append("                  where t.offering_envelope_id = e.id   " + vbCrLf)
                            .Append("                  and p.id = e.person_id  " + vbCrLf)

                            If String.IsNullOrEmpty(personIds) = False Then
                                .Append("                  and p.id in (" + personIds + ") ")
                            End If

                            .Append("                  and t.date = '" + datesList.Item(i) + "' )" + vbCrLf)
                            .Append("       group by " + vbCrLf)
                            .Append(PersonGroupbyStr + vbCrLf)
                            .Append(" date " + vbCrLf)
                            Logger.LogInfo("datesList=[" + datesList.Item(i) + "]")
                        Next
                    End If
                    .Append(") z " + vbCrLf)
                    .Append("order by z.date, " + vbCrLf)
                    .Append(PersonMainStr + vbCrLf)
                    .Append("z.lvl, z.offering_type desc " + vbCrLf)
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

                cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(startDate, Date)
                cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(endDate, Date)

                da.SelectCommand = cmd
                da.Fill(getOfferingFlexibleReport1, "getOfferingFlexibleReport1")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingFlexibleReport1 : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingFlexibleReport1
        End Function

        'This function is to create a offering report flexibly
        Public Function getOfferingFlexibleReport2(ByVal startDate As String, ByVal endDate As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingFlexibleReport2 = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("startDate=[" + startDate + "], endDate=[" + endDate + "]")

            '2 - 주단위 헌금 통계

            With sb
                .Append("select x.date, to_char(x.amount, 'FM9,999,999,999.00') total, " + vbCrLf)
                .Append("       ROUND((100 * (x.amount - lag(x.amount, 1) over (order by x.date)) / lag(x.amount, 1) over (order by x.date))::numeric, 2) || '%' grow_rate, " + vbCrLf)
                .Append("       RANK() OVER (ORDER BY x.amount desc) amount_rank, " + vbCrLf)
                .Append("       x.cnt_single, " + vbCrLf)
                .Append("       x.avg_single, " + vbCrLf)
                .Append("       x.cnt_multi, " + vbCrLf)
                .Append("       x.avg_multi " + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("Select t.date, " + vbCrLf)
                .Append("       SUM(t.amount) amount, " + vbCrLf)
                .Append("       COUNT(distinct e.id) cnt_single, " + vbCrLf)
                .Append("       to_char(SUM(t.amount) / COUNT(distinct e.id)::real, 'FM999.00') avg_single, " + vbCrLf)
                .Append("       COUNT(e.id) cnt_multi, " + vbCrLf)
                .Append("       to_char(SUM(t.amount) / COUNT(e.id)::real, 'FM999.00') avg_multi " + vbCrLf)
                .Append("from public.offering t, public.offering_envelopes e " + vbCrLf)
                .Append("where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("and t.date between @startDate and @endDate " + vbCrLf)
                .Append("group by t.date " + vbCrLf)
                .Append(") x " + vbCrLf)
                .Append("order by x.date ASC" + vbCrLf)
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(startDate, Date)
                cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(endDate, Date)

                da.SelectCommand = cmd
                da.Fill(getOfferingFlexibleReport2, "getOfferingFlexibleReport2")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingFlexibleReport2 : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingFlexibleReport2
        End Function
        'This function is to create a offering report flexibly
        Public Function getOfferingFlexibleReport3(ByVal typeFlag As Boolean, ByVal startDate As String, ByVal endDate As String, ByVal personIds As String, ByVal allFlag As Boolean) As DataSet
            Dim sql As String = String.Empty
            getOfferingFlexibleReport3 = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("typeFlag=[" + CType(typeFlag, String) + "], startDate=[" + startDate + "], endDate=[" + endDate + "], personIds=[" + personIds + "], allFlag=[" + CType(allFlag, String) + "]")

            '3 - 월단위 헌금

            If typeFlag = False And allFlag = True Then
                With sb
                    .Append("select to_char(t.date, 'YYYY-mm') as month, " + vbCrLf)
                    .Append("       e.offering_number as offer_num, p.korean_name as korean_name,  " + vbCrLf)
                    .Append("       to_char(SUM(amount), 'FM9,999,999,999.00') amount " + vbCrLf)
                    '.Append("       to_char(SUM(amount) OVER (PARTITION BY to_char(t.date, 'YYYY-mm'), e.offering_number), 'FM9,999,999,999.00') total_date_offerNumber, " + vbCrLf)
                    '.Append("       to_char(SUM(amount) OVER (PARTITION BY to_char(t.date, 'YYYY-mm')), 'FM9,999,999,999.00') total_date, " + vbCrLf)
                    '.Append("       to_char(SUM(amount) OVER (PARTITION BY e.offering_number), 'FM9,999,999,999.00') total_person " + vbCrLf)
                    .Append("from public.offering t, public.offering_envelopes e, public.people p  " + vbCrLf)
                    .Append("where t.offering_envelope_id = e.id   " + vbCrLf)
                    .Append("and p.id = e.person_id  " + vbCrLf)
                    .Append("and t.date between @startDate and @endDate " + vbCrLf)
                    .Append("group by to_char(t.date, 'YYYY-mm'), e.offering_number, p.korean_name " + vbCrLf)
                    .Append("order by to_char(t.date, 'YYYY-mm'), e.offering_number, p.korean_name " + vbCrLf)
                End With
            Else
                Dim OfferingTypeStr As String = ""
                If typeFlag = True Then
                    OfferingTypeStr = "z.offering_type, z.other_reason, "
                End If

                Dim PersonMainStr As String = ""
                Dim PersonSubStr As String = ""
                Dim PersonGroupbyStr As String = ""

                If String.IsNullOrEmpty(personIds) = False Or allFlag = True Then
                    PersonMainStr = "z.offer_num, z.korean_name, "
                End If

                If String.IsNullOrEmpty(personIds) = False Or allFlag = True Then
                    PersonSubStr = "  e.offering_number as offer_num, p.korean_name as korean_name, "
                    PersonGroupbyStr = " p.korean_name, e.offering_number, "
                Else
                    PersonSubStr = "  MAX(e.offering_number) as offer_num, MAX(p.korean_name) as korean_name, "
                    PersonGroupbyStr = ""
                End If

                With sb
                    .Append("select z.month, " + vbCrLf)
                    .Append(PersonMainStr + vbCrLf)
                    .Append("z.offering_type, z.other_reason, " + vbCrLf)
                    .Append("       to_char(SUM(z.amount), 'FM9,999,999,999.00') amount " + vbCrLf)
                    '.Append("       to_char(SUM(z.amount) OVER (PARTITION BY z.offer_num), 'FM9,999,999,999.00') offerNumber_total " + vbCrLf)
                    .Append("from ( " + vbCrLf)
                    .Append("       select a.lvl, a.date, to_char(a.date, 'YYYY-mm') as month,  " + vbCrLf)
                    .Append("              to_char(a.date, 'YYYY') as year, a.offering_type, a.other_reason,  " + vbCrLf)
                    .Append("              a.offer_num, a.korean_name, a.amount, a.offering_numbers " + vbCrLf)
                    .Append("       from (  " + vbCrLf)
                    .Append("           select t.date,  " + vbCrLf)
                    .Append("                  MAX(CASE WHEN t.offering_type = 'weekly' THEN 1  " + vbCrLf)
                    .Append("                  WHEN t.offering_type = 'tithe' THEN 2  " + vbCrLf)
                    .Append("                  WHEN t.offering_type = 'missions' THEN 3  " + vbCrLf)
                    .Append("                  WHEN t.offering_type = 'thanksgiving' THEN 4  " + vbCrLf)
                    .Append("                  Else 5 End) lvl,  " + vbCrLf)
                    .Append("                  t.offering_type, t.other_reason,  " + vbCrLf)
                    .Append(PersonSubStr + vbCrLf)
                    .Append("                  SUM(t.amount) amount,  string_agg(cast(e.offering_number as varchar), ',' ORDER BY e.offering_number) offering_numbers   " + vbCrLf)
                    .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                    .Append("           where t.offering_envelope_id = e.id  " + vbCrLf)
                    .Append("           and p.id = e.person_id " + vbCrLf)
                    .Append("           and t.date between @startDate and @endDate  " + vbCrLf)
                    If String.IsNullOrEmpty(personIds) = False Then
                        .Append("           and p.id in ( " + personIds + ") " + vbCrLf)
                    End If
                    .Append("           group by " + vbCrLf)
                    .Append(PersonGroupbyStr + vbCrLf)
                    .Append(" t.date, " + vbCrLf)
                    .Append(" t.offering_type, t.other_reason  " + vbCrLf)
                    .Append("       ) a  " + vbCrLf)
                    If allFlag = False Then
                        .Append("       Union all " + vbCrLf)
                        .Append("       select Null as lvl, null as date, to_char(t.date, 'YYYY-mm') as month,  " + vbCrLf)
                        .Append("              null as year, 'Total', null as other_reason,  " + vbCrLf)
                        .Append(PersonSubStr + vbCrLf)
                        .Append("              SUM(amount) amount, NULL offering_numbers " + vbCrLf)
                        .Append("       from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                        .Append("       where t.offering_envelope_id = e.id  " + vbCrLf)
                        .Append("       and p.id = e.person_id " + vbCrLf)
                        .Append("       and t.date between @startDate and @endDate " + vbCrLf)
                        If String.IsNullOrEmpty(personIds) = False Then
                            .Append("           and p.id in (" + personIds + ") " + vbCrLf)
                        End If
                        .Append("       group by " + vbCrLf)
                        .Append(PersonGroupbyStr + vbCrLf)
                        .Append(" to_char(t.date, 'YYYY-mm') " + vbCrLf)
                    End If
                    .Append(") z " + vbCrLf)
                    .Append("group by z.month, ")
                    .Append(PersonMainStr + vbCrLf)
                    .Append("z.offering_type, z.other_reason, z.lvl " + vbCrLf)
                    .Append("order by z.month, " + vbCrLf)
                    .Append(PersonMainStr + vbCrLf)
                    .Append("z.lvl, z.offering_type desc, z.other_reason " + vbCrLf)
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

                cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(startDate, Date)
                cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(endDate, Date)

                da.SelectCommand = cmd
                da.Fill(getOfferingFlexibleReport3, "getOfferingFlexibleReport3")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingFlexibleReport3 : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingFlexibleReport3
        End Function
        'This function is to create a offering report flexibly
        Public Function getOfferingFlexibleReport4(ByVal startDate As String, ByVal endDate As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingFlexibleReport4 = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("startDate=[" + startDate + "], endDate=[" + endDate + "]")

            '4 - 월단위 헌금 통계

            With sb
                .Append("select x.month, to_char(x.amount, 'FM9,999,999,999.00') total, " + vbCrLf)
                .Append("       ROUND((100 * (x.amount - lag(x.amount, 1) over (order by x.month)) / lag(x.amount, 1) over (order by x.month))::numeric, 2) || '%' rate,  " + vbCrLf)
                .Append("       RANK() OVER (ORDER BY x.amount desc) amount_rank, " + vbCrLf)
                .Append("       x.cnt_single,  " + vbCrLf)
                .Append("       x.avg_single,  " + vbCrLf)
                .Append("       x.cnt_multi,  " + vbCrLf)
                .Append("       x.avg_multi " + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("   Select TO_CHAR(t.date, 'yyyy-mm') as month,  " + vbCrLf)
                .Append("          SUM(t.amount) amount, " + vbCrLf)
                .Append("          COUNT(distinct e.id) cnt_single, " + vbCrLf)
                .Append("          to_char(SUM(t.amount) / COUNT(distinct e.id)::real, 'FM999.00') avg_single, " + vbCrLf)
                .Append("          COUNT(e.id) cnt_multi, " + vbCrLf)
                .Append("          to_char(SUM(t.amount) / COUNT(e.id)::real, 'FM999.00') avg_multi " + vbCrLf)
                .Append("from public.offering t, public.offering_envelopes e " + vbCrLf)
                .Append("where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("and t.date between @startDate and @endDate " + vbCrLf)
                .Append("group by TO_CHAR(t.date, 'yyyy-mm') " + vbCrLf)
                .Append(") x " + vbCrLf)
                .Append("order by x.month ASC" + vbCrLf)
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(startDate, Date)
                cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(endDate, Date)

                da.SelectCommand = cmd
                da.Fill(getOfferingFlexibleReport4, "getOfferingFlexibleReport4")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingFlexibleReport4 : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingFlexibleReport4
        End Function
        'This function is to create a offering report flexibly
        Public Function getOfferingFlexibleReport5(ByVal typeFlag As Boolean, ByVal startDate As String, ByVal endDate As String, ByVal personIds As String, ByVal allFlag As Boolean) As DataSet
            Dim sql As String = String.Empty
            getOfferingFlexibleReport5 = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("typeFlag=[" + CType(typeFlag, String) + "], startDate=[" + startDate + "], endDate=[" + endDate + "], personIds=[" + personIds + "], allFlag=[" + CType(allFlag, String) + "]")

            '5 - 년단위 헌금

            If typeFlag = False And allFlag = True Then
                With sb
                    .Append("select to_char(t.date, 'YYYY') as year, " + vbCrLf)
                    .Append("       e.offering_number as offer_num, p.korean_name as korean_name,  " + vbCrLf)
                    .Append("       to_char(SUM(amount), 'FM9,999,999,999.00') amount " + vbCrLf)
                    .Append("from public.offering t, public.offering_envelopes e, public.people p  " + vbCrLf)
                    .Append("where t.offering_envelope_id = e.id   " + vbCrLf)
                    .Append("and p.id = e.person_id  " + vbCrLf)
                    .Append("and to_char(t.date, 'YYYY') between @startDate and @endDate " + vbCrLf)
                    .Append("group by to_char(t.date, 'YYYY'), e.offering_number, p.korean_name " + vbCrLf)
                    .Append("order by to_char(t.date, 'YYYY'), e.offering_number, p.korean_name " + vbCrLf)
                End With
            Else
                Dim OfferingTypeStr As String = ""
                If typeFlag = True Then
                    OfferingTypeStr = "z.offering_type, "
                End If

                Dim PersonMainStr As String = ""
                Dim PersonSubStr As String = ""
                Dim PersonGroupbyStr As String = ""

                If String.IsNullOrEmpty(personIds) = False Or allFlag = True Then
                    PersonMainStr = "z.offer_num, z.korean_name, "
                End If

                If String.IsNullOrEmpty(personIds) = False Or allFlag = True Then
                    PersonSubStr = "  e.offering_number as offer_num, p.korean_name as korean_name, "
                    PersonGroupbyStr = " p.korean_name, e.offering_number, "
                Else
                    PersonSubStr = "  MAX(e.offering_number) as offer_num, MAX(p.korean_name) as korean_name, "
                    PersonGroupbyStr = ""
                End If

                With sb
                    .Append("select z.year, " + vbCrLf)
                    .Append(PersonMainStr + vbCrLf)
                    .Append("z.offering_type, z.other_reason, " + vbCrLf)
                    .Append("       to_char(SUM(z.amount), 'FM9,999,999,999.00') amount " + vbCrLf)
                    .Append("from ( " + vbCrLf)
                    .Append("       select a.lvl, a.date, to_char(a.date, 'YYYY-mm') as month,  " + vbCrLf)
                    .Append("              to_char(a.date, 'YYYY') as year, a.offering_type, a.other_reason,  " + vbCrLf)
                    .Append("              a.offer_num, a.korean_name, a.amount, a.offering_numbers " + vbCrLf)
                    .Append("       from (  " + vbCrLf)
                    .Append("           select t.date,  " + vbCrLf)
                    .Append("                  MAX(CASE WHEN t.offering_type = 'weekly' THEN 1  " + vbCrLf)
                    .Append("                  WHEN t.offering_type = 'tithe' THEN 2  " + vbCrLf)
                    .Append("                  WHEN t.offering_type = 'missions' THEN 3  " + vbCrLf)
                    .Append("                  WHEN t.offering_type = 'thanksgiving' THEN 4  " + vbCrLf)
                    .Append("                  Else 5 End) lvl,  " + vbCrLf)
                    .Append("                  t.offering_type, t.other_reason,  " + vbCrLf)
                    .Append(PersonSubStr + vbCrLf)
                    .Append("                  SUM(t.amount) amount,  string_agg(cast(e.offering_number as varchar), ',' ORDER BY e.offering_number) offering_numbers   " + vbCrLf)
                    .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                    .Append("           where t.offering_envelope_id = e.id  " + vbCrLf)
                    .Append("           and p.id = e.person_id " + vbCrLf)
                    .Append("           and to_char(t.date, 'YYYY') between @startDate and @endDate  " + vbCrLf)
                    If String.IsNullOrEmpty(personIds) = False Then
                        .Append("           and p.id in ( " + personIds + ") " + vbCrLf)
                    End If
                    .Append("           group by " + vbCrLf)
                    .Append(PersonGroupbyStr + vbCrLf)
                    .Append(" t.date, " + vbCrLf)
                    .Append(" t.offering_type, t.other_reason  " + vbCrLf)
                    .Append("       ) a  " + vbCrLf)
                    If allFlag = False Then
                        .Append("       Union all " + vbCrLf)
                        .Append("       select null as lvl, null as date, null as month,  " + vbCrLf)
                        .Append("              to_char(t.date, 'YYYY') as year, 'Total', null,  " + vbCrLf)
                        .Append(PersonSubStr + vbCrLf)
                        .Append("              SUM(amount) amount, NULL offering_numbers " + vbCrLf)
                        .Append("       from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                        .Append("       where t.offering_envelope_id = e.id  " + vbCrLf)
                        .Append("       and p.id = e.person_id " + vbCrLf)
                        .Append("       and to_char(t.date, 'YYYY') between @startDate and @endDate " + vbCrLf)
                        If String.IsNullOrEmpty(personIds) = False Then
                            .Append("           and p.id in (" + personIds + ") " + vbCrLf)
                        End If
                        .Append("       group by " + vbCrLf)
                        .Append(PersonGroupbyStr + vbCrLf)
                        .Append(" to_char(t.date, 'YYYY') " + vbCrLf)
                    End If
                    .Append(") z " + vbCrLf)
                    .Append("group by z.year, ")
                    .Append(PersonMainStr + vbCrLf)
                    .Append("z.offering_type, z.other_reason, z.lvl ")
                    .Append("order by z.year, " + vbCrLf)
                    .Append(PersonMainStr + vbCrLf)
                    .Append("z.lvl, z.offering_type desc " + vbCrLf)
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

                cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = CType(startDate, String)
                cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = CType(endDate, String)
                'If String.IsNullOrEmpty(personIds) = False Then
                '    cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personIds, Integer)
                'End If
                da.SelectCommand = cmd
                da.Fill(getOfferingFlexibleReport5, "getOfferingFlexibleReport5")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingFlexibleReport5 : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingFlexibleReport5
        End Function
        'This function is to create a offering report flexibly
        Public Function getOfferingFlexibleReport6(ByVal typeFlag As Boolean, ByVal startYear As String, ByVal endYear As String, ByVal personIds As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingFlexibleReport6 = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("typeFlag=[" + CType(typeFlag, String) + "], startYear=[" + startYear + "], endYear=[" + endYear + "], personIds=[" + personIds + "]")

            '6 - 년단위 월별

            Dim OfferingTypeStr As String = ""
            If typeFlag = True Then
                OfferingTypeStr = "z.offering_type, z.other_reason, "
            End If

            With sb
                .Append("select z.year, " + vbCrLf)
                .Append(OfferingTypeStr)
                .Append("       to_char(z.jan, 'FM9,999,999,999.00') As January, to_char(z.feb, 'FM9,999,999,999.00') As February, to_char(z.mar, 'FM9,999,999,999.00') As March, to_char(z.apr, 'FM9,999,999,999.00') As April, to_char(z.may, 'FM9,999,999,999.00') As May, " + vbCrLf)
                .Append("       to_char(z.jun, 'FM9,999,999,999.00') As Jun, to_char(z.jul, 'FM9,999,999,999.00') As July, to_char(z.aug, 'FM9,999,999,999.00') As August, to_char(z.sept, 'FM9,999,999,999.00') As September, to_char(z.nov, 'FM9,999,999,999.00') As November, to_char(z.decm, 'FM9,999,999,999.00') As December, z.total " + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("       Select x.year, x.type_lvl, x.offering_type, x.other_reason, SUM(x.Jan) jan, SUM(x.Feb) feb, SUM(x.Mar) mar, SUM(x.Apr) apr, SUM(x.May) may, " + vbCrLf)
                .Append("              SUM(x.Jun) jun, SUM(x.Jul) jul, SUM(x.Aug) aug, SUM(x.Sept) sept, SUM(x.Nov) nov, SUM(x.Dec) decm, " + vbCrLf)
                .Append("              SUM(x.Jan+x.Feb+x.Mar+x.Apr+x.May+x.Jun+x.Jul+x.Aug+x.Sept+x.Nov+x.Dec) total " + vbCrLf)
                .Append("       from ( " + vbCrLf)
                .Append("           Select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl,  " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  SUM(t.amount) as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jan' and e.year between @startYear and @endYear " + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, SUM(t.amount) as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Feb' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason  " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl,  " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, SUM(t.amount) as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Mar' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, SUM(t.amount) as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Apr' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, SUM(t.amount) as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'May' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  SUM(t.amount) as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jun' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, SUM(t.amount) as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jul' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, SUM(t.amount) as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Aug' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, SUM(t.amount) as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Sep' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, SUM(t.amount) as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Oct' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, SUM(t.amount) as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Nov' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, SUM(t.amount) as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Dec' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("       ) x " + vbCrLf)
                .Append("       group by x.year, x.type_lvl, x.offering_type, x.other_reason " + vbCrLf)
                .Append("       union all " + vbCrLf)
                .Append("       select x.year, null, 'Total', null, SUM(x.Jan) jan, SUM(x.Feb) feb, SUM(x.Mar) mar, SUM(x.Apr) apr, SUM(x.May) may, " + vbCrLf)
                .Append("              SUM(x.Jun) jun, SUM(x.Jul) jul, SUM(x.Aug) aug, SUM(x.Sept) sept, SUM(x.Nov) nov, SUM(x.Dec) decm, " + vbCrLf)
                .Append("              SUM(x.Jan+x.Feb+x.Mar+x.Apr+x.May+x.Jun+x.Jul+x.Aug+x.Sept+x.Nov+x.Dec) total " + vbCrLf)
                .Append("       from ( " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl,  " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  SUM(t.amount) as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jan' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, SUM(t.amount) as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Feb' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason  " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl,  " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, SUM(t.amount) as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Mar' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, SUM(t.amount) as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Apr' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, SUM(t.amount) as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'May' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  SUM(t.amount) as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jun' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, SUM(t.amount) as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Jul' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, SUM(t.amount) as Aug, 0 as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Aug' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, SUM(t.amount) as Sept, 0 as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Sep' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, SUM(t.amount) as Oct, 0 as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Oct' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, SUM(t.amount) as Nov, 0 as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Nov' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select '1' lvl, e.year, " + vbCrLf)
                .Append("                  CASE WHEN offering_type = 'weekly' THEN 1 " + vbCrLf)
                .Append("                  WHEN offering_type = 'tithe' THEN 2 " + vbCrLf)
                .Append("                  WHEN offering_type = 'missions' THEN 3 " + vbCrLf)
                .Append("                  WHEN offering_type = 'thanksgiving' THEN 4 " + vbCrLf)
                .Append("                  Else 5 End type_lvl, " + vbCrLf)
                .Append("                  t.offering_type, t.other_reason, " + vbCrLf)
                .Append("                  0 as Jan, 0 as Feb, 0 as Mar, 0 as Apr, 0 as May, " + vbCrLf)
                .Append("                  0 as Jun, 0 as Jul, 0 as Aug, 0 as Sept, 0 as Oct, 0 as Nov, SUM(t.amount) as Dec " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and p.id = e.person_id " + vbCrLf)
                .Append("           and TO_CHAR(t.date, 'Mon') = 'Dec' and e.year between @startYear and @endYear" + vbCrLf)
                .Append("           group by e.year, t.offering_type, t.other_reason " + vbCrLf)
                .Append("       ) x " + vbCrLf)
                .Append("       group by x.year " + vbCrLf)
                .Append(") z " + vbCrLf)
                .Append("order by z.year, z.type_lvl" + vbCrLf)
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@startYear", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(startYear, Integer)
                cmd.Parameters.Add(New NpgsqlParameter("@endYear", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(endYear, Integer)

                da.SelectCommand = cmd
                da.Fill(getOfferingFlexibleReport6, "getOfferingFlexibleReport6")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingFlexibleReport6 : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingFlexibleReport6
        End Function
        'This function is to create a offering report flexibly
        Public Function getOfferingFlexibleReport7(ByVal startDate As String, ByVal endDate As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingFlexibleReport7 = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("startDate=[" + startDate + "], endDate=[" + endDate + "]")

            '7 - 년단위 헌금통계

            With sb
                .Append("select x.year, to_char(x.amount, 'FM9,999,999,999.00') total, " + vbCrLf)
                .Append("       ROUND((100 * (x.amount - lag(x.amount, 1) over (order by x.year)) / lag(x.amount, 1) over (order by x.year))::numeric, 2) || '%' rate, " + vbCrLf)
                .Append("       RANK() OVER (ORDER BY x.amount desc) amount_rank, " + vbCrLf)
                .Append("       x.cnt_single, " + vbCrLf)
                .Append("       x.avg_single, " + vbCrLf)
                .Append("       x.cnt_multi, " + vbCrLf)
                .Append("       x.avg_multi " + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("Select to_char(t.date, 'yyyy') as year, " + vbCrLf)
                .Append("       SUM(t.amount) amount, " + vbCrLf)
                .Append("       COUNT(distinct e.id) cnt_single, " + vbCrLf)
                .Append("       to_char(SUM(t.amount) / COUNT(distinct e.id)::real, 'FM9,999,999,999.00') avg_single, " + vbCrLf)
                .Append("       COUNT(e.id) cnt_multi, " + vbCrLf)
                .Append("       to_char(SUM(t.amount) / COUNT(e.id)::real, 'FM9,999,999,999.00') avg_multi " + vbCrLf)
                .Append("from public.offering t, public.offering_envelopes e " + vbCrLf)
                .Append("where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("and e.year between @startDate and @endDate " + vbCrLf)
                .Append("group by to_char(t.date, 'yyyy') " + vbCrLf)
                .Append(") x " + vbCrLf)
                .Append("order by x.year ASC" + vbCrLf)
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(startDate, Integer)
                cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(endDate, Integer)
                'If String.IsNullOrEmpty(personIds) = False Then
                '    cmd.Parameters.Add(New NpgsqlParameter("@personId", NpgsqlTypes.NpgsqlDbType.Integer)).Value = CType(personIds, Integer)
                'End If
                da.SelectCommand = cmd
                da.Fill(getOfferingFlexibleReport7, "getOfferingFlexibleReport7")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingFlexibleReport7 : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingFlexibleReport7
        End Function

        'This function is to get the list of people whose offering records are not found
        Public Function getPeopleWithoutOffering(ByVal startDate As String, ByVal endDate As String, ByVal personIds As String, ByVal offeringYears As String) As DataTable
            Dim sql As String = String.Empty
            getPeopleWithoutOffering = New DataTable
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("startDate=[" + startDate + "], endDate=[" + endDate + "], personIds=[" + personIds + "], offeringYear=[" + offeringYears + "]")

            Dim offerColumnStr As String = ""
            If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_OFFICE Then
                offerColumnStr = " ,year ,offering_number "
            End If

            With sb
                .Append("select y.person_id, y.korean_name " + offerColumnStr + ", y.cell_phone, y.home_phone, y.work_phone, y.address " + vbCrLf)
                .Append("from ( " + vbCrLf)
                .Append("       Select x.person_id, x.year, x.offering_number, x.korean_name, x.cell_phone, x.home_phone, x.work_phone, INITCAP(x.street) || ' ' || INITCAP(x.city) || ' ' || INITCAP(x.province) || ' ' || x.postal_code as address, SUM(x.amount) amount " + vbCrLf)
                .Append("       from ( " + vbCrLf)
                .Append("           select e.person_id, e.year, e.offering_number, p.korean_name, p.cell_phone, p.home_phone, p.work_phone, a.street, a.city, a.province, a.postal_code, SUM(t.amount) amount, " + vbCrLf)
                .Append("                  (select k.name " + vbCrLf)
                .Append("                   from ( " + vbCrLf)
                .Append("                       select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("                       from public.milestones m where e.person_id = m.person_id and type = 'MembershipMilestone' " + vbCrLf)
                .Append("                   ) k " + vbCrLf)
                .Append("                   where k.row_num = 1 " + vbCrLf)
                .Append("                   ) status " + vbCrLf)
                .Append("           from public.offering t, public.offering_envelopes e, public.people p, public.addresses a " + vbCrLf)
                .Append("           where t.offering_envelope_id = e.id " + vbCrLf)
                .Append("           and e.person_id = p.id " + vbCrLf)
                .Append("           and p.id = a.person_id " + vbCrLf)
                .Append("           and a.address_type = 'home' " + vbCrLf)
                .Append("           and p.deceased_on is null " + vbCrLf)

                If String.IsNullOrEmpty(personIds) = False Then
                    .Append("   and p.id in (" + personIds + ") ")
                End If

                .Append("           and t.date between @startDate and @endDate " + vbCrLf)
                .Append("           group by e.person_id, e.year, e.offering_number, p.korean_name, p.cell_phone, p.home_phone, p.work_phone,  a.street, a.city, a.province, a.postal_code " + vbCrLf)
                .Append("           union all " + vbCrLf)
                .Append("           select e.person_id, e.year, e.offering_number, p.korean_name,  p.cell_phone, p.home_phone, p.work_phone, a.street, a.city, a.province, a.postal_code, 0 amount, " + vbCrLf)
                .Append("                   (select k.name " + vbCrLf)
                .Append("                    from ( " + vbCrLf)
                .Append("                       select name, ROW_NUMBER() OVER (PARTITION BY m.type ORDER BY effective_date desc) row_num " + vbCrLf)
                .Append("                       from public.milestones m where e.person_id = m.person_id and type = 'MembershipMilestone' " + vbCrLf)
                .Append("                   ) k " + vbCrLf)
                .Append("                   where k.row_num = 1 " + vbCrLf)
                .Append("                   ) status " + vbCrLf)
                .Append("           from public.offering_envelopes e, public.people p, public.addresses a " + vbCrLf)
                .Append("           where e.person_id = p.id " + vbCrLf)
                .Append("           and p.id = a.person_id " + vbCrLf)
                .Append("           and a.address_type = 'home' " + vbCrLf)
                .Append("           and p.deceased_on is null " + vbCrLf)
                .Append("           and e.year in (" + offeringYears + ") " + vbCrLf)

                If String.IsNullOrEmpty(personIds) = False Then
                    .Append("       and p.id in (" + personIds + ") ")
                End If

                .Append("   ) x " + vbCrLf)
                .Append("   where x.status = 'church_member' " + vbCrLf)
                .Append("   group by x.person_id, x.year, x.offering_number, x.korean_name, x.cell_phone, x.home_phone, x.work_phone, x.street, x.city, x.province, x.postal_code " + vbCrLf)
                .Append(") y " + vbCrLf)
                .Append("where y.amount = 0 " + vbCrLf)
                .Append("order by y.korean_name COLLATE " + ControlChars.Quote + "C" + ControlChars.Quote)
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(startDate, Date)
                cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(endDate, Date)

                dr = cmd.ExecuteReader

                If dr.HasRows = True Then
                    With getPeopleWithoutOffering
                        .Load(dr)
                        'If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_OFFICE Then
                        '    .Columns(0).ColumnName = "person_id"
                        '    .Columns(1).ColumnName = "한글명"
                        '    .Columns(2).ColumnName = "헌금번호"
                        '    .Columns(3).ColumnName = "셀폰번호"
                        '    .Columns(4).ColumnName = "집 전화번호"
                        '    .Columns(5).ColumnName = "직장 전화번호"
                        '    .Columns(6).ColumnName = "주소"
                        'Else
                        '    .Columns(0).ColumnName = "person_id"
                        '    .Columns(1).ColumnName = "한글명"
                        '    .Columns(2).ColumnName = "셀폰번호"
                        '    .Columns(3).ColumnName = "집 전화번호"
                        '    .Columns(4).ColumnName = "직장 전화번호"
                        '    .Columns(5).ColumnName = "주소"
                        'End If
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
                MessageBox.Show("Error at OfferingDAL.getPeopleWithoutOffering : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
            Return getPeopleWithoutOffering
        End Function

        'This function is to create a offering report flexibly
        Public Function getOfferingDates(ByVal flag As Integer, ByVal startDate As String, ByVal endDate As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingDates = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("flag=[" + CType(flag, String) + "], startDate=[" + startDate + "], endDate=[" + endDate + "]")

            Dim formattedDate As String = "date"
            If flag = 1 Then '월단위 날짜
                formattedDate = "to_char(date, 'YYYY-mm) as date"
            ElseIf flag = 2 Then '년단위 날짜
                formattedDate = "to_char(date, 'YYYY) as date"
            End If

            With sb
                .Append("       select distinct ")
                .Append(formattedDate + vbCrLf)
                .Append("       from public.offering t " + vbCrLf)
                .Append("       where t.date between @startDate and @endDate " + vbCrLf)
                .Append(" order by 1 " + vbCrLf)
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()
                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                cmd.Parameters.Add(New NpgsqlParameter("@startDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(startDate, Date)
                cmd.Parameters.Add(New NpgsqlParameter("@endDate", NpgsqlTypes.NpgsqlDbType.Date)).Value = CType(endDate, Date)

                da.SelectCommand = cmd
                da.Fill(getOfferingDates, "getOfferingDates")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingDates : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingDates
        End Function

        'This function is to create a offering list for details
        Public Function getOfferingForDetails(ByVal personIds As String) As DataSet
            Dim sql As String = String.Empty
            getOfferingForDetails = New DataSet
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim da As NpgsqlDataAdapter = New NpgsqlDataAdapter()
            Dim sb As StringBuilder = New StringBuilder

            Logger.LogInfo("personIds=[" + personIds + "]")

            With sb
                .Append("select p.id, e.year, e.offering_number number, p.korean_name, UPPER(p.last_name) last_name, UPPER(p.first_name) first_name, " + vbCrLf)
                .Append("       INITCAP(a.street) street, INITCAP(a.city) city, INITCAP(a.province) province, UPPER(a.postal_code) postal_code " + vbCrLf)
                .Append("from public.offering_envelopes e, public.people p, public.addresses a " + vbCrLf)
                .Append("where p.id = e.person_id " + vbCrLf)
                .Append("and e.person_id = a.person_id " + vbCrLf)
                .Append("and a.person_id = p.id " + vbCrLf)
                .Append("and p.id in (" + personIds + ") " + vbCrLf)
                .Append("order by e.year desc")
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
                da.Fill(getOfferingForDetails, "getOfferingForDetails")

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at OfferingDAL.getOfferingForDetails : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(da, cn, cmd)
                sb = Nothing
            End Try
            Return getOfferingForDetails
        End Function
    End Class
End Namespace