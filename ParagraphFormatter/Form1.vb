Public Class Form1

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Dim s = Me.TextBox1.Text.Trim()

        Dim paragraphs = New List(Of String)
        Dim sb_paragraph = New Text.StringBuilder

        For Each s_line In s.Split(vbCrLf)
            s_line = s_line.Trim()
            If s_line = "" Then
                ' 空白行
                If sb_paragraph.Length = 0 Then
                    Continue For
                End If
                paragraphs.Add(sb_paragraph.ToString().Trim())
                sb_paragraph.Clear()

            Else
                ' 空白行でない
                If Not (sb_paragraph.Length > 0 AndAlso sb_paragraph.ToString().Last() = "-") Then
                    sb_paragraph.Append(" ")
                End If
                sb_paragraph.Append(s_line)

            End If
        Next
        If sb_paragraph.Length > 0 Then
            paragraphs.Add(sb_paragraph.ToString().Trim())
        End If

        Dim sb = New Text.StringBuilder()
        Dim n = 0
        Dim sb_section = New Text.StringBuilder()

        For Each s_para In paragraphs
            If n + s_para.Length > 4900 Then
                If sb_section.Length > 0 Then
                    If sb.Length > 0 Then
                        sb.AppendLine()
                        sb.AppendLine("########################################")
                        sb.AppendLine()
                    End If
                    sb.AppendLine(sb_section.ToString().Trim())
                    sb_section.Clear()
                    sb_section.AppendLine(s_para)
                    n = s_para.Length
                    Continue For
                End If
            End If

            n += s_para.Length
            sb_section.Append(s_para & vbCrLf & vbCrLf)

        Next
        If sb_section.Length > 0 Then
            If sb.Length > 0 Then
                sb.AppendLine()
                sb.AppendLine("########################################")
                sb.AppendLine()
            End If
            sb.AppendLine(sb_section.ToString().Trim())
        End If
        Me.TextBox2.Text = sb.ToString()

    End Sub


End Class
