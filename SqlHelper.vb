Imports System
Imports System.Data
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Collections.Generic

Public Class SqlHelper
#Region "定数"
    Private Const commStr As String = "connStr" 'DB连接字符串读取
#End Region
    Private Shared ReadOnly connString As String = ConfigurationManager.ConnectionStrings(commStr).ConnectionString
    ''' <summary>
    ''' `SELECT文运行
    ''' </summary>
    ''' <param name="Sql">SQL文</param>
    ''' <param name="type">数据解析方法</param>
    ''' <param name="pars">SQL文参数</param>
    ''' <returns>数据DataTable</returns>
    ''' <remarks></remarks>
    Public Shared Function GetTable(ByVal Sql As String,
                                    type As CommandType,
                                    ParamArray pars As SqlParameter()) As DataTable
        Using conn As SqlConnection = New SqlConnection(connString)
            Using apter As SqlDataAdapter = New SqlDataAdapter(Sql, conn)
                apter.SelectCommand.CommandType = type
                If Not apter Is Nothing Then
                    apter.SelectCommand.Parameters.AddRange(pars)
                End If
                Dim da As DataTable = New DataTable()
                apter.Fill(da)
                Return da
            End Using
        End Using
    End Function
    ''' <summary>
    ''' `DELECT,INSERT文运行
    ''' </summary>
    ''' <param name="Sql">SQL文</param>
    ''' <param name="type">数据解析方法</param>
    ''' <param name="pars">SQL文参数</param>
    ''' <returns>添加或删除行数</returns>
    ''' <remarks></remarks>
    Public Shared Function ExecuteNonquery(ByVal sql As String,
                                            ByVal type As CommandType,
                                            ParamArray pars As SqlParameter()) As Integer
        Using conn As SqlConnection = New SqlConnection(connString)
            Using cmd As SqlCommand = New SqlCommand(sql, conn)
                cmd.CommandType = type
                If Not pars Is Nothing Then
                    cmd.Parameters.Add(pars)
                End If
                conn.Open()
                Return cmd.ExecuteNonQuery()
            End Using
        End Using
    End Function
End Class
